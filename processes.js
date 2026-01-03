const { execFile } = require("child_process");
const { promisify } = require("util");
const iconv = require("iconv-lite");

const execFileAsync = promisify(execFile);
const MAX_PROCESS_BUFFER = 10 * 1024 * 1024;

const ENCODING_CANDIDATES = [
  "utf8",
  "gb18030",
  "big5",
  "shift_jis",
  "euc-kr",
  "windows-1252",
];

const MOJIBAKE_MARKERS = ["\u8107", "\u8117", "\u8292"];

function decodeWithEncoding(bytes, encoding) {
  if (encoding === "utf8") {
    return bytes.toString("utf8");
  }
  return iconv.decode(bytes, encoding);
}

function looksUtf16Le(bytes) {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return true;
  }
  let evenZeros = 0;
  let oddZeros = 0;
  const sampleSize = Math.min(bytes.length, 64);
  for (let i = 0; i < sampleSize; i += 1) {
    if (bytes[i] === 0) {
      if (i % 2 === 0) {
        evenZeros += 1;
      } else {
        oddZeros += 1;
      }
    }
  }
  return oddZeros > evenZeros * 2 && oddZeros > 2;
}

function countBadChars(text) {
  let penalty = 0;
  for (const ch of text) {
    const code = ch.charCodeAt(0);
    if (code === 0xfffd) {
      penalty += 3;
      continue;
    }
    if (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) {
      penalty += 1;
    }
  }
  for (const marker of MOJIBAKE_MARKERS) {
    if (text.includes(marker)) {
      penalty += 1;
    }
  }
  return penalty;
}

function chooseBestDecoded(bytes, original) {
  let bestText = original;
  let bestPenalty = countBadChars(original);
  for (const encoding of ENCODING_CANDIDATES) {
    const decoded = decodeWithEncoding(bytes, encoding);
    const penalty = countBadChars(decoded);
    if (penalty < bestPenalty) {
      bestPenalty = penalty;
      bestText = decoded;
    }
  }
  return bestText;
}

function autoDecodeText(value) {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (Buffer.isBuffer(value)) {
    if (looksUtf16Le(value)) {
      return decodeWithEncoding(value, "utf16le").replace(/^\uFEFF/, "");
    }
    return chooseBestDecoded(value, decodeWithEncoding(value, "utf8"));
  }
  if (typeof value !== "string") {
    return String(value);
  }
  if (!value) {
    return value;
  }
  const bytes = Buffer.from(value, "latin1");
  return chooseBestDecoded(bytes, value);
}

const parseTasklistNames = (stdout) => {
  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      if (line.startsWith('"')) {
        const trimmed = line.length >= 2 ? line.slice(1, -1) : line;
        const [name] = trimmed.split('","');
        return name;
      }
      return line.split(/\s+/)[0];
    })
    .filter(Boolean);
};

const parseTasklistFirstName = (stdout) => {
  const names = parseTasklistNames(stdout);
  return names[0] ?? "";
};

const listWindowsNames = async () => {
  const { stdout } = await execFileAsync("tasklist", ["/FO", "CSV", "/NH"], {
    encoding: "buffer",
    windowsHide: true,
    maxBuffer: MAX_PROCESS_BUFFER,
  });
  const decoded = autoDecodeText(stdout);
  if (!decoded) {
    return [];
  }
  return parseTasklistNames(decoded);
};

const parseWmicProcessPaths = (stdout) => {
  const map = new Map();
  let name = "";
  let path = "";
  const flush = () => {
    if (name && path && !map.has(name)) {
      map.set(name, path);
    }
    name = "";
    path = "";
  };
  stdout.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      flush();
      return;
    }
    const [keyRaw, ...rest] = trimmed.split("=");
    if (!keyRaw) {
      return;
    }
    const key = keyRaw.trim().toLowerCase();
    const value = rest.join("=").trim();
    if (key === "name") {
      name = value;
    } else if (key === "executablepath") {
      path = value;
    }
  });
  flush();
  return map;
};

const parsePowerShellProcessPaths = (stdout) => {
  const map = new Map();
  const trimmed = stdout.trim();
  if (!trimmed) {
    return map;
  }
  let data;
  try {
    data = JSON.parse(trimmed);
  } catch {
    return map;
  }
  const items = Array.isArray(data) ? data : [data];
  items.forEach((item) => {
    if (!item || typeof item !== "object") {
      return;
    }
    const name = typeof item.Name === "string" ? item.Name : "";
    const filePath =
      typeof item.ExecutablePath === "string"
        ? item.ExecutablePath
        : typeof item.Path === "string"
          ? item.Path
          : "";
    if (name && filePath && !map.has(name)) {
      map.set(name, filePath);
    }
  });
  return map;
};

const parseWmicValue = (stdout, keyName) => {
  let result = "";
  stdout.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) {
      return;
    }
    const [keyRaw, ...rest] = trimmed.split("=");
    if (!keyRaw) {
      return;
    }
    const key = keyRaw.trim().toLowerCase();
    if (key === keyName.toLowerCase()) {
      result = rest.join("=").trim();
    }
  });
  return result;
};

const listWindowsProcessPaths = async () => {
  try {
    const { stdout } = await execFileAsync(
      "wmic",
      ["process", "get", "Name,ExecutablePath", "/VALUE"],
      {
        encoding: "buffer",
        windowsHide: true,
        maxBuffer: MAX_PROCESS_BUFFER,
      }
    );
    const decoded = autoDecodeText(stdout) ?? "";
    if (decoded) {
      const map = parseWmicProcessPaths(decoded);
      if (map.size) {
        return map;
      }
    }
  } catch {
    // Fall back to PowerShell when WMIC is missing or blocked.
  }
  try {
    const { stdout } = await execFileAsync(
      "powershell",
      [
        "-NoProfile",
        "-Command",
        "Get-CimInstance Win32_Process | Select-Object -Property Name,ExecutablePath | ConvertTo-Json -Compress -Depth 2",
      ],
      {
        encoding: "utf8",
        windowsHide: true,
        maxBuffer: MAX_PROCESS_BUFFER,
      }
    );
    return parsePowerShellProcessPaths(stdout);
  } catch {
    return new Map();
  }
};

const getWindowsProcessNameByPid = async (pid) => {
  if (!Number.isFinite(pid) || pid <= 0) {
    return "";
  }
  try {
    const { stdout } = await execFileAsync(
      "wmic",
      ["process", "where", `processid=${pid}`, "get", "Name", "/VALUE"],
      {
        encoding: "buffer",
        windowsHide: true,
        maxBuffer: MAX_PROCESS_BUFFER,
      }
    );
    const decoded = autoDecodeText(stdout) ?? "";
    if (decoded) {
      const name = parseWmicValue(decoded, "Name");
      if (name) {
        return name;
      }
    }
  } catch {
    // Fall back to tasklist.
  }
  try {
    const { stdout } = await execFileAsync(
      "tasklist",
      ["/FO", "CSV", "/NH", "/FI", `PID eq ${pid}`],
      {
        encoding: "buffer",
        windowsHide: true,
        maxBuffer: MAX_PROCESS_BUFFER,
      }
    );
    const decoded = autoDecodeText(stdout) ?? "";
    if (!decoded) {
      return "";
    }
    return parseTasklistFirstName(decoded);
  } catch {
    return "";
  }
};

const listUnixNames = async () => {
  const { stdout } = await execFileAsync("ps", ["-A", "-o", "comm="], {
    encoding: "utf8",
  });
  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
};

const listDarwinNames = async () => {
  return listUnixNames();
};

const listRunningProcessNames = async () => {
  try {
    let names;
    if (process.platform === "win32") {
      names = await listWindowsNames();
    } else if (process.platform === "darwin") {
      names = await listDarwinNames();
    } else {
      names = await listUnixNames();
    }
    return [...new Set(names.map((name) => name.trim()).filter(Boolean))];
  } catch {
    return [];
  }
};

module.exports = {
  listRunningProcessNames,
  listWindowsProcessPaths,
  getWindowsProcessNameByPid,
};
