$ErrorActionPreference = 'Stop'

function Write-Utf8NoBom {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content
    )
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Backup-File {
    param([Parameter(Mandatory = $true)][string]$Path)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backup = "$Path.bak.$stamp"
    Copy-Item -LiteralPath $Path -Destination $backup -Force
    return $backup
}

function Ensure-Replaced {
    param(
        [Parameter(Mandatory = $true)][string]$Content,
        [Parameter(Mandatory = $true)][string]$OldValue,
        [Parameter(Mandatory = $true)][string]$NewValue,
        [Parameter(Mandatory = $true)][string]$Label
    )
    if ($Content.Contains($NewValue)) {
        return $Content
    }
    if (-not $Content.Contains($OldValue)) {
        throw "Patch anchor not found: $Label"
    }
    return $Content.Replace($OldValue, $NewValue)
}

function Replace-IfPresent {
    param(
        [Parameter(Mandatory = $true)][string]$Content,
        [Parameter(Mandatory = $true)][string]$OldValue,
        [Parameter(Mandatory = $true)][string]$NewValue
    )
    if ($Content.Contains($NewValue)) {
        return $Content
    }
    if (-not $Content.Contains($OldValue)) {
        return $Content
    }
    return $Content.Replace($OldValue, $NewValue)
}

function Add-Before {
    param(
        [Parameter(Mandatory = $true)][string]$Content,
        [Parameter(Mandatory = $true)][string]$Anchor,
        [Parameter(Mandatory = $true)][string]$InsertText,
        [Parameter(Mandatory = $true)][string]$Marker,
        [Parameter(Mandatory = $true)][string]$Label
    )
    if ($Content.Contains($Marker)) {
        return $Content
    }
    if (-not $Content.Contains($Anchor)) {
        throw "Insert anchor not found: $Label"
    }
    return $Content.Replace($Anchor, "$InsertText$Anchor")
}

function Patch-File {
    param([Parameter(Mandatory = $true)][string]$Path)

    $content = [System.IO.File]::ReadAllText($Path)
    $original = $content

    $decodeMarker = '/* autoclaw-windows-exec-decode */'
    $decodeHelper = @"
$decodeMarker
const WINDOWS_CODEPAGE_ENCODING_MAP = {
	65001: "utf-8",
	54936: "gb18030",
	936: "gbk",
	950: "big5",
	932: "shift_jis",
	949: "euc-kr",
	1252: "windows-1252"
};
let cachedWindowsConsoleEncoding;
function parseWindowsCodePage(raw) {
	if (!raw) return null;
	const match = raw.match(/\b(\d{3,5})\b/);
	if (!match?.[1]) return null;
	const codePage = Number.parseInt(match[1], 10);
	if (!Number.isFinite(codePage) || codePage <= 0) return null;
	return codePage;
}
function resolveWindowsConsoleEncoding() {
	if (process.platform !== "win32") return null;
	if (cachedWindowsConsoleEncoding !== void 0) return cachedWindowsConsoleEncoding;
	try {
		const result = spawnSync("cmd.exe", ["/d", "/s", "/c", "chcp"], {
			windowsHide: true,
			encoding: "utf8",
			stdio: ["ignore", "pipe", "pipe"]
		});
		const codePage = parseWindowsCodePage(`${result.stdout ?? ""}\n${result.stderr ?? ""}`);
		cachedWindowsConsoleEncoding = codePage !== null ? WINDOWS_CODEPAGE_ENCODING_MAP[codePage] ?? null : null;
	} catch {
		cachedWindowsConsoleEncoding = null;
	}
	return cachedWindowsConsoleEncoding;
}
function decodeCapturedOutputBuffer(buffer) {
	const utf8 = buffer.toString("utf8");
	if (process.platform !== "win32") return utf8;
	const encoding = resolveWindowsConsoleEncoding();
	if (!encoding || encoding.toLowerCase() === "utf-8") return utf8;
	try {
		return new TextDecoder(encoding).decode(buffer);
	} catch {
		return utf8;
	}
}

"@

    $textFileMarker = '/* autoclaw-windows-textfile-decode */'
    $textFileHelper = @"
$textFileMarker
function decodeTextFileBuffer(buffer) {
	if (!buffer || buffer.length === 0) return "";
	if (buffer.length >= 2) {
		if (buffer[0] === 255 && buffer[1] === 254) return new TextDecoder("utf-16le").decode(buffer);
		if (buffer[0] === 254 && buffer[1] === 255) {
			const swapped = new Uint8Array(buffer.length);
			for (let i = 0; i < buffer.length; i += 2) {
				swapped[i] = buffer[i + 1] ?? 0;
				swapped[i + 1] = buffer[i] ?? 0;
			}
			return new TextDecoder("utf-16le").decode(swapped);
		}
	}
	if (buffer.length >= 3 && buffer[0] === 239 && buffer[1] === 187 && buffer[2] === 191) {
		return buffer.toString("utf8");
	}
	try {
		return new TextDecoder("utf-8", { fatal: true }).decode(buffer);
	} catch {}
	if (process.platform === "win32") {
		for (const encoding of ["gb18030", "gbk"]) {
			try {
				return new TextDecoder(encoding).decode(buffer);
			} catch {}
		}
	}
	return buffer.toString("utf8");
}

"@

    $content = Add-Before -Content $content -Anchor 'function sanitizeBinaryOutput(' -InsertText $decodeHelper -Marker $decodeMarker -Label "$Path decode helper"
    $content = Replace-IfPresent -Content $content -OldValue 'sanitizeBinaryOutput(data.toString())' -NewValue 'sanitizeBinaryOutput(decodeCapturedOutputBuffer(data))'
    $content = Replace-IfPresent -Content $content -OldValue 'sanitizeBinaryOutput(data2.toString())' -NewValue 'sanitizeBinaryOutput(decodeCapturedOutputBuffer(data2))'

    $content = Add-Before -Content $content -Anchor 'async function readOptionalUtf8File(params) {' -InsertText $textFileHelper -Marker $textFileMarker -Label "$Path text-file helper"
    $content = Replace-IfPresent -Content $content -OldValue ')).toString("utf-8");' -NewValue '));'
    $content = Replace-IfPresent -Content $content -OldValue 'return await fs$1.readFile(params.absolutePath, "utf-8");' -NewValue 'return decodeTextFileBuffer(await fs$1.readFile(params.absolutePath));'
    $content = Replace-IfPresent -Content $content -OldValue 'return await fs.readFile(params.absolutePath, "utf-8");' -NewValue 'return decodeTextFileBuffer(await fs.readFile(params.absolutePath));'
    $content = Replace-IfPresent -Content $content -OldValue 'return (await params.sandbox.bridge.readFile({' -NewValue 'return decodeTextFileBuffer(await params.sandbox.bridge.readFile({'

    if ($content -ne $original) {
        $backup = Backup-File -Path $Path
        Write-Utf8NoBom -Path $Path -Content $content
        [pscustomobject]@{ Path = $Path; Backup = $backup; Changed = $true }
    } else {
        [pscustomobject]@{ Path = $Path; Backup = ''; Changed = $false }
    }
}

$targets = @(
    'C:\Program Files\AutoClaw\resources\gateway\openclaw\dist\compact-1mmJ_KWL.js',
    'C:\Program Files\AutoClaw\resources\gateway\openclaw\gateway-bundle.mjs'
)

$results = foreach ($target in $targets) {
    if (-not (Test-Path $target)) {
        throw "Target not found: $target"
    }
    Patch-File -Path $target
}

$results | Format-Table -AutoSize
