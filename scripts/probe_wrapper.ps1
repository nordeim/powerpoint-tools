# probe_wrapper.ps1
# Usage: .\probe_wrapper.ps1 -File C:\path\to\presentation.pptx
param(
  [Parameter(Mandatory=$true)][string]$File,
  [int]$TimeoutSeconds = 15,
  [int]$MaxRetries = 3
)

function Emit-ErrorJson {
  param($Code, $Message, $Retryable)
  $obj = @{ error = @{ error_code = $Code; message = $Message; retryable = $Retryable } }
  $obj | ConvertTo-Json -Depth 5
}

if (-not [System.IO.Path]::IsPathRooted($File)) {
  Emit-ErrorJson "RELATIVE_PATH_NOT_ALLOWED" "Absolute path required" $false
  exit 1
}

if (-not (Test-Path -Path $File -PathType Leaf)) {
  Emit-ErrorJson "FILE_NOT_FOUND" "File does not exist" $false
  exit 1
}

if (-not (Get-Item $File).IsReadOnly -and -not (Get-Acl $File)) {
  # best-effort permission check; continue
  $null = $null
}

# Disk space check
$drive = Get-PSDrive -Name ([System.IO.Path]::GetPathRoot($File).TrimEnd('\'))
if ($drive.Free -lt 100MB) {
  Emit-ErrorJson "LOW_DISK_SPACE" "Available space less than 100MB" $false
  exit 1
}

# Tool availability
if (-not (Get-Command ppt_capability_probe.py -ErrorAction SilentlyContinue)) {
  Emit-ErrorJson "TOOL_MISSING" "ppt_capability_probe.py not found" $false
  exit 1
}

$attempt = 0
while ($attempt -lt $MaxRetries) {
  $attempt++
  try {
    $proc = Start-Process -FilePath "ppt_capability_probe.py" -ArgumentList "--file", $File, "--deep", "--json" -NoNewWindow -RedirectStandardOutput "$env:TEMP\probe_out.json" -Wait -PassThru -ErrorAction Stop -Timeout $TimeoutSeconds
    Get-Content "$env:TEMP\probe_out.json" | Out-String
    Remove-Item "$env:TEMP\probe_out.json" -ErrorAction SilentlyContinue
    exit 0
  } catch {
    Start-Sleep -Seconds ([math]::Pow(2, $attempt))
  }
}

# Fallback probes
if (Get-Command ppt_get_info.py -ErrorAction SilentlyContinue -and Get-Command ppt_get_slide_info.py -ErrorAction SilentlyContinue) {
  try {
    & ppt_get_info.py --file $File --json > "$env:TEMP\info.json"
    & ppt_get_slide_info.py --file $File --slide 0 --json > "$env:TEMP\slide0.json"
    $info = Get-Content "$env:TEMP\info.json" -Raw | ConvertFrom-Json
    $slide0 = Get-Content "$env:TEMP\slide0.json" -Raw | ConvertFrom-Json
    $merged = $info | Add-Member -PassThru -NotePropertyName probe_fallback -NotePropertyValue $true
    $merged | Add-Member -PassThru -NotePropertyName slide0 -NotePropertyValue $slide0
    $merged | ConvertTo-Json -Depth 10
    Remove-Item "$env:TEMP\info.json","$env:TEMP\slide0.json" -ErrorAction SilentlyContinue
    exit 0
  } catch {
    Emit-ErrorJson "PROBE_FALLBACK_FAILED" "Both deep probe and fallback probes failed" $true
    exit 3
  }
} else {
  Emit-ErrorJson "PROBE_AND_FALLBACK_TOOLS_MISSING" "Fallback tools not available" $false
  exit 1
}
