# preflight_check.ps1
# Usage: .\preflight_check.ps1 -File C:\path\to\presentation.pptx
param(
  [Parameter(Mandatory=$true)][string]$File
)

$MinSpaceMB = 100

if (-not [System.IO.Path]::IsPathRooted($File)) {
  Write-Output '{"error":"Absolute path required"}'
  exit 1
}

if (-not (Test-Path -Path $File -PathType Leaf)) {
  Write-Output '{"error":"File not readable"}'
  exit 1
}

$dir = [System.IO.Path]::GetDirectoryName($File)
# Simple write check by creating temp file
try {
  $testFile = Join-Path $dir ([System.Guid]::NewGuid().ToString())
  New-Item -Path $testFile -ItemType File -Force | Out-Null
  Remove-Item $testFile -Force
} catch {
  Write-Output '{"error":"No write permission to destination directory"}'
  exit 1
}

$drive = Get-PSDrive -Name ([System.IO.Path]::GetPathRoot($File).TrimEnd('\'))
if ($drive.Free -lt ($MinSpaceMB * 1MB)) {
  Write-Output ('{"error":"Low disk space: ' + [math]::Round($drive.Free / 1MB) + 'MB available"}')
  exit 1
}

# Run probe wrapper
try {
  $probeJson = & .\probe_wrapper.ps1 -File $File
  if ($LASTEXITCODE -eq 0) {
    $probeObj = $probeJson | ConvertFrom-Json
    $result = @{
      file = $File
      preflight = @{ status = "ok" }
      probe = $probeObj
    }
    $result | ConvertTo-Json -Depth 10
    exit 0
  } else {
    Write-Output '{"error":"Probe failed"}'
    exit 2
  }
} catch {
  Write-Output '{"error":"Probe failed"}'
  exit 2
}
