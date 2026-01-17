$qBittorrentURL  = "http://localhost:8080"
$qBittorrentUser = "admin"
$PasswordFilePath = "$env:LOCALAPPDATA\qBittorrentPassword.txt"
$LogPath          = "$env:LOCALAPPDATA\Proton\Proton VPN\Logs\"


function Get-StoredPassword {
    if (Test-Path $PasswordFilePath) {
        try { return (Get-Content -Path $PasswordFilePath | ConvertTo-SecureString) }
        catch { Write-Error "Failed to decrypt password."; return $null }
    }
    Write-Error "Password file not found. Run the setup command first."
    return $null
}

function Get-ProtonPortFromLogs {
    # Find the most recent Log File
    if (-not (Test-Path $LogPath)) { Write-Error "Log path not found: $LogPath"; return $null }
    
    $latestLog = Get-ChildItem -Path $LogPath -Recurse -Filter "*.txt" | 
                 Sort-Object LastWriteTime -Descending | 
                 Select-Object -First 1

    if (-not $latestLog) { Write-Error "No log files found."; return $null }

    # Read the last 200 lines
    # Looking for something like: "Port pair 12345->12345"
    $content = Get-Content -Path $latestLog.FullName -Tail 200
    
    # Find pattern
    $matches = $content | Select-String -Pattern "Port pair (\d+)" -AllMatches
    
    if ($matches) {
        # Most recent match
        $lastMatch = $matches[-1].Matches.Groups[1].Value
        return [int]$lastMatch
    }
    
    return $null
}

function Get-QbitSession {
    param ($Url, $User, $Password)
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
    $PlainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    
    try {
        $body = @{ username = $User; password = $PlainPass }
        Invoke-RestMethod -Uri "$Url/api/v2/auth/login" -Method Post -WebSession $session -Body $body -ErrorAction Stop | Out-Null
        return $session
    }
    catch { Write-Error "qBittorrent Login Failed."; return $null }
}

function Get-QbitPort {
    param ($Url, $Session)
    try {
        $pref = Invoke-RestMethod -Uri "$Url/api/v2/app/preferences" -Method Get -WebSession $Session -ErrorAction Stop
        return $pref.listen_port
    }
    catch { return $null }
}

function Set-QbitPort {
    param ($Url, $Session, [int]$Port)
    $json = @{ listen_port = $Port } | ConvertTo-Json -Compress
    try {
        Invoke-RestMethod -Uri "$Url/api/v2/app/setPreferences" -Method Post -WebSession $Session -Body @{ json = $json } -ErrorAction Stop | Out-Null
        Write-Host "SUCCESS: qBittorrent port updated to $Port" -ForegroundColor Green
        return $true
    }
    catch { Write-Error "Failed to set port."
            return $false 
    }
}


# Check if qBittorrent is running
$qProcess = Get-Process -Name 'qbittorrent' -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $qProcess) {
    Write-Host "qBittorrent is not running. Exiting." -ForegroundColor Yellow
    exit
}

# Get Credentials
$securePass = Get-StoredPassword
if (-not $securePass) { exit }

# Get Port from ProtonVPN Logs
$protonPort = Get-ProtonPortFromLogs
if (-not $protonPort) {
    Write-Host "Could not find 'Port pair' in the latest ProtonVPN log." -ForegroundColor Red
    exit
}

# Authenticate qBittorrent
$qbitSession = Get-QbitSession -Url $qBittorrentURL -User $qBittorrentUser -Password $securePass
if (-not $qbitSession) { exit }

# Check and Update
$currentPort = Get-QbitPort -Url $qBittorrentURL -Session $qbitSession

if ($currentPort -eq $protonPort) {
    Write-Host "Ports match ($protonPort). No action needed." -ForegroundColor Cyan
}
else {
    Write-Host "Mismatch! Proton: $protonPort | qBit: $currentPort" -ForegroundColor Yellow
    
    # Update qBittorrent port
    $updateSuccess = Set-QbitPort -Url $qBittorrentURL -Session $qbitSession -Port $protonPort
    
    if ($updateSuccess) {
        # Restart app (fixes torrents getting stalled bug)
        Write-Host "Restarting qBittorrent to apply changes..." -ForegroundColor Magenta
        
        # Capture the executable path before we kill the process
        $exePath = $qProcess.Path
        
        # Stop process
        Stop-Process -Id $qProcess.Id -Force
        
        # Wait 5 seconds for file locks to release
        Start-Sleep -Seconds 5
        
        # Start the process again
        if ($exePath -and (Test-Path $exePath)) {
            Start-Process -FilePath $exePath
            Write-Host "qBittorrent restarted successfully." -ForegroundColor Green
        }
        else {
            Write-Error "Could not find qBittorrent executable at '$exePath' to restart it. Please start it manually."
        }
    }
}
