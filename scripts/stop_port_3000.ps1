# Stops processes listening on TCP port 3000
try {
    $p = (netstat -ano | Select-String ':3000' | ForEach-Object { $_.ToString().Trim().Split()[-1] }) -join ','
    if ($p -and $p -ne '') {
        $p.Split(',') | ForEach-Object {
            if ([int]::TryParse($_, [ref]$null)) {
                Stop-Process -Id ([int]$_) -Force -ErrorAction SilentlyContinue
                Write-Output "Stopped PID $_"
            }
        }
    } else {
        Write-Output 'No process on port 3000'
    }
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
}