$global:port = $null
$global:config = @{}
$global:scrollback = [System.Collections.ArrayList]::new()
$global:commandmode = $false
$global:localecho = $false
$global:addlf = $false
$global:captureactive = $false
$global:capturefile = "logs/mtcom_$(Get-Date -Format yyyyMMdd_HHmmss).log"

function load-config {
    $default = @{
        baudrate = 9600
        parity = "none"
        databits = 8
        stopbits = "one"
        historysize = 2000
        downloaddir = "downloads"
        uploaddir = "uploads"
        statusline = $true
    }
    if (Test-Path "config\mtcom.dfl.json") {
        $global:config = Get-Content "config\mtcom.dfl.json" | ConvertFrom-Json
    } else {
        $global:config = $default
        $global:config | ConvertTo-Json -Depth 10 | Set-Content "config\mtcom.dfl.json"
    }
}

function save-config {
    $global:config | ConvertTo-Json -Depth 10 | Set-Content "config\mtcom.dfl.json"
}

function load-macros {
    if (Test-Path "config\macros.json") {
        $global:macros = Get-Content "config\macros.json" | ConvertFrom-Json -AsHashtable
    } else {
        $global:macros = @{}
    }
}

function show-statusline {
    $status = "mtcom | $($global:config.port) | $($global:config.baudrate) $($global:config.parity)$($global:config.databits)$($global:config.stopbits) | capture: $($global:captureactive ? 'on' : 'off')"
    Write-Host $status -ForegroundColor Yellow
}

function send-break {
    $global:port.BreakState = $true
    Start-Sleep -Milliseconds 400
    $global:port.BreakState = $false
}

function toggle-capture {
    $global:captureactive = -not $global:captureactive
    if ($global:captureactive) { New-Item -ItemType Directory -Force -Path "logs" | Out-Null }
    Write-Host "capture: $($global:captureactive ? 'on' : 'off') -> $global:capturefile"
}

function send-files {
    $files = Get-ChildItem $global:config.uploaddir -File
    if ($files.Count -eq 0) { Write-Host "no files in upload folder"; return }
    $choice = $files | Out-GridView -Title "select file to send" -PassThru
    if ($choice) {
        . .\protocols\xmodem.ps1
        send-xmodem $choice.FullName
    }
}

function receive-files {
    . .\protocols\xmodem.ps1
    receive-xmodem $global:config.downloaddir
}

function show-scrollback {
    Clear-Host
    $global:scrollback | ForEach-Object { Write-Host $_ -NoNewline }
    Write-Host "`n--- scrollback (q to quit) ---"
    while ($true) {
        if ($host.UI.RawUI.KeyAvailable) {
            $k = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ($k.Character -eq 'q') { Clear-Host; break }
        }
        Start-Sleep -Milliseconds 50
    }
}

function show-help {
    Clear-Host
    Write-Host @"
mtcom help (ctrl+a z)

commands (ctrl+a + lowercase):
  a   toggle add lf after cr
  b   show scrollback buffer
  c   clear screen
  e   toggle local echo
  f   send break
  l   toggle capture/log
  r   receive file (xmodem)
  s   send file (xmodem)
  w   toggle line wrap
  x   exit
  z   show this help

f1-f10   macros (config/macros.json)
"@
    Read-Host "press enter to continue"
    Clear-Host
}

function handle-command {
    param([string]$cmd)
    switch ($cmd) {
        "a" { $global:addlf = -not $global:addlf; Write-Host "add lf: $($global:addlf)" }
        "b" { show-scrollback }
        "c" { Clear-Host }
        "e" { $global:localecho = -not $global:localecho; Write-Host "local echo: $($global:localecho)" }
        "f" { send-break }
        "l" { toggle-capture }
        "r" { receive-files }
        "s" { send-files }
        "w" { $global:linewrap = -not $global:linewrap; Write-Host "line wrap: $($global:linewrap)" }
        "x" { throw "exit" }
        "z" { show-help }
    }
}

function start-mtcom {
    param($port, $baudrate, $parity, $databits, $stopbits)

    load-config
    load-macros

    $global:config.port = $port
    $global:config.baudrate = $baudrate

    $global:port = New-Object System.IO.Ports.SerialPort $port,$baudrate,[System.IO.Ports.Parity]$parity,$databits,[System.IO.Ports.StopBits]$stopbits
    $global:port.DtrEnable = $true
    $global:port.RtsEnable = $true

    try {
        $global:port.Open()
        Write-Host "mtcom started: $port @ $baudrate" -ForegroundColor Green
        if ($global:config.statusline) { show-statusline }

        Register-ObjectEvent -InputObject $global:port -EventName DataReceived -Action {
            $data = $global:port.ReadExisting()
            Write-Host $data -NoNewline -ForegroundColor Cyan
            if ($global:localecho) { Write-Host $data -NoNewline }
            $global:scrollback.Add($data) | Out-Null
            while ($global:scrollback.Count -gt $global:config.historysize) { $global:scrollback.RemoveAt(0) }
            if ($global:captureactive) { Add-Content $global:capturefile $data -NoNewline }
        } | Out-Null

        while ($true) {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                $char = $key.KeyChar

                if ($key.Modifiers -band [ConsoleModifiers]::Control -and $char -eq [char]1) {
                    $global:commandmode = $true
                    continue
                }

                if ($global:commandmode) {
                    $global:commandmode = $false
                    handle-command ($char.ToString().ToLower())
                    continue
                }

                if ($key.Key -ge [ConsoleKey]::F1 -and $key.Key -le [ConsoleKey]::F10 -and $global:macros.Count -gt 0) {
                    $fnum = [int]$key.Key - 111
                    $macro = $global:macros."f$fnum"
                    if ($macro) { $global:port.Write($macro) }
                    continue
                }

                $send = $char
                if ($global:addlf -and $char -eq "`r") { $send += "`n" }
                $global:port.Write($send)
                if ($global:localecho) { Write-Host $char -NoNewline }
                if ($global:captureactive) { Add-Content $global:capturefile $char -NoNewline }
            }
            Start-Sleep -Milliseconds 10
        }
    }
    catch {
        if ($_.Exception.Message -eq "exit") {
            Write-Host "exiting..." -ForegroundColor Yellow
        } else {
            Write-Host "error: $_" -ForegroundColor Red
        }
    }
    finally {
        if ($global:port.IsOpen) { $global:port.Close() }
        Write-Host "mtcom terminated." -ForegroundColor Green
    }
}

Export-ModuleMember -Function start-mtcom
