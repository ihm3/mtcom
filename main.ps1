param (
    [string]$port = "com3",
    [int]$baudrate = 9600,
    [string]$parity = "none",
    [int]$databits = 8,
    [string]$stopbits = "one",
    [switch]$s,
    [string]$capture,
    [string]$script,
    [switch]$h
)

if ($h) {
    Write-Host @"
mtcom - serial terminal tool for windows

usage:
  .\main.ps1 [options]

options:
  -port <name>      serial port (default: com3)
  -baudrate <n>         baud rate (default: 9600)
  -parity <none|even|odd|mark|space> (default: none)
  -databits <5-8>       data bits (default: 8)
  -stopbits <one|two|onepointfive> (default: one)
  -s                    setup mode (configuration menu)
  -capture <file>       log all input/output to file
  -script <path>        run startup script
  -h                    show this help

runtime commands:
  ctrl+a then:
    a   toggle add lf after cr
    b   view scrollback buffer
    c   clear screen
    e   toggle local echo
    f   send break signal
    l   toggle capture/log
    r   receive file (xmodem)
    s   send file (xmodem)
    w   toggle line wrap
    x   exit
    z   show help

  f1-f10   macros (defined in config/macros.json)
"@
    return
}

Import-Module .\mtcom.psm1 -Force

if ($s) {
    show-configmenu
    return
}

if ($capture) { $global:capturefile = $capture }

start-mtcom -port $port -baudrate $baudrate -parity $parity -databits $databits -stopbits $stopbits
