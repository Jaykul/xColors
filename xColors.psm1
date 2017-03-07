Add-MetadataConverter @{
    [Tuple[ConsoleColor,Drawing.Color]] = { "ConsoleColorColor {0} '{1}'" -f $_.Item1, ("{0},{1},{2}" -f $_.Item2.R, $_.Item2.G, $_.Item2.B) }
    ConsoleColorColor = { param([ConsoleColor]$name, [Drawing.Color]$color) [Tuple[ConsoleColor,Drawing.Color]]::new($name, $color) }
}

# Install-Package ObscureWare.Console.Core -Destination .\lib\ -SkipDependencies
# Add-Type -Path .\lib\Colourful*\lib\net46\Colourful.dll
$Controller = [ObscureWare.Console.ConsoleController]::new()

# To keep it backward compatible, don't use PS classes:
# enum xColor { Black, DarkRed, DarkGreen, DarkYellow, DarkBlue, DarkMagenta, DarkCyan, Gray, DarkGray, Red, Green, Yellow, Blue, Magenta, Cyan, White }
Add-Type -TypeDefinition "public enum xColor { Black, DarkRed, DarkGreen, DarkYellow, DarkBlue, DarkMagenta, DarkCyan, Gray, DarkGray, Red, Green, Yellow, Blue, Magenta, Cyan, White }"

function Set-xColorTheme {
    <#
        .Synopsis
            Set the color theme of the Windows console from a theme on xcolors.net
        .Description
            Downloads the theme from xcolors.net and parses it, then saves it
    #>
    [CmdletBinding()]
    param(
        # The name of a theme from xcolors.net
        [string]$themeName = "grandshell"
    )

    try {
        $ErrorActionPreference = "Stop"
        $ThemeLines = (Invoke-RestMethod "http://www.xcolors.net/dl/$themeName") -split "[\r\n]+"
    } catch {
        Write-Error "Unable to fetch the theme $themeName from xcolors.net"
    }
    $ErrorActionPreference = "Continue"
    $ThemeLines = ($ThemeLines | Where {$_ -notmatch "^!" -and $_ -match "color\d+" }) -replace "\s+|\*\.?" -replace '.*(color[\d]+)\:','$1:'

    Write-Verbose ($ThemeLines -join "`n")

    $Theme = @{}
    $ThemeLines.ForEach{

            $index, $value = $_.split(@("color",":"),2,"RemoveEmptyEntries")
            if($value -notmatch "#[abcdef0123456789]{6}") {
                $value = $value -replace "(?:.*([abcdef0123456789]{2}).*)(?:.*([abcdef0123456789]{2}).*)(?:.*([abcdef0123456789]{2}).*)",'#$1$2$3'
            }
            $Theme.([ConsoleColor][string][xColor]$index) = [Drawing.Color]$value

    }

    Write-Verbose (($Theme | Format-Table | Out-String -stream) -join "`n")

    Set-Theme $Theme -ErrorVariable E
    Write-Verbose "Theme set to $themeName with $($E.Count) errors"
    if($E.Count -eq 0) {
        Write-Verbose "Exporting Theme!"
        Export-Configuration @{Theme = $Theme}
    }
}

function Show-Color {
    [CmdletBinding()]
    param()
    Write-Host "                      40m    41m    42m    43m    44m    45m    46m    47m    100m   107m"
    0..15 | % {
    $Color = [String]([ConsoleColor]$_)
        Write-Host " Hello " -NoNewline -Fore "White" -Back $Color
        Write-Host " $Color ".PadRight(13) -Fore "Black" -Back $Color -NoNewline
        Write-Host "  txt  " -Back Black -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkRed -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkGreen -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkYellow -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkBlue -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkMagenta -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkCyan -Fore $Color -NoNewline
        Write-Host "  txt  " -Back DarkGray -Fore $Color -NoNewline
        Write-Host "  txt  " -Back Gray -Fore $Color -NoNewline
        Write-Host "  txt  " -Back White -Fore $Color
    }
}

function Set-Theme {
    [CmdletBinding()]
    param([Tuple[ConsoleColor,Drawing.Color][]]$Theme)
    try {
        $Controller.ReplaceConsoleColors($theme)
    } catch {
        Write-Error $_
    }
}

function Get-Theme {
    [CmdletBinding()]
    param()
    $Theme = Import-Configuration
    if($Theme) {
        Write-Verbose "Imported Theme $($Theme.Theme | ft -aut | out-string)"
        Set-Theme $Theme.Theme
    }
}

Get-Theme -ErrorAction SilentlyContinue