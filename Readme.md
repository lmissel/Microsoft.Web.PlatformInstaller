# Microsoft.Web.PlatformInstaller

Dieses PowerShell Module ermöglicht es Applikationen, Features und Tools aus den PlatformInstaller über die PowerShell zu installieren.

## Voraussetzungen

Stelle sicher, dass Du den Microsoft Web Platform Installer installiert hast. Wenn Du ihn nicht installiert hast, kannst du das Skript WebPlatformInstaller\Install-WebPlatformInstaller.ps1 nutzen.

## Installation

Kopiere das Module in eins der PowerShell Modulepfade.

## Verwendung

```PowerShell
# Laden des Modules
Import-Module Microsoft.Web.PlatformInstaller

# WebPlatformInstaller : Start
New-WebPlatfromInstaller

# Produkt "Url Rewrite 2.1" suchen
$Product = Get-Product -ProductId "UrlRewrite2"

# Produkt installieren
Install-Product -Product $Product
```
