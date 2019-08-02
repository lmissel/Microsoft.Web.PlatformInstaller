Import-Module Microsoft.Web.PlatformInstaller

# -------------------------------------------------------------------------------------------
# PlatformInstaller : ProductManager erstellen und Produkte laden
# -------------------------------------------------------------------------------------------
$ProductManager = New-ProductManager
$culture = get-culture
$Language = Get-Language -LanguageId ($culture.TwoLetterISOLanguageName)

# -------------------------------------------------------------------------------------------
# Produkt "Web Deploy 3.6" abrufen und Installationsstatus überprüfen und ggfs. ein die 
# Auftragsliste hinzufügen.
# -------------------------------------------------------------------------------------------
$Product = Get-Product -ProductId "WDeploy36"
if (-not (Test-ProductIsInstalled -Product $Product))
{
	$InstallerCollection = New-InstallerCollection -product $Product -Language $Language
}
else
{
    Write-Verbose "Produkt [$($Product.Title)] ist bereits installiert." -Verbose
}

# -------------------------------------------------------------------------------------------
# Softwareinstallation: InstallerManager erstellen, die Aufträge übergeben und die 
# Softwareinstallationen starten
# -------------------------------------------------------------------------------------------
if ($InstallerCollection.Count -gt 0)
{
    $InstallManager = New-InstallManager
    Set-InstallManager -InstallerCollection $InstallerCollection
	Start-Installation
}