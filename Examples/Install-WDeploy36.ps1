Import-Module Microsoft.Web.PlatformInstaller

# -------------------------------------------------------------------------------------------
# PlatformInstaller : Start
# -------------------------------------------------------------------------------------------
$ProductManager = New-ProductManager
$InstallManager = New-InstallManager
$culture = get-culture
$Language = Get-Language -LanguageId ($culture.TwoLetterISOLanguageName)

# -------------------------------------------------------------------------------------------
# "Web Deploy 3.6"
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
# Softwareinstallation
# -------------------------------------------------------------------------------------------
if ($InstallerCollection.Count -gt 0)
{
    Set-InstallManager -InstallerCollection $InstallerCollection
	Start-Installation
}