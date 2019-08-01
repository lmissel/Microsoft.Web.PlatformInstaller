# ----------------------------------------------------------------
# Module: Microsoft.Web.PlatformInstaller
# ----------------------------------------------------------------
[reflection.assembly]::LoadWithPartialName("Microsoft.Web.PlatformInstaller") | Out-Null

# https://webgallery.microsoft.com/feed/webapplicationlist.xml
# https://webpifeed.blob.core.windows.net/webpifeed/WebProductList.xml
# https://webpifeed.blob.core.windows.net/webpifeed/MediaProductList.xml
# https://webpifeed.blob.core.windows.net/webpifeed/ToolsProductList.xml
# https://webpifeed.blob.core.windows.net/webpifeed/EnterpriseProductList.xml

# ----------------------------------------------------------------
# ProductManager
# ----------------------------------------------------------------
#region
function New-ProductManager
{
    $Global:ProductManager = [Microsoft.Web.PlatformInstaller.ProductManager]::New()

    Register-ObjectEvent -InputObject $Global:ProductManager -EventName "WebPlatformInstallerUpdateAvailable" -Action {
        (New-Event -SourceIdentifier "WebPlatformInstallerUpdateAvailable" -Sender $args[0] –EventArguments $args[1].SourceEventArgs.NewEvent.TargetInstance)
    }

    # Produkte laden...
    $Global:ProductManager.Load()

    return $Global:ProductManager
}

function Get-ProductManager
{
    return $Global:ProductManager
}

function Set-ProductManager
{
    param(
        [String] $resourceLanguage
    )

    $Global:ProductManager.SetResourceLanguage($resourceLanguage)
}

function Get-Keyword
{
    param(
        [String] $Id,
        [Switch] $All
    )
    
    $Global:ProductManager.GetKeyword($Id)
    $Global:ProductManager.Keywords
}

function Get-Language
{
    param(
        [String] $LanguageId,
        [Switch] $All
    )
    
    $Global:ProductManager.GetLanguage($LanguageId)
    $Global:ProductManager.Languages
}

function Get-Product
{
    param(
        [String] $ProductId,
        [Switch] $All
    )

    $Global:ProductManager.Products
    $Global:ProductManager.GetProduct($ProductId)
}

function Get-Tab
{
    param(
        [String] $tabId,
        [Switch] $All
    )

    $Global:ProductManager.Tabs
    $Global:ProductManager.GetTab($tabId)
}

function Update-WebPlatformBinary
{
    $Global:ProductManager.UpdateWebPlatformBinary()
}
#endregion

# ----------------------------------------------------------------
# InstallManager
# ----------------------------------------------------------------
#region
function New-InstallManager
{
    $Global:InstallManager = [Microsoft.Web.PlatformInstaller.InstallManager]::New()

    Register-ObjectEvent -InputObject $Global:InstallManager -EventName "InstallCompleted" -Action {
        (New-Event -SourceIdentifier "InstallCompleted" -Sender $args[0] –EventArguments $args[1].SourceEventArgs.NewEvent.TargetInstance)
    }

    Register-ObjectEvent -InputObject $Global:InstallManager -EventName "InstallerStatusUpdated" -Action {
        (New-Event -SourceIdentifier "InstallerStatusUpdated" -Sender $args[0] –EventArguments $args[1].SourceEventArgs.NewEvent.TargetInstance)
    }

    return $Global:InstallManager
}

function Get-InstallManager
{
    return $Global:InstallManager
}

function Set-InstallManager
{
    param(
        [System.Collections.Generic.List[Microsoft.Web.PlatformInstaller.Installer]] $InstallerCollection
    )

    $Global:InstallManager.Load($InstallerCollection)
}

function Export-InstallManagerConfiguration
{
    param(
        [String] $fileLocation
    )

    $global:InstallManager.Save($fileLocation)
}

function Import-InstallManagerConfiguration
{
    param(
        [String] $ConfigurationFile
    )

    $global:InstallManager.Load($fileLocation)
}

function Start-Installation
{
    $failureReason = $null
    foreach($InstallerContext in $Global:InstallManager.InstallerContexts)
    {
        $Global:InstallManager.DownloadInstallerFile($InstallerContext, [ref]$failureReason)
    }

    $Global:InstallManager.StartInstallation()
    # überwachen...

    $Global:InstallManager.StartApplicationInstallation()
    # überwachen...

    $Global:InstallManager.StartSynchronousInstallation()
    # überwachen...
}
#endregion

# ----------------------------------------------------------------
# InstallerCollection / Warenkorb :-)
# ----------------------------------------------------------------
#region

function New-InstallerCollection
{
    param(
        [Microsoft.Web.PlatformInstaller.Product[]] $product,
        [Microsoft.Web.PlatformInstaller.Language] $Language
    )
    
    $Global:InstallerCollection = [System.Collections.Generic.List[Microsoft.Web.PlatformInstaller.Installer]]::new()

    foreach ($p in $product)
    {
        $Global:InstallerCollection.Add($p.GetInstaller($Language))
    }

    return $Global:InstallerCollection
}

function Get-InstallerCollection
{
    Begin
    {
        if (-not ($Global:InstallerCollection))
        {
            $Global:InstallerCollection = [System.Collections.Generic.List[Microsoft.Web.PlatformInstaller.Installer]]::new()
        }
    }

    Process
    {
        return $Global:InstallerCollection
    }

    End
    {
    }
}

function Set-InstallerCollection
{
    [CmdletBinding(DefaultParameterSetName='Installer', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Installer')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Web.PlatformInstaller.Installer] $Installer,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Product')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Web.PlatformInstaller.Product] $product,

        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Product')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Web.PlatformInstaller.Language] $Language
    )

    Begin
    {
        if (-not ($Global:InstallerCollection))
        {
            $Global:InstallerCollection = [System.Collections.Generic.List[Microsoft.Web.PlatformInstaller.Installer]]::new()
        }
    }

    Process
    {
        if ($PSCmdlet.ParameterSetName -eq "Installer")
        {
            $Global:InstallerCollection.Add($Installer)
            return $true
        }

        if ($PSCmdlet.ParameterSetName -eq "Product")
        {
            $Global:InstallerCollection.Add($product.GetInstaller($Language))
            return $true
        }
    }

    End
    {
        return $false
    }
}

function Remove-InstallerCollection
{
    if ($Global:InstallerCollection)
    {
        Remove-Variable $Global:InstallerCollection
    }
}

#endregion