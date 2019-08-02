# ----------------------------------------------------------------
# Module: Microsoft.Web.PlatformInstaller
# 
# Links:
# -> https://webgallery.microsoft.com/
# -> https://webgallery.microsoft.com/feed/webapplicationlist.xml
# -> https://webpifeed.blob.core.windows.net/webpifeed/WebProductList.xml
# -> https://webpifeed.blob.core.windows.net/webpifeed/MediaProductList.xml
# -> https://webpifeed.blob.core.windows.net/webpifeed/ToolsProductList.xml
# -> https://webpifeed.blob.core.windows.net/webpifeed/EnterpriseProductList.xml
# ----------------------------------------------------------------
[reflection.assembly]::LoadWithPartialName("Microsoft.Web.PlatformInstaller") | Out-Null

function New-WebPlatformInstaller
{
    New-ProductManager | Out-Null
    New-InstallManager | Out-Null
    $Global:WebPiLanguage = Get-Language -LanguageId ((get-culture).TwoLetterISOLanguageName)
}

function Install-Product
{
    [CmdletBinding(DefaultParameterSetName='Product', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([Object])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Product')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Web.PlatformInstaller.Product[]] $Product,

        [Microsoft.Web.PlatformInstaller.Language] $Language
    )

    if ($PSCmdlet.ParameterSetName -eq "Product")
    {
        if (-not ($Language)) { $InstallerCollection = New-InstallerCollection -product $Product } else { $InstallerCollection = New-InstallerCollection -product $Product -Language $Language }
        Set-InstallManager -InstallerCollection $InstallerCollection
        Start-Installation
    }
}

# ----------------------------------------------------------------
# ProductManager
# ----------------------------------------------------------------
#region
function New-ProductManager
{
    $Global:ProductManager = [Microsoft.Web.PlatformInstaller.ProductManager]::New()

    Register-ObjectEvent -InputObject $Global:ProductManager -EventName "WebPlatformInstallerUpdateAvailable" -Action {
        (New-Event -SourceIdentifier "WebPlatformInstallerUpdateAvailable" -Sender $args[0] –EventArguments $args[1].SourceEventArgs.NewEvent.TargetInstance)
    } | Out-Null

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
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([Microsoft.Web.PlatformInstaller.Keyword[]])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Keyword')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $Id,

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $All
    )

    if ($PSCmdlet.ParameterSetName -eq "Keyword")
    {
        $Global:ProductManager.GetKeyword($Id)
    }

    if ($PSCmdlet.ParameterSetName -eq "All")
    {
        $Global:ProductManager.Keywords
    }
}

function Get-Language
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([Microsoft.Web.PlatformInstaller.Language[]])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Language')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $LanguageId,

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $All
    )

    if ($PSCmdlet.ParameterSetName -eq "Language")
    {
        $Global:ProductManager.GetLanguage($LanguageId)
    }

    if ($PSCmdlet.ParameterSetName -eq "All")
    {
        $Global:ProductManager.Languages
    }
}

function Get-Product
{
    [CmdletBinding(DefaultParameterSetName='Language', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([Microsoft.Web.PlatformInstaller.Product[]])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Product')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $ProductId,

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Language')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $LanguageId = ((get-culture).TwoLetterISOLanguageName),

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $All
    )

    if ($PSCmdlet.ParameterSetName -eq "Product")
    {
        $Global:ProductManager.GetProduct($ProductId)
    }

    if ($PSCmdlet.ParameterSetName -eq "Language")
    {
        $Global:WebPiLanguage = Get-Language -LanguageId $LanguageId
        $Global:WebPiLanguage.AvailableProducts
    }

    if ($PSCmdlet.ParameterSetName -eq "All")
    {
        $Global:ProductManager.Products
    }
}

function Get-Tab
{
    [CmdletBinding(DefaultParameterSetName='All', 
                  SupportsShouldProcess=$false, 
                  PositionalBinding=$true,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([Microsoft.Web.PlatformInstaller.Tab[]])]
    param(
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Tab')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String] $tabId,

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='All')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Switch] $All
    )

    if ($PSCmdlet.ParameterSetName -eq "tab")
    {
        $Global:ProductManager.GetTab($tabId)
    }

    if ($PSCmdlet.ParameterSetName -eq "All")
    {
        $Global:ProductManager.Tabs
    }
}

function Update-WebPlatformBinary
{
    $Global:ProductManager.UpdateWebPlatformBinary()
}

function Test-ProductIsInstalled
{
    param(
        [Microsoft.Web.PlatformInstaller.Product] $Product
    )

    $Product.IsInstalled($false)

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
    } | Out-Null

    Register-ObjectEvent -InputObject $Global:InstallManager -EventName "InstallerStatusUpdated" -Action {
        (New-Event -SourceIdentifier "InstallerStatusUpdated" -Sender $args[0] –EventArguments $args[1].SourceEventArgs.NewEvent.TargetInstance)
    } | Out-Null

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
        $result = $Global:InstallManager.DownloadInstallerFile($InstallerContext, [ref] $failureReason)
        if (-not ($result))
        {
            throw $failureReason
        }
    
        while ($true)
        {
            $event = Wait-Event -SourceIdentifier InstallerStatusUpdated
            if ($event.Sender.InstallerContexts[0].InstallationState -eq "Downloaded")
            {
                Remove-Event -EventIdentifier $event.EventIdentifier
                break
            }
        }

        if ($InstallerContext.Product.IsApplication)
        {
            $Global:InstallManager.StartApplicationInstallation()
            while ($true)
            {
                $event = Wait-Event -SourceIdentifier InstallCompleted
                if ($event.Sender.InstallerContexts[0].InstallationState -eq "InstallCompleted")
                {        
                    $ReturnCode = $event.Sender.InstallerContexts[0].ReturnCode
                    Remove-Event -EventIdentifier $event.EventIdentifier
                    $ReturnCode
                    break
                }
            }
        }
        elseif ($InstallerContext.Product.IsIisComponent)
        {
            $Global:InstallManager.StartSynchronousInstallation()
            while ($true)
            {
                $event = Wait-Event -SourceIdentifier InstallCompleted
                if ($event.Sender.InstallerContexts[0].InstallationState -eq "InstallCompleted")
                {        
                    $ReturnCode = $event.Sender.InstallerContexts[0].ReturnCode
                    Remove-Event -EventIdentifier $event.EventIdentifier
                    $ReturnCode
                    break
                }
            }
        }
        else
        {
            $Global:InstallManager.StartInstallation()
            while ($true)
            {
                $event = Wait-Event -SourceIdentifier InstallCompleted
                if ($event.Sender.InstallerContexts[0].InstallationState -eq "InstallCompleted")
                {        
                    $ReturnCode = $event.Sender.InstallerContexts[0].ReturnCode
                    Remove-Event -EventIdentifier $event.EventIdentifier
                    $ReturnCode
                    break
                }
            }
        }
    }
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
        [Microsoft.Web.PlatformInstaller.Language] $Language = $Global:WebPiLanguage
    )
    
    $Global:InstallerCollection = [System.Collections.Generic.List[Microsoft.Web.PlatformInstaller.Installer]]::new()

    foreach ($p in $product)
    {
        if (-not (Test-ProductIsInstalled -Product $p))
        {
            $Global:InstallerCollection.Add($p.GetInstaller($Language))
        }
        else
        {
            Write-Warning "The product is already installed and was skipped."
        }
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

function Add-Installer
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

        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=1,
                   ParameterSetName='Product')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Web.PlatformInstaller.Language] $Language = $Global:WebPiLanguage
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