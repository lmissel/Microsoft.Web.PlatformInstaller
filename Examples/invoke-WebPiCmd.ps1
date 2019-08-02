function Invoke-WebPiCmd
{
    [CmdletBinding(DefaultParameterSetName='Default', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'https://docs.microsoft.com/de-de/dotnet/api/microsoft.web.platforminstaller',
                  ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Default')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [String]$Arguments
    )

    Begin
    {
        $fileName  = "$env:ProgramFiles\Microsoft\Web Platform Installer\WebpiCmd-x64.exe"
        if (-not (Test-Path -Path $fileName))
        {
            throw New-Object System.InvalidOperationException ("Web Platform Installer not installed exception!")
        }
    }

    Process
    {
        if ($pscmdlet.ShouldProcess($env:COMPUTERNAME, "invoke WebPiCmd command."))
        {
            try
            {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.CreateNoWindow = $true 
                $psi.UseShellExecute = $false 
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.FileName = $fileName
                $psi.Arguments = $Arguments

                $process = New-Object System.Diagnostics.Process 
                $process.StartInfo = $psi
                $process.Start() > $null
                $output = $process.StandardOutput.ReadToEnd()
                $process.StandardOutput.ReadLine()
                $process.WaitForExit() 
            
                return $output 
            }
            catch
            {
                $outputError = $process.StandardError.ReadToEnd()
                throw $_ + $outputError
            }
            finally
            {
                if ($null -ne $psi){ $psi = $null}
                if ($null -ne $process){ $process.Dispose() }
            }
        }
    }

    End
    {
    }
}