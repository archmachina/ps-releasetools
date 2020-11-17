<#
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Stages
)

########
# Global settings
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2

########
# Modules
Remove-Module Noveris.ModuleMgmt -EA SilentlyContinue
Import-Module ./Noveris.ModuleMgmt/source/Noveris.ModuleMgmt/Noveris.ModuleMgmt.psm1

Remove-Module noveris.build -EA SilentlyContinue
Import-Module -Name noveris.build -RequiredVersion (Install-PSModuleWithSpec -Name noveris.build -Major 0 -Minor 4)

########
# Capture version information
$version = Get-BuildVersionInfo -Sources @(
    $Env:GITHUB_REF,
    $Env:BUILD_SOURCEBRANCH,
    $Env:CI_COMMIT_TAG,
    $Env:BUILD_VERSION,
    "v0.1.0"
)

########
# Build stage
Invoke-BuildStage -Name "Build" -Filters $Stages -Script {
    # Template PowerShell module definition
    Write-Information "Templating Noveris.ReleaseTools.psd1"
    Format-TemplateFile -Template source/Noveris.ReleaseTools.psd1.tpl -Target source/Noveris.ReleaseTools/Noveris.ReleaseTools.psd1 -Content @{
        __FULLVERSION__ = $version.Full
    }

    # Copy modulemgmt to release tools to include in module
    Write-Information "Copy ModuleMgmt to source directory"
    Copy-Item ./Noveris.ModuleMgmt/source/Noveris.ModuleMgmt/Noveris.ModuleMgmt.psm1 ./source/Noveris.ReleaseTools/

    # Attempt to import the module
    Write-Information "Attempting to import module"
    Import-Module ./source/Noveris.ReleaseTools -EA Stop
}

Invoke-BuildStage -Name "Publish" -Filters $Stages -Script {
    # Publish module
    Publish-Module -Path ./source/Noveris.ReleaseTools -NuGetApiKey $Env:NUGET_API_KEY
}
