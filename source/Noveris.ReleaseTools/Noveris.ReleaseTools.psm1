
#Requires -Modules @{"ModuleName"="VMware.VimAutomation.Core";"RequiredVersion"="12.1.0.16997984"}

################
# Global settings
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
Set-StrictMode -Version 2

<#
#>
Function Get-VMIPv4Addresses
{
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNull()]
        [VMware.VimAutomation.ViCore.Impl.V1.VM.UniversalVirtualMachineImpl]$VM,

        [Parameter(mandatory=$false)]
        [switch]$First = $false
    )

    process
    {
        $addresses = $VM.Guest.IPAddress -match "^.*[.].*[.].*[.].*$"
        $addrCount = ($addresses | Measure-Object).Count
        Write-Verbose "Found $addrCount addresses for VM"

        if ($First)
        {
            if ($addrCount -lt 1)
            {
                Write-Error "Missing IPv4 addresses for system"
            }

            $single = $addresses | Select-Object -First 1
            Write-Verbose "First IPv4 address for VM: $single"
            $single
        } else {
            $addresses
        }
    }
}

<#
#>
Function Invoke-ScriptRetry
{
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$false)]
        [ValidateNotNull()]
        [int]$Attempts = 10,

        [Parameter(mandatory=$false)]
        [ValidateNotNull()]
        [int]$WaitSeconds = 5,

        [Parameter(mandatory=$true)]
        [ValidateNotNull()]
        [ScriptBlock]$Script
    )

    process
    {
        $attempt = 1
        while ($true)
        {
            try {
                $result = Invoke-Command -ScriptBlock $Script
                $result
                break
            } catch {
                Write-Information "Error running script block (attempt $attempt): $_"
                if ($attempt -ge $Attempts)
                {
                    throw $_
                } else {
                    Write-Information ("Error: " + ($_ | Out-String))
                }
            }

            Write-Information "Waiting $WaitSeconds seconds..."
            Start-Sleep $WaitSeconds
            $attempt++
        }
    }
}

<#
#>
Function New-ReleaseEnvVM
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Prefix,

        [Parameter(mandatory=$true)]
        [ValidateNotNull()]
        [HashTable]$VMArgs
    )

    process
    {
        $newArgs = $VMArgs.Clone()

        $notes = ("ReleaseEnv:{0}:AutoRemove" -f $Prefix)
        if ($VMArgs.Keys -contains "OSCustomizationSpec" -and $null -ne $VMArgs["OSCustomizationSpec"])
        {
            $notes += ":OSCustomizationSpec"
        }

        # Apply datastore location, if supplied
        if (![string]::IsNullOrEmpty($Env:RELEASEENV_VCENTER_DATASTORE) -and $newArgs.Keys -notcontains "Datastore")
        {
            $newArgs["Datastore"] = Get-Datastore $Env:RELEASEENV_VCENTER_DATASTORE
        }

        # Apply VMhost specification, if supplied
        if (![string]::IsNullOrEmpty($Env:RELEASEENV_VCENTER_VMHOST) -and $newArgs.Keys -notcontains "VMHost")
        {
            $newArgs["VMHost"] = $Env:RELEASEENV_VCENTER_VMHOST
        }

        $newArgs["Name"] = ("{0}-{1}" -f $Prefix, $newArgs["Name"])
        $newArgs["Notes"] = $notes

        if ($PSCmdlet.ShouldProcess($newArgs["Name"], "Create"))
        {
            New-VM @newArgs
        }
    }
}

<#
#>
Function Get-ReleaseEnvPrefix
{
    [CmdletBinding()]
    param(
    )

    process
    {
        $prefix = $Env:RELEASEENV_PREFIX
        if (![string]::IsNullOrEmpty($prefix))
        {
            $prefix
            return
        }

        $prefix = ($Env:BUILD_REPOSITORY_NAME + "-" + $Env:BUILD_BUILDID)
        if ([string]::IsNullOrEmpty($prefix) -or $prefix -eq "-")
        {
            Write-Error "Missing RELEASEENV_PREFIX and could not determine valid prefix"
            return
        }

        $prefix
    }
}

enum VMCustomiseState {
    Unknown = 0
    Started
    Succeeded
    Failed
}

<#
#>
Function Start-ReleaseEnv
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Prefix,

        [Parameter(mandatory=$false)]
        [ValidateNotNull()]
        [int]$TimeoutMinutes = 20
    )

    process
    {
        $vms = Get-ReleaseEnvVMs -Prefix $Prefix

        # Start all VMs
        $list = @()
        $vms | ForEach-Object {
            $name = $_.Name

            if ($PSCmdlet.ShouldProcess($_.Name, "Start"))
            {
                Write-Information "Starting VM: $name"
                $list += $_
                $_ | Start-VM | Out-Null
            }
        }
        $vms = $list

        # Only work on VMs that had an OSCustomisationSpec applied
        $vms = $vms | Where-Object {$_.Notes.Contains("OSCustomizationSpec")}

        Write-Information "VMs Pending Customisations:"
        $vms | ForEach-Object {Write-Information $_}

        # Exit here if there are no VMs to customise
        if (($vms | Measure-Object).Count -lt 1)
        {
            Write-Information "no VMs to customise"
            return
        }

        # Wait for VMs to perform customisation
        $start = [DateTime]::Now
        $state = @{}
        $vms | ForEach-Object { $state[$_] = [VMCustomiseState]::Unknown }

        while ([DateTime]::Now -lt $start.AddMinutes($TimeoutMinutes))
        {
            # Initialise counters
            $unknownCount = 0
            $startedCount = 0
            $succeededCount = 0
            $failedCount = 0

            # Review state of all VMs
            $keys = $state.Keys | ForEach-Object { $_ }
            foreach ($key in $keys)
            {
                # Check for status update on VM if the current state is unknown or started
                if ($state[$key] -eq [VMCustomiseState]::Unknown -or $state[$key] -eq [VMCustomiseState]::Started)
                {
                    $events = Get-VIEvent -Entity $key

                    if (($events | Where-Object {$_ -is "VMware.Vim.CustomizationSucceeded"} | Measure-Object).Count -gt 0)
                    {
                        $state[$key] = [VMCustomiseState]::Succeeded
                    } elseif (($events | Where-Object {$_ -is "VMware.Vim.CustomizationFailed"} | Measure-Object).Count -gt 0)
                    {
                        $state[$key] = [VMCustomiseState]::Failed
                    } elseif (($events | Where-Object {$_ -is "VMware.Vim.CustomizationStartedEvent"} | Measure-Object).Count -gt 0)
                    {
                        $state[$key] = [VMCustomiseState]::Started
                    }
                }

                # Update totals
                if ($state[$key] -eq [VMCustomiseState]::Unknown) {
                    $unknownCount++
                } elseif ($state[$key] -eq [VMCustomiseState]::Started) {
                    $startedCount++
                } elseif ($state[$key] -eq [VMCustomiseState]::Succeeded) {
                    $succeededCount++
                } elseif ($state[$key] -eq [VMCustomiseState]::Failed) {
                    $failedCount++
                }
            }

            Write-Information ("State -> Unknown({0}), Started({1}), Succeeded({2}), Failed({3}): Elapsed Time: {4} minutes" -f $unknownCount,
                $startedCount, $succeededCount, $failedCount, ([DateTime]::Now - $start).TotalMinutes.ToString("0.00"))

            if ($unknownCount -lt 1 -and $startedCount -lt 1)
            {
                break
            }

            Start-Sleep 15
        }

        $deployFailed = $false

        # Check if we had any unknowns
        $unknown = $state.Keys | Where-Object { $state[$_] -eq [VMCustomiseState]::Unknown }
        if (($unknown | Measure-Object).Count -gt 0)
        {
            $deployFailed = $true
            Write-Information ("VMs never started customisation: " + ($unknown.Name -join ", "))
        }

        # Check if we had any unfinished VMs
        $started = $state.Keys | Where-Object { $state[$_] -eq [VMCustomiseState]::Started }
        if (($started | Measure-Object).Count -gt 0)
        {
            $deployFailed = $true
            Write-Information ("VMs started, but did not finish customisation: " + ($started.Name -join ", "))
        }

        # Check if we had any failures
        $failed = $state.Keys | Where-Object { $state[$_] -eq [VMCustomiseState]::Failed }
        if (($failed | Measure-Object).Count -gt 0)
        {
            $deployFailed = $true
            Write-Information ("VMs failed customisation: " + ($failed.Name -join ", "))
        }

        if ($deployFailed)
        {
            Write-Error "Customisation did not complete successfully."
        } else {
            Write-Information "Customisation successful."
        }
    }
}

<#
#>
Function Get-ReleaseEnvVMs
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Prefix
    )

    process
    {
        Get-VM | Where-Object {$_.Name.StartsWith($Prefix) -and $_.Notes.Contains("ReleaseEnv:" + $Prefix)}
    }
}

<#
#>
Function Stop-ReleaseEnv
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Prefix
    )

    process
    {
        Get-ReleaseEnvVMs -Prefix $Prefix | Where-Object {$_.Notes.Contains("AutoRemove")} | ForEach-Object {
            if ($PSCmdlet.ShouldProcess($_.Name, "DELETE"))
            {
                Write-Information "Stopping VM: $_"
                Stop-VM $_ -Confirm:$false -EA Ignore | Out-Null
                Start-Sleep 1
                Write-Information "Removing VM: $_"
                Remove-VM -DeletePermanently -Confirm:$false $_ | Out-Null
            }
        }
    }
}

<#
#>
Function New-VMwareSession {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$vCenter = $Env:RELEASEENV_VCENTER_HOST,

        [Parameter(mandatory=$false)]
        [ValidateNotNull()]
        [PSCredential]$Credential,

        [Parameter(mandatory=$false)]
        [switch]$NoPrompt = $false
    )

    process
    {
        # Check for vcenter server
        if ([string]::IsNullOrEmpty($vCenter))
        {
            Write-Error "Missing vCenter server name"
        }

        # Check for credentials
        if ($PSBoundParameters.Keys -notcontains "Credentials")
        {
            # Credentials not supplied via parameters
            $username = $Env:RELEASEENV_VCENTER_USERNAME
            $password = $Env:RELEASEENV_VCENTER_PASSWORD
            if ([string]::IsNullOrEmpty($username) -and [string]::IsNullOrEmpty($password))
            {
                if ($NoPrompt)
                {
                    Write-Error "Missing Credential parameter, no credential environment variables set and prompting not allowed"
                } else {
                    $Credential = Get-Credential -Title "VMware Credentials"
                    if ($null -eq $Credential) {
                        Write-Error "Invalid or no credentials supplied"
                    }
                }
            } else {
                $netcred = [System.Net.NetworkCredential]::new($username, $password)
                $Credential = [PSCredential]::new($netcred.Username, $netcred.SecurePassword)
            }
        }

        if ($PSCmdlet.ShouldProcess("vCenter", "Connect"))
        {
            # Disconnect any existing instances
            try {
                $response = Disconnect-VIServer * -Force -Confirm:$false -EA Ignore
                Write-Verbose ($response | Out-String)
            } catch {
                # Don't really care if this fails. Best effort
                Write-Verbose "Disconnect-VIServer threw error: $_"
            }

            # Configure PowerCli session settings
            try {
                Write-Information "Configuring PowerCLI session settings"
                $response = Set-PowerCliConfiguration -DefaultVIServerMode single -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false -Scope Session -Confirm:$false
                Write-Verbose ($response | Out-String)
            } catch {
                Write-Information "Error settings powercli session settings"
                throw $_
            }

            # Connect to VIServer
            try {
                Write-Information "Connecting to vCenter server: $vCenter"
                $response = Connect-VIServer -Server $vCenter -Credential $Credential
                Write-Verbose ($response | Out-String)
            } catch {
                Write-Information "Error connecting to vCenter server"
                throw $_
            }
        }
    }
}

<#
#>
Function Get-VMSourceSnapshot
{
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNull()]
        $VM,

        [Parameter(mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    process
    {
        $vmRef = Get-VM $VM
        $snapshots = $vmRef | Get-Snapshot

        # If snapshot requested by name, try to return that one
        if ($PSBoundParameters.Keys -contains "Name")
        {
            $snapshots = $snapshots | Where-Object {$_.Name -eq $Name}

            if (($snapshots | Measure-Object).Count -lt 1)
            {
                Write-Error "Could not find referenced snapshot name: $Name"
            }

            $snapshots | Select-Object -First 1

            return
        }

        # If there are no snapshots, create one
        if (($snapshots | Measure-Object).Count -lt 1)
        {
            $vmRef | New-Snapshot -Name ([DateTime]::Now.ToString("yyyyMMdd_HHmm-AutoRelEnv")) | Out-Null
            $snapshots = $vmRef | Get-Snapshot
        }

        # Return the latest snapshot
        $vmRef | Get-Snapshot | Sort-Object -Property {$_.ExtensionData.CreateTime} -Descending | Select-Object -First 1
    }
}
