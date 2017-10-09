<#
.SYNOPSIS
    SCConfigMgr Application & Package reporting

.DESCRIPTION
	This script allows you to report on applications and packages in your environment for the
	purpose of clean up tasks.

.INPUTS
	The script has two switches for either Applications or Packages and requires you to enter
	the site server name.

.EXAMPLE
	.\Get-AppPackageCleanUp.ps1 -PackageType Packages -SiteServer YOURSITESERVER

	.\Get-AppPackageCleanUp.ps1 -PackageType Packages -SiteServer YOURSITESERVER -ExportCSV True

.NOTES
    FileName:    Get-AppPackgeCleanUp.ps1
    Authors:     Maurice Daly, Austin WongCarter
    Contact:     @modaly_it, @wongcarter
    Created:     2017-08-11
    Updated:     2017-09-20
    
    Version history:
    1.0.0 - (2017-08-11) Script created (Maurice Daly)
	1.0.1 - (2017-08-21) Added task sequence names as requested (Maurice Daly)
    1.0.2 - (2017-09-15) Changed logic to speed up processing (Austin WongCarter)
    1.0.3 - (2017-09-20) Added Dependent Applications, other minor improvements (Austin WongCarter)
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
	[parameter(Position = 0, HelpMessage = "Please specify whether you want to report on applications or packages")]
	[ValidateSet("Applications","Packages")]
	[string]$PackageType,
	[parameter(Position = 0, HelpMessage = "Please specify your SCCM site server")]
	[ValidateNotNullOrEmpty()]
	[string]$SiteServer,
	[parameter(Position = 0, HelpMessage = "Generated CSV output option")]
	[switch]$ExportCSV,
    [switch]$IncludeAppDependency
	
)
function DeploymentReport ($PackageType) 
{
	
	# Set required variables
	$PackageReport = @()
	$TaskSequences = Get-CMTaskSequence | Select-Object Name, PackageID, References | Where-Object {$_.References -ne $null}
	# Run package report
	if ($PackageType -eq "Packages")
	{
        $Packages = Get-CMPackage
        $Deployments = Get-CMDeployment

		Foreach ($Package in $Packages)
		{
			$TaskSequenceCount = 0
			$TaskSequenceList = $null
			
            $Deployed = ($Deployments | Where-Object { $_.PackageID -match $Package.PackageID }).count
            						
			foreach ($TaskSequence in $TaskSequences)
			{
				if (($TaskSequence | Select-Object -ExpandProperty References | Where-Object { $_.Package -contains $Package.PackageID }) -ne $null)
				{
					$TaskSequenceCount++
					$TaskSequenceList = $TaskSequenceList + $TaskSequence.Name + ";"
				}
			}
			$TaskSequenceMatch = New-Object PSObject
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Package Name' -Value $Package.Name
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Package ID' -Value $Package.PackageID
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Last Refresh Time' -Value $Package.LastRefreshTime
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Deployment References' -Value $Deployed
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Task Sequence References' -Value $TaskSequenceCount
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Task Sequences' -Value $TaskSequenceList
			$PackageReport += $TaskSequenceMatch
		}
		Return $PackageReport
	}
	
	# Run application report
	if ($PackageType -eq "Applications")
	{
        $Applications = Get-CMApplication
        $Dependencies = @()

        ## Build Dependency Table
        if($IncludeAppDependency)
        {
            foreach ($Application in $Applications)
		    {
                $Application | Get-CMDeploymentType | where { $_.NumberOfDependedDTs -gt 0} | Get-CMDeploymentTypeDependencyGroup | Get-CMDeploymentTypeDependency|`
                    % {$Dependencies += New-Object -TypeName PSObject -Property @{'ApplicationID'="$($Application.CI_ID)";'DependencyAppID'=$(($_ | Get-CMApplication).CI_ID);'ApplicationDisplayName'="$($Application.LocalizedDisplayName)"}}
            }
            $Dependencies = $Dependencies | select * -Unique
		}
        
        foreach ($Application in $Applications)
		{
			$TaskSequenceCount = 0
			$TaskSequenceList = $null
						
			foreach ($TaskSequence in $TaskSequences)
			{
				If ($($TaskSequence.References.Package) -contains $Application.ModelName)
				{
					$TaskSequenceCount++
					$TaskSequenceList = $TaskSequenceList + $TaskSequence.Name + ";"
				}
			}
            
            $DepCount = 0
            $DependentAppNames = ""
            $Dependencies | where {$_.DependencyAppID -eq $Application.CI_ID} | % {$DependentAppNames += "$($_.ApplicationDisplayName);"; $DepCount++}

            $TaskSequenceMatch = New-Object PSObject
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Application Name' -Value $Application.LocalizedDisplayName
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Application ID' -Value $Application.CI_ID
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Application Deployed' -Value $Application.IsDeployed
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Task Sequence References' -Value $TaskSequenceCount
			$TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Task Sequences' -Value $TaskSequenceList
            if($IncludeAppDependency)
            {
                $TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Dependent Application Count' -Value $DepCount
                $TaskSequenceMatch | Add-Member -type NoteProperty -Name 'Dependent Applications' -Value $DependentAppNames
            }
			$PackageReport += $TaskSequenceMatch
		}
        Return $PackageReport
	}
}

function ConnectSCCM ($SiteServer)
{
	
	if ((Test-WSMan -ComputerName $SiteServer).wsmid -ne $null)
	{
		# Import SCCM PowerShell Module
		$ModuleName = (Get-Item $env:SMS_ADMIN_UI_PATH).parent.FullName + "\ConfigurationManager.psd1"
		if ($ModuleName -ne $null)
		{
			Import-Module $ModuleName
			$SiteCode = QuerySiteCode -SiteServer $SiteServer
			Return $SiteCode
		}
		else
		{
			Write-Error "Error: ConfigMgr PowerShell Module Not Found" -Severity 3
		}
	}
	else
	{
		Write-Error "Error: ConfigMgr Server Specified Not Found - $SiteServer" -Severity 3
	}
}

function QuerySiteCode ($SiteServer)
{
	try
	{
		$SiteCodeObjects = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer -ErrorAction Stop
		$SiteCodeError = $false
	}
	Catch
	{
		$SiteCodeError = $true
	}
	
	if (($SiteCodeObjects -ne $null) -and ($SiteCodeError -ne $true))
	{
		foreach ($SiteCodeObject in $SiteCodeObjects)
		{
			if ($SiteCodeObject.ProviderForLocalSite -eq $true)
			{
				$SiteCode = $SiteCodeObject.SiteCode

			}
		}
		Return $SiteCode
	}
}

function Get-ScriptDirectory
{
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

# Get current directory
[string]$CurrentDirectory = (Get-ScriptDirectory)

# Connect to the SCCM environment and discover the site code
$SiteCode = ConnectSCCM ($SiteServer)
# Set the location to the site code
Set-Location -Path ($SiteCode + ":")

# Start deployment report process
$Report = DeploymentReport ($PackageType)
Set-Location -Path $CurrentDirectory

if ($ExportCSV -eq $true)
{
	if ($PackageType -eq "Applications")
	{
		$Report[1 .. ($Report.Count - 1)] | Export-CSV -Path .\Application-CleanUpReport.csv -NoTypeInformation -Append
	}
	else
	{
		$Report[1 .. ($Report.Count - 1)] | Export-CSV -Path .\Package-CleanUpReport.csv -NoTypeInformation -Append
	}
}
else
{
	$Report | Out-GridView
}