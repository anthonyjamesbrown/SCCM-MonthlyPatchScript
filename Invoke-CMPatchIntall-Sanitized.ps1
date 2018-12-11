<#
    .SYNOPSIS
        Start to finish Patching process for the specified month
    .NOTES
        Author: Anthony Brown (anthonyjamesbrown@gmail.com)
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
try
{
    #$session = New-PSSession -ComputerName "LABINFCMSS01"
    #Invoke-Command -Session $session -ScriptBlock {

        # Variables
        $Site           = "SAN"
        $SiteServer     = "SanitizedSiteServer"
        $Domain         = "Sanitized.com"
        $SiteServerFQDN = "$SiteServer.$Domain"
        $AdminPath      = split-path -Path $ENV:SMS_ADMIN_UI_PATH
        $ModulePath     = "$adminpath\ConfigurationManager.psd1"
        $WMINameSpace   = "Root\SMS\Site_$Site"
        $MaxJobQueue    = 10
        $CollectionName = "Sanitized_Collection"                   # This is the name of the collection that the SUG will be deployed to
        $DP             = "SanitizedDP.Santized.com"               # This is the name of the DP that the content will be pushed to

        # Patch time frame.  You can provide a start and end for scoping the timeframe of updates released to add to the SUG
        # The times below grab everything from July, you should be able to see how to change this to other start and end dates.
        $PatchMonthDate = get-date('12/1/2018')
        $PatchMonthEndDate = ($PatchMonthDate).AddMonths(1)

        <#
            Note: You will need to run the powershell console as a user with the required level of access to SCCM console.

            The powershell module can only be ran on a computer that has the SCCM admin console installed.
            When the console is installed it will create a enviornment varable called 'SMS_ADMIN_UI_PATH'.
            The default path is: C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386
            The path you need is one level higher in the \bin directory.  You can use split-path to remove the last
            Leaf in the path.  
            The name of the module is ConfigurationManager.psd1.
        #>

        Import-Module $ModulePath

        if((Test-Path -Path "$($Site):") -eq $false) { New-PSDrive -Name $Site -PSProvider "CMSite" -Root "$SiteServerFQDN" }

        # Change to the Pri: PSdrive.
        Set-Location -Path "$($Site):"

        # Pull the updates from the Site Server
        $Updates = Get-WMIObject -Query "Select * from SMS_SoftwareUpdate" -ComputerName $SiteServer -Namespace $WMINameSpace -ErrorAction Stop        

        # Determine which updates are in scope
        $ScopedUpdates = $Updates | Where-Object {((get-date -Date ([System.Management.ManagementDateTimeConverter]::ToDateTime($_.DateRevised))) -gt $PatchMonthDate -and (get-date -Date ([System.Management.ManagementDateTimeConverter]::ToDateTime($_.DateRevised))) -lt $PatchMonthEndDate) -and 
                        ($_.CategoryInstance_UniqueIDs -contains "Product:9f3dd20a-1004-470e-ba65-3dc62d982958" -or # This is the product id for SilverLight
                         $_.CategoryInstance_UniqueIDs -contains "Product:7f44c2a7-bc36-470b-be3b-c01b6dc5dd4e" -or # This is the prodcut id for Windows Server 2003, Datacenter Edition
                         $_.CategoryInstance_UniqueIDs -contains "Product:dbf57a08-0d5a-46ff-b30c-7715eb9498e9" -or # This is the prodcut id for Windows Server 2003
                         $_.CategoryInstance_UniqueIDs -contains "Product:fdfe8200-9d98-44ba-a12a-772282bf60ef" -or # This is the prodcut id for Windows Server 2008 R2
                         $_.CategoryInstance_UniqueIDs -contains "Product:d31bd4c3-d872-41c9-a2e7-231f372588cb"     # This is the prodcut id for Windows Server 2012 R2
                        ) -and 
                        ($_.CategoryInstance_UniqueIDs -contains "UpdateClassification:e6cf1350-c01b-414d-a61f-263d14d133b4" -or # This the the update id for Critical Updates
                         $_.CategoryInstance_UniqueIDs -contains "UpdateClassification:0fa1201d-4330-4fa8-8ae9-b877473b6441"     # This the the update id for Security Updates
                        ) -and
                        $_.CustomServerity -ne 2 -and                        # A custom severity of low in the GUI will equal 2 here.
                        $_.LocalizedDisplayName -notlike "*.Net*" -and       # Filter out .NET
                        $_.LocalizedDisplayName -notlike "*Itanium*" -and    # Filter out Itanium
                        $_.IsExpired -eq $false -and                         # Filter out Expired
                        $_.IsSuperseded -eq $false                           # Filter out Superseded
        } # end where-object

        $SUGName = "Server Updates - $(get-date($PatchMonthDate) -Format MMMM) $((get-date($PatchMonthDate)).Year) (Automation Created)" # Determine the Name to use for the new SUG based on the month.

        if($null -eq (Get-CMSoftwareUpdateGroup -Name $SUGName)) 
        {
            $NewSUG = New-CMSoftwareUpdateGroup -Name $SUGName -Description "Created by AB Automation" -UpdateId $ScopedUpdates.CI_ID    # Create the new Software Update Group

            #region Download Content
            $DLDirectory  = "ServerUpdatePackage-$((get-date($PatchMonthDate)).Year)-$(get-date($PatchMonthDate) -Format MMMM)" # Determine the name for the content download directory
            $DLFullPath   = "\\$SiteServer\MS Packages\WindowsUpdates\"                       # Set the root path for the new download directory
            $DownloadPath = "M:\$DLDirectory"                                                 # Set the relative path that will be used for coping the content

            if((Test-Path -Path "M:") -eq $false) { New-PSDrive -Name M -PSProvider FileSystem -Root $DLFullPath } # Create a local M: drive

            # Check if the folder exists, if not then create it.
            if(!(test-path -Path $downloadpath)) { $null = New-Item -ItemType directory -Path $DownloadPath }

            # Query WMI on SiteServer to build a custom object that contains the download source, destination, and ID for each of the in scope updates.
            $DownloadInfo = $ScopedUpdates | ForEach-Object {      
                $CI_ID = $_.CI_ID
                #$ContentID = Get-CimInstance -Query "Select ContentID,ContentUniqueID,ContentLocales from SMS_CITOContent Where CI_ID='$CI_ID'"  @hash
                $ContentID = Get-WMIObject -Query "Select ContentID,ContentUniqueID,ContentLocales from SMS_CITOContent Where CI_ID='$CI_ID'" -ComputerName $SiteServer -Namespace $WMINameSpace
                $ContentID = $ContentID  | Where-Object {($_.ContentLocales -eq "Locale:9") -or ($_.ContentLocales -eq "Locale:0") }

                foreach ($ID in $ContentID)
                {
                    #$ContentFile = Get-CimInstance -Query "Select FileName,SourceURL from SMS_CIContentfiles WHERE ContentID='$($ID.ContentID)'" @hash
                    $ContentFile = Get-WMIObject -Query "Select FileName,SourceURL from SMS_CIContentfiles WHERE ContentID='$($ID.ContentID)'" -ComputerName $SiteServer -Namespace $WMINameSpace
                    [pscustomobject]@{ID = $ID.ContentID; Source = $ContentFile.SourceURL ; Destination = "$DownloadPath\$($ID.ContentID)\$($ContentFile.FileName)";}
                } # end foreach
            } # end foreach-object

            # Test and create the Destination Folders if needed
            $DownloadInfo.destination | ForEach-Object -Process {
                if(-not (test-path -Path "filesystem::$(Split-Path -Path $_)"))
                {
                    $null = New-Item -ItemType directory -Path "$(Split-Path -Path $_)"
                } # end if
            } # end foreach

            # This Scriptblock is used to download the patch content from the source location and copy it to the destination.  This will be called as a job later in the script.
            $ScriptBlock = {
                param 
                (
                    [Parameter(Mandatory=$true)][string]$Source,
                    [Parameter(Mandatory=$true)][string]$Destination,
                    [Parameter(Mandatory=$true)][string]$DLFullPath       
                ) # end param

                if((Test-Path -Path "M:") -eq $false) { New-PSDrive -Name M -PSProvider FileSystem -Root $DLFullPath }
                Invoke-WebRequest -uri $Source -OutFile $Destination
            } # end scriptblock

            # Start the Download Jobs.
            $jobs = $DownloadInfo | ForEach-Object -Begin {$i=0} -Process { 
                $Args = "$($_.Source)","$($_.Destination)","$DLFullPath" 
                Write-Progress -Activity "Downloading Content: $($_.Source)" -Status "Progress: $i of $($DownloadInfo.Count)" -PercentComplete ($i/$DownloadInfo.count*100)
                While(@(Get-Job -State Running).Count -ge $MaxJobQueue)
                { 
                    Write-Warning -Message "More than $MaxJobQueue jobs in queue, waiting 2 seconds"
                    Start-Sleep -Seconds 2
                } # end while

                Start-job -Name $_.ID -ScriptBlock $ScriptBlock -ArgumentList $Args
                ++$i
            } # end foreach

            # Check every 5 seconds to see if all the download jobs are done running.
            While (@($jobs | Where-Object { $_.State -eq 'Running' }) -ne 0)
            {
                $RunningJobs = @($jobs | Where-Object { $_.State -eq 'Running' })
                Write-Progress -Activity 'Waiting for Download jobs to complete' -PercentComplete $(($jobs.Count - $RunningJobs.Count)/$($jobs.Count)*100 -as [int]) -Status "$(($jobs.Count - $RunningJobs.Count)/$($jobs.Count)*100 -as [int])% jobs finished"
                Start-Sleep 10
            } # end while

            # All jobs should be completed now.  Clean them up.
            remove-job *
            $Check = $true
            $DownloadInfo | ForEach-Object {
                $result = test-path -Path $_.Destination
                if($result -eq $false){ $Check = $false; Write-Error "$($_.Destination) was not created."}
            } # end foreach

            if($Check) 
            {
                # Create Deployment Package
                $class = Get-WmiObject -ComputerName $SiteServer -Namespace $WMINameSpace -Class SMS_SoftwareUpdatesPackage -List

                # Instantiate the Class Object 
                $DeployPackage = $class.CreateInstance()

                # Set the appropriate properties on the Instance
                $DeployPackage.Name = "$SUGName"
                $DeployPackage.SourceSite = "$Site"
                $DeployPackage.PkgSourcePath = "$DLFullPath$DLDirectory"
                $DeployPackage.Description = "$SUGName Patch Automation"
                $DeployPackage.PkgSourceFlag = [int32]2

                # Persist the changes 
                $DeployPackage.put()

                # Get the latest WMI Instance back
                $DeployPackage.get()

                # Get the Array of content source path
                $contentsourcepath = Get-ChildItem  -path "M:$DLDirectory" | Select-Object -ExpandProperty Fullname

                # Get the array of ContentIDs
                $allContentIDs =  $contentsourcepath | ForEach-Object {Split-Path  -Path $_ -Leaf}

                # Add the downloaded content for each update to the package
                $DeployPackage.AddUpdateContent($allContentIDs,$contentsourcepath,$true)

                # Final cleanup of the original source files now that they have been copied and renamed for use in the package.
                $DownloadInfo.destination | ForEach-Object -Process { Remove-Item -Path (split-path -path $_) -Recurse}  
                Start-CMContentDistribution -DeploymentPackageName $SUGName -DistributionPointName $DP
                
                $Props = @{
                    'DeploymentName'                = "$SUGName - $CollectionName";
                    'DeploymentType'                = "Required";
                    'ProtectedType'                 = "NoInstall";
                    'UnprotectedType'               = "NoInstall";
                    'TimeBasedOn'                   = "UTC";
                    'EnforcementDeadlineDay'        = (Get-Date).ToShortDateString();
                    'EnforcementDeadline'           = (Get-Date).AddMinutes(5).ToShortTimeString();
                    'RestartServer'                 = $false;
                    'AllowRestart'                  = $false;                
                    'DisableOperationsManagerAlert' = $false;
                    'DownloadFromMicrosoftUpdate'   = $False;
                    'PersistOnWriteFilterDevice'    = $False;               
                    'SoftwareInstallation'          = $False;        
                } # end hash

                Start-CMSoftwareUpdateDeployment -SoftwareUpdateGroupName $SUGName -CollectionName $CollectionName @Props
            }
            else
            {
                    Write-Error "There were missing content files."
            } # end if
        }
        else
        {
                Write-Error "A Software update group called $SUGName already exists."
        } # End If ((Get-CMSoftwareUpdateGroup -Name $SUGName) -ne $null)
   # } # Session Script Block
} catch {
    Throw $_.Exception
}