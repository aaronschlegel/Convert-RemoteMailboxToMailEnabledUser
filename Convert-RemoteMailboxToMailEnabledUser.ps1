
param(
    $DomainController,
    $AdConnectServer,
    [Parameter(Mandatory = $true)]
    [string]$Csvfile
)


#Requires -Version 5.1
#Requires -RunAsAdministrator
function Remove-O365DirectAssignedLicense {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName = 'Array')]
        [array]$userUPNArray,
        [Parameter(ParameterSetName = 'FileInput')]
        [ValidateScript({
                if ( -Not ($_ | Test-Path) ) {
                    throw "File or folder does not exist"
                }
                return $true
            })]
        [System.IO.FileInfo]$userUPNFilePath,
        [Parameter(ParameterSetName = 'SingleUser')]
        [ValidateNotNullOrEmpty()]
        [string]$user,
        [Parameter(ParameterSetName = 'AllUsers')]
        [switch]$processAllUsers
    )

    #Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users.Actions
    #Requires -Version 5.1

   
    Import-Module Microsoft.Graph.Users
    Import-module Microsoft.Graph.Identity.DirectoryManagement
    Import-Module Microsoft.Graph.Users.Actions


    try {

        Connect-Graph -Scopes User.ReadWrite.All, Organization.Read.All

    }
    catch {

        write-host -ForegroundColor Red "Error connecting to Graph $_"
    }

    $allSkus = Get-MgSubscribedSku -all | select -exp skupartnumber
    $allUsers = $Null 

    $processedUsers = @()

    if ($user) {
        $allusers = $user 
    }
    elseif ($userUPNArray) {
        
        $allUsers = $userUPNArray
    }
    elseif ($userUPNFilePath) {

        $allUsers = get-content $userUPNFilePath
    }
    elseif ($processAllUsers.IsPresent) {

        $allUsers = (get-MgUser -all -Property "userprincipalname").userprincipalname


    }
    else {

        write-host -ForegroundColor Red "Please specifiy users for input"
        sleep 5 
        exit
    }

    for ($i = 0; $i -lt $allSkus.count - 1 ; ++$i ) {

        write-host "$($i + 1). $($allSkus[$i])"

    }

    $skuToremoveidx = read-host "Select License to Remove Plan from"

    $skuToremove = $allSkus[$skuToremoveidx - 1]


    $planSkus = (Get-MgSubscribedSku -all |  ? { $_.skupartnumber -eq $skuToremove } | select -exp serviceplans).serviceplanName

    for ($i = 0; $i -lt $planSkus.count - 1; ++$i ) {

        write-host "$($i + 1). $($planSkus[$i])"

    }

    $servicePlanToRemoveidx = read-host "Select Service Plan to remove"

    $servicePlanToRemove = $planSkus[$servicePlanToRemoveidx - 1]

    write-host -ForegroundColor Magenta "Found $($allUsers.count) users to remove $servicePlanToRemove from $skuToremove"
    $cont = Read-Host "Continue(y/n)?"

    if ($cont -ne "y") {

        exit;
    }




    foreach ($o365User in $allUsers) {
        ## Get the services that have already been disabled for the user.
        $o365User = $o365user.replace("'", "''")
        $userLicense = $Null 

        $userLicense = Get-MgUserLicenseDetail -UserId "$o365User" | ? { $_.skupartnumber -eq $skuToremove }



        $userDisabledPlans = $userLicense.ServicePlans | ? { $_.ProvisioningStatus -eq "Disabled" } | Select -ExpandProperty ServicePlanId
        if ($userLicense.Id.count -gt 0) {


            write-host -ForegroundColor Green "$o365User has $userDisabledPlans disabled"


            ## Get the new service plans that are going to be disabled
            $skuInfo = Get-MgSubscribedSku -All | Where SkuPartNumber -eq $skuToremove

            $newDisabledPlans = $Null 

            $newDisabledPlans = $skuInfo.ServicePlans | ? { $_.ServicePlanName -in ($servicePlanToRemove) } | Select -ExpandProperty ServicePlanId

            if ($null -ne $newDisabledPlans) {

                ## Merge the new plans that are to be disabled with the user's current state of disabled plans
                $disabledPlans = ($userDisabledPlans + $newDisabledPlans) | Select -Unique
                write-host -ForegroundColor Yellow "Disabling $newDisabledPlans"
                $addLicenses = @(
                    @{
                        SkuId         = $skuInfo.SkuId
                        DisabledPlans = $disabledPlans
                    }
 
                )
                write-host -ForegroundColor Cyan "Setting disabled Plans $($addLicenses.disabledPlans)"
                ## Update user's license
                try {
                    Set-MgUserLicense -UserId $o365User -AddLicenses $addLicenses -RemoveLicenses @()

                    $processedUsers += $o365User
                }
                catch {

                    write-host -ForegroundColor Red "Error removing license for $o365User $_"
                }

            }
            else {
            
                write-host -ForegroundColor Cyan "$o365 user does not have entitilement $servicePlanToRemove enabled"
            }

        }
        else {
            write-host -ForegroundColor Red "$o365User has no licenses"

        }

    }

    return $processedUsers

}
function Run-DirSync {
	param (
		$type = 'Delta',
		$AdConnectServer = $defaultAdConnectServer
	)



	if ($type -eq 'Delta') {
		Invoke-Command -ComputerName $AdConnectServer -ScriptBlock { Import-Module ADSync 
			Start-ADSyncSyncCycle -PolicyType Delta }
	}
	if ($type -eq 'Full') {
		Invoke-Command -ComputerName $AdConnectServer -ScriptBlock { Import-Module ADSync 
			Start-ADSyncSyncCycle -PolicyType Initial }

	}
}

Get-packageprovider nuget -force
if(-not(get-module Microsoft.graph.users -ListAvailable)) {
    if(-not(Get-PackageProvider nuget)) {
        Install-PackageProvider nuget -force
    }
    $installedGraphUsers = $true
    Install-Module -Name Microsoft.Graph.Users -force
}
if(-not(get-module Microsoft.graph.identity.directorymanagement -ListAvailable)) {
         if(-not(Get-PackageProvider nuget)) {
        Install-PackageProvider nuget -force
    }
    $installedGraphIdent = $true
    Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -force
}

add-pssnapin *exchange*

$MailboxList = import-csv $Csvfile

#remove archive mailbox
foreach ($Mbx in $MailboxList) {

    $escapedUpn = $mbx.SourceUPN.Replace("'", "''")
    Disable-RemoteMailbox   $escapedUpn -IgnoreLegalHold -Archive -Confirm:$false -erroraction silentlycontinue

}

Remove-O365DirectAssignedLicense -userUPNArray $([array]$MailboxList.SourceUPN)

Run-DirSync -type Delta -AdConnectServer $AdConnectServer

#remove mailbox and create mail user
foreach($mbx in $MailboxList) {

    $escapedUpn = $mbx.SourceUPN.Replace("'", "''")
    #Gets the custom attributes so they can be applied to the mailuser
    $Mailbox = Get-RemoteMailbox   $escapedUpn -DomainController $DomainController 
    #Disables the remote mailbox
    Disable-RemoteMailbox   $escapedUpn -IgnoreLegalHold -Confirm:$false -erroraction silentlycontinue
    #Disables the remote mailbox
    
    #Enables as a mail user
    try {
        Enable-MailUser -Identity $escapedUpn -ExternalEmailAddress $mbx.TargetMailbox -DomainController $DomainController
    }
    catch {

        write-host "Error enabling mail user for $($escapedUpn)"
        Continue 
    }

    $customAttributes = @{}

    # Loop through properties that start with "Custom" in $Mailbox
    foreach ($property in $Mailbox.PSObject.Properties) {
        if ($property.Name -like "Custom*") {
            $customAttributes[$property.Name] = $property.Value
        }
    }

    #Sets the custom attributes back on the mailuser that were defined on the remote mailbox
    Set-MailUser -Identity $escapedUPN -DomainController $DomainController @customAttributes
}

Run-DirSync -type Delta -AdConnectServer $AdConnectServer

if($installedGraphUsers) {
    Uninstall-Module -Name Microsoft.Graph.Users -force
}

if($installedGraphIdent) {
    Uninstall-Module -Name Microsoft.Graph.Identity.DirectoryManagement -force
}