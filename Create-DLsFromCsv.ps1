# Create DLs

# RUN AS ADMIN FROM EXCHANGE MANAGEMENT SHELL

[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True)]
   [string]$CsvFile
)


#get-pssession | remove-pssession
#Get-Module -Name tmp_* | Remove-Module
#Remove-Module -name MSOnline -ea 0

Try {
	
    #$localExchange = "exch.domain.local"

	#login to EXO
	#"$(Get-Date -UFormat %m-%d-%y_%H-%M-%S_%Z) : Logging in to EXO"
	#$cloudAdminU = "test@domain.onmicrosoft.com"
	#$cloudAdminP = ConvertTo-secureString "pwd" -AsPlainText -Force
    #$cloudCred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $cloudAdminU,$cloudAdminP
    #$cloudCred = Get-Credential
	#$ps = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid -Credential $cloudCred -Authentication Basic -AllowRedirection -ea stop
	#Import-PSSession $ps -Prefix "Online" -ea stop -AllowClobber

	#Login to EX
	
    <#
    "$(Get-Date -UFormat %m-%d-%y_%H-%M-%S_%Z) : Logging in to EX"
	$adCredU = ""
	$adCredP = ConvertTo-secureString "" -AsPlainText -Force
	$adCred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $adCredU,$adCredP
	$sessionsOption = New-PSSessionOption -SkipCNCheck
	$psSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://$localExchange/powershell/" -Credential $adCred -Authentication Basic -AllowRedirection -SessionOption $sessionsOption
	Import-PSSession $psSession -Prefix "Local"
    #>
    # No CREDS login:
    #$sessionsOption = New-PSSessionOption -SkipCNCheck
    #$psSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ex-server.domain.local/powershell/" -AllowRedirection -SessionOption $sessionsOption
    #Import-PSSession $psSession -Prefix "Local"
    
	
	#Login to AD
	"$(Get-Date -UFormat %m-%d-%y_%H-%M-%S_%Z) : Importing AD Module"
    Import-Module ActiveDirectory -ea stop
    #$creds = Get-Credential -UserName "domain\user" -Message "Enter password" -ea 1
	
} Catch {

	"$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Error($($error[0].InvocationInfo.ScriptLineNumber)) Failed to login - $($error[0].Exception.Message)"
	Exit

}

$CsvUsers = Import-Csv $CsvFile

#$objMembers = $CsvUsers | Get-Member
<# 
Required Columns: PrimarySMTPAddress,DisplayName,Name,ManagedBy,Members,EmailAddresses,RequireSenderAuthenticationEnabled,Guid,LegacyExchangeDN
#>

mkdir NewDL_Logs -ea 0
Start-Transcript ".\NewDL_Logs\NewDL_$(Get-Date -UFormat %m%d%y_%H%M%S_%Z).txt"

$dlPrefix = "#DLPREFIX - "
$dlOU = "domain.local/Company/DLs"
$DomainController = "DC01.domain.local"
$customAttributeValue = "MigratedDLs"
$SitePrefix = "SITE1"
$TargetDomain = "domain.com"

foreach ($dl in $CsvUsers) {

    write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Creating $($dl.PrimarySMTPAddress)" -ForegroundColor Green
    Try {
        
        $checkDL = $null
        $currentPrefix = $dl.PrimarySMTPAddress.Split('@')[0]
        $checkDL = Get-DistributionGroup "$($SitePrefix)-$($currentPrefix)@$($TargetDomain)" -ea 0

        if ($checkDL) {
            write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : DL already exists for $($SitePrefix)-$($currentPrefix)@$($TargetDomain)"
            $newAlias = $checkDL.Alias
        } Else {
            $currentPrefix = $dl.PrimarySMTPAddress.Split('@')[0]
            $newPrimary = "$($SitePrefix)-$($currentPrefix)@$($TargetDomain)"
            $newAlias = "$($SitePrefix)-$($dl.Alias)"
            $newName = "$($dlPrefix)$($dl.Name)"
            $newDisplayName = "$($dlPrefix)$($dl.DisplayName)"
            New-DistributionGroup -DisplayName $newDisplayName -OrganizationalUnit $dlOU -Name $newName -PrimarySMTPAddress $newPrimary -Alias $newAlias -DomainController $DomainController -ErrorAction 1 | Out-Null
            Start-Sleep -Seconds 30
        }

        

        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Accessing new DL object for $($newAlias)"
        $localDL = Get-DistributionGroup $newAlias -ea 1
        $localDLCount = $localDL | Measure-Object | %{ $_.Count }
        if ($localDLCount -ne 1) {
            Write-Warning "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Skipping because there are multiple matches for $($newAlias)"
            Continue
        }



        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up group owners for $($newAlias)"
        $ownersAdded = 0
        $dlOwnersList = @()
        if (!($dl.ManagedBy -match ';')) {
            $dlOwnersList += $dl.ManagedBy
        } Else {
            $dlOwnersList += $dl.ManagedBy -split ';'
        }
		foreach ($dlowner in $dlOwnersList) {
	    	if ((Get-User $dlowner -ea 0)) {
                Set-DistributionGroup $newAlias -ManagedBy @{Add="$($dlowner)"} -ea 1
                $ownersAdded++
			} Else {
                write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : DL owner $dlowner does not exist in WORLD for $($newAlias)"
            }
        }
        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Added $ownersAdded owners for $($newAlias)"



        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up group members for $($newAlias)"
        $membersAdded = 0
		foreach ($dlmem in ($dl.Members -split ';')) {
	    	if ((Get-Recipient $dlmem -ea 0)) {
                Add-DistributionGroupMember $newAlias -Member "$($dlmem)" -ea 0
                $membersAdded++
			} Else {
                write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : DL member $dlmem does not exist in WORLD for $($newAlias)"
            }
        }
        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Added $membersAdded members for $($newAlias)"
        


        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up proxyAddresses for $($newAlias)"
		foreach ($proxy in ($dl.EmailAddresses -split '\|')) {
            if ($localDL.EmailAddresses -contains $proxy) {
                write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Proxy $proxy already exists for $($newAlias)"
            } Else {
                Try {
                    if ($proxy -cmatch "SMTP") {
                        Set-DistributionGroup $newAlias -EmailAddresses @{Add="$($proxy.ToLower())"} -ea 1    
                    } Else {
                        Set-DistributionGroup $newAlias -EmailAddresses @{Add="$($proxy)"} -ea 1
                    }
                } Catch {
                    write-host "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : ProxyError($($error[0].InvocationInfo.ScriptLineNumber)): $($newAlias) : $($error[0].Exception.Message)" -ForegroundColor Red
                }
            }
        }
        


		write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up additional properties for $($newAlias)"
        Set-DistributionGroup $newAlias -RequireSenderAuthenticationEnabled ($dl.RequireSenderAuthenticationEnabled -eq 'true') -EmailAddressPolicyEnabled $true -CustomAttribute11 $customAttributeValue -DomainController $DomainController -ea 1
        



        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up legacyexchangedn and ms-DS-ConsistencyGuid for $($newAlias)"
		$msdsConGuid = New-Object -TypeName System.Guid -ArgumentList $dl.GUID
        set-adobject $localDL.distinguishedname -replace @{'mS-DS-ConsistencyGuid' = $msdsConGUID.ToByteArray() }
        set-adobject $localDL.distinguishedname -replace @{'legacyExchangeDN' = $dl.LegacyExchangeDN }
        #$ADPath = "LDAP://" + $DomainController + "/" + $localDL.DistinguishedName
        #$ADGroup = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $ADPath
        #$ADGroup.psbase.invokeset('mS-DS-ConsistencyGuid',$msdsConGUID.ToByteArray())
		#$ADGroup.psbase.invokeset('legacyExchangeDN',$dl.legacyExchangeDN)
        #$ADGroup.psbase.CommitChanges()
        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : DL operations completed for $($newAlias)"
        
    } Catch {
        write-host "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : Error($($error[0].InvocationInfo.ScriptLineNumber)): $($dl.PrimarySMTPAddress) : $($error[0].Exception.Message)" -ForegroundColor Red
    }
}
Stop-Transcript