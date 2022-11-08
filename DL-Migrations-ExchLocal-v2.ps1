# DL Migration
[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True)]
   [string]$CsvFile
)
# Required column: EmailAddress
# Requires ActiveDirectory module and Exchange Management Shell
# Define routing address below in $routingDomain before running the script

function IsValidEmail { 
    param([string]$EmailAddress)

    try {
        $null = New-Object System.Net.Mail.MailAddress $EmailAddress
        return $true
    }
    catch {
        return $false
    }
}

Function ExpandGroup ($DistGroup, $CheckNested) {
    $membersList = @()
    #$nonMailEnabled = @()
    
    foreach ($Group in $DistGroup)
    {
        $nestedList = @()
        $CurMems = Get-DistributionGroupMember -identity $Group.Identity -resultsize unlimited
        foreach ( $Member in $CurMems) {
            $NewMember = New-Object PSObject -Property @{ DistGrpName = $Group.DisplayName ;
            DistGrpAlias = $Group.Alias ; 
            DistGrpPrimarySMTP = $Group.PrimarySmtpAddress ;
            GrpMemberPrimarySMTP = $Member.PrimarySmtpAddress ; 
            GrpMemberAlias = $Member.Alias ; 
            GrpMemberRecTypeDetail = $Member.RecipientTypeDetails ; 
            GrpMemberRecType = $Member.RecipientType ; 
            GrpMemberDisplayName = $Member.DisplayName ; 
            GrpMemberSamAccName = $Member.SamAccountName }
            if ($CheckNested) {
                if ($Member.RecipientTypeDetails -match 'group' -and $Member.RecipientTypeDetails -match 'mail') {
                    $dgObject = get-distributiongroup -identity $Member.Identity -ea 0
                    $membersList += ExpandGroup -DistGroup $dgObject -CheckNested $true
                    $nestedListObject = "" | Select GroupSMTP,MemberDLSMTP
                    $nestedListObject.GroupSMTP = $Group.PrimarySMTPAddress
                    $nestedListObject.MemberDLSMTP = $dgObject.PrimarySMTPAddress
                    $nestedList += $nestedListObject

                } <# Else {
                    $obj = "" | select Group,Member
                    $obj.Group = $Group.PrimarySMTPAddress
                    $obj.Member = $Member.SamAccountName
                    $nonMailEnabled += $obj
                } #>
            }
            $membersList += $NewMember
        }
        if ($nestedList) { $nestedList | Export-Csv ".\GroupExports\NestedDL_$($Group.SamAccountName)_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation }
    }
    #if ($nonMailEnabled) { $nonMailEnabled | Export-Csv ".\GroupExports\NonMailEnabledGrpMems$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation }
    
    return $membersList
    
}

Import-Module ActiveDirectory -ea 1
Set-ADServerSettings -ViewEntireForest:$true
$routingDomain = "test.mail.onmicrosoft.com"
$logpath = ".\GroupExports"
if (!(Test-Path $logpath)) {
	mkdir $logpath | out-null
}
Start-Transcript -Path ".\GroupExports\DLOperations_Log_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).txt" | Out-Null

do {
	
	""
	Write-Host "Distribution Group Operations" -foregroundcolor Green
    ""
    Write-Host "1. Export local DLs provided from CSV" -foregroundcolor White
    Write-Host "2. Remove local DLs" -foregroundcolor White
    Write-Host "3. Create local contacts" -foregroundcolor White
    write-host "4. Fix memberships of DLs" -foregroundcolor White
    write-host "5. Get current MemberOf of DL contacts" -foregroundcolor White
	Write-Host "6. Exit" -foregroundcolor White
	""
	$reply = Read-Host "Select your choice"

	if ($reply -eq '6') {
        ""
        Write-Host "Goodbye!" -ForegroundColor Magenta
        ""
		break
    }

    if ($reply -eq '1') {
        ""
        $groups = Import-Csv $CsvFile -ea 1
        $groupCount = $groups | measure | %{ $_.Count }
        "Group count in CSV: $($groupCount)"
        #------DISTRIBUTION GROUPS AND GROUP MEMBERS EXPORT (Not compatible with Powershell v1.0)
        Write-Host "Getting Distribution groups.." -NoNewline
        $DistGroup = $groups | %{ Get-DistributionGroup $_.EmailAddress -ea 0 }
        Write-Host "Done" -ForegroundColor Green

        if (!$DistGroup) {
            "Error: DL lookup was empty"
            break
        }
        $lookupCount = $DistGroup | measure | %{ $_.Count }
        "Groups found by lookup: $($lookupCount)"
        if ($groupCount -ne $lookupCount) {
            write-host 'DLs not found in Exchange:'
            Compare-Object -ReferenceObject $groups -DifferenceObject ($DistGroup | select @{n='EmailAddress';e={$_.PrimarySMTPAddress}}) -PassThru | select -expand EmailAddress
            write-host "WARNING: Some DLs were not found." -foregroundcolor Yellow
        }
        $secCheck = $DistGroup | where{$_.RecipientTypeDetails -match 'security'}
        $secCheckCount = $secCheck | measure | %{$_.Count}
        if ($secCheckCount -ge 1) { write-host "WARNING: Security groups present in current DL batch." -foregroundcolor Yellow } else { "No security groups present in current batch." }

        mkdir "GroupExports" -ea 0 | out-null
        "Exporting group csv file"
        $DistGroup | Export-Csv ".\GroupExports\DistGroup_Export_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        "Exporting group xml file"
        $DistGroup | Export-CliXml ".\GroupExports\DistGroup_XmlExport_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).xml"
        "Processing group attributes"
        $expandedExport = @()
        $moStoreExport = @()
        foreach ($dl in $DistGroup) {
            $obj = "" | select PrimarySMTPAddress,GroupType,SamAccountName,ManagedBy,Alias,OrganizationalUnit,DisplayName,EmailAddresses,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,LegacyExchangeDN,RecipientTypeDetails,Name,DistinguishedName,ExchangeVersion,Guid,RequireSenderAuthenticationEnabled,RoutingAddress,Members,AcceptMessagesOnlyFromSendersOrMembers,MemberOf,Notes
            $obj.PrimarySMTPAddress = $dl.PrimarySMTPAddress
            $obj.GroupType = $dl.GroupType
            $obj.SamAccountName = $dl.SamAccountName
            $managedByStore = @()
            foreach ($mbUser in $dl.ManagedBy) {
                $gUser = $null
                $gUser = Get-User $mbUser -ea 0
                if ($gUser.WindowsEmailAddress) {
                    $managedByStore += $gUser.WindowsEmailAddress
                } Else {
                    $obj.Notes += "ManByNotFound: $($gUser.Name)"
                }
            }
            $obj.ManagedBy = $managedByStore -join ";"
            $obj.Alias = $dl.Alias
            $obj.OrganizationalUnit = $dl.OrganizationalUnit
            $obj.DisplayName = $dl.DisplayName
            $obj.EmailAddresses = $dl.EmailAddresses -join "|"
            $routingAddress = $dl.EmailAddresses | where{$_ -match 'mail.onmicrosoft.com'}
            if (!$routingAddress) { $obj.RoutingAddress = "$($dl.Alias)@$($routingDomain)"; $obj.Notes += "RtgAddNotFound. " } else { $obj.RoutingAddress = ($routingAddress.ToString()).Split(':')[1] }
            $grantSendStore = @()
            foreach ($gsUser in $dl.GrantSendOnBehalfTo) {
                $gUser = $null
                $gUser = Get-User $gsUser -ea 0
                $grantSendStore += $gUser.WindowsEmailAddress
            }
            $obj.GrantSendOnBehalfTo = $grantSendStore -join ";"
            $obj.HiddenFromAddressListsEnabled = $dl.HiddenFromAddressListsEnabled
            $obj.LegacyExchangeDN = $dl.LegacyExchangeDN
            $obj.RecipientTypeDetails = $dl.RecipientTypeDetails
            $obj.Name = $dl.Name
            $obj.DistinguishedName = $dl.DistinguishedName
            $obj.ExchangeVersion = $dl.ExchangeVersion
            $obj.Guid = $dl.Guid.ToString()
            $obj.RequireSenderAuthenticationEnabled = $dl.RequireSenderAuthenticationEnabled
            $membersStore = @()
            $dgmVar = $null
            $dgmVar = Get-DistributionGroupMember $dl.Identity
            $dgmVar | Export-CliXml ".\GroupExports\$($dl.Alias)_Member_XmlExport_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).xml"
            foreach ($dlMember in $dgmVar) {
                if (IsValidEmail $dlMember.PrimarySMTPAddress) {
                    $membersStore += $dlMember.PrimarySMTPAddress
                } else {
                    $obj.Notes += "GrpMemSMTPNotFound:$($dlMember.Name). "
                }
            }
            $obj.Members = $membersStore -join ';'

            $drStore = @()
            foreach ($drUser in $dl.AcceptMessagesOnlyFromSendersOrMembers) {
                $drUserDetails = $null
                $drUserDetails = Get-Recipient $drUser -ea 0
                if ($drUserDetails.PrimarySMTPAddress) {
                    $drStore += $drUserDetails.PrimarySMTPAddress
                } Else {
                    $obj.Notes += "DRMemSMTPNotFound:$($drUser). "
                }
            }
            $obj.AcceptMessagesOnlyFromSendersOrMembers = $drStore -join ';'

            $moStore = @()
            foreach ($memberOfObject in (Get-ADPrincipalGroupMembership $dl.SamAccountName -ea 0)) {
                $dgObjectLookup = Get-DistributionGroup $memberOfObject.ObjectGuid.ToString() -ea 0
                if ($dgObjectLookup) {
                    $moObj = "" | select GroupPrimary,MemberOf,MemberOfGuid
                    $moObj.GroupPrimary = $dl.PrimarySMTPAddress
                    $moObj.MemberOf = $dgObjectLookup.PrimarySMTPAddress
                    $moObj.MemberOfGuid = $memberOfObject.ObjectGuid.ToString()
                    $moStore += $dgObjectLookup.PrimarySMTPAddress
                    $moObjCount = $moObj | measure | %{ $_.Count }
                    if ($moObjCount -ge 1) {
                        $moStoreExport += $moObj
                    }
                } Else {
                    $obj.Notes += "DGMemOfDGObjNotFound:$($memberOfObject.SamAccountName). "
                }
            }
            $obj.MemberOf = $moStore -join ';'
            $expandedExport += $obj
        }
        $expandedExport | Export-Csv ".\GroupExports\DistGroupExpanded_Export_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        $moStoreExport | Export-Csv ".\GroupExports\DistGroupMemberOf_Export_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        $ExportFile = @()
        if ($DistGroup) {
            Write-Host "Looking for group members.." -NoNewline
            $ExportFile += ExpandGroup -DistGroup $DistGroup -CheckNested $True
            <# foreach ($Group in $DistGroup)
            {
                
                $CurMems = Get-DistributionGroupMember -identity $Group.Identity -resultsize unlimited
                foreach ( $Member in $CurMems) {
                    $NewMember = New-Object PSObject -Property @{ DistGrpName = $Group.DisplayName ;
                    DistGrpAlias = $Group.Alias ; 
                    DistGrpPrimarySMTP = $Group.PrimarySmtpAddress ;
                    GrpMemberPrimarySMTP = $Member.PrimarySmtpAddress ; 
                    GrpMemberAlias = $Member.Alias ; 
                    GrpMemberRecTypeDetail = $Member.RecipientTypeDetails ; 
                    GrpMemberRecType = $Member.RecipientType ; 
                    GrpMemberDisplayName = $Member.DisplayName ; 
                    GrpMemberSamAccName = $Member.SamAccountName }
                    $ExportFile += $NewMember
                }
            } #>
            Write-Host "Done" -ForegroundColor Green
                
            $ExportFile | Export-Csv ".\GroupExports\GroupMembers_Export_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            $ExportFile | sort DistGrpPrimarySMTP -Unique | select DistGrpPrimarySMTP | Export-Csv ".\GroupExports\AllNestedGrps_Export_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            $nestedAllList = $ExportFile | sort DistGrpPrimarySMTP -Unique | select @{n='PrimarySMTPAddress';e={$_.DistGrpPrimarySMTP}}
            $comparisonResult = Compare-Object -ReferenceObject ($DistGroup | Select PrimarySMTPAddress) -DifferenceObject $nestedAllList -PassThru | select -expand PrimarySMTPAddress
            if ($comparisonResult) {
                "Nested DL List has additional DLs that are not included in this migration batch:"
                $comparisonResult
            }
        }

        $allowMembers = @()
        write-host "Looking for message delivery restrictions for groups.." -NoNewline
        foreach ($grp in $DistGroup) {
            if ($grp.AcceptMessagesOnlyFromSendersOrMembers -ne $null) {
                foreach ($allowsoMem in $grp.AcceptMessagesOnlyFromSendersOrMembers) {
                    $obj2 = "" | Select-Object GrpDisplayName,GrpPrimarySMTPAddress,AllowedMember,AllowedMemberPrimarySMTP
                    $obj2.GrpDisplayName = $grp.DisplayName
                    $obj2.GrpPrimarySMTPAddress = $grp.PrimarySMTPAddress
                    $obj2.AllowedMember = $allowsoMem
                    $obj2.AllowedMemberPrimarySMTP = Get-Recipient $allowsoMem -ea 0 | select -ExpandProperty PrimarySMTPAddress
                    $allowMembers += $obj2
                }
            }
        }
        if ($allowMembers) {
            $allowMembers | Export-CSV ".\GroupExports\DistGroupDR_Export_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            write-host "Done" -ForegroundColor Green
        } else { write-host "No restrictions found" -ForegroundColor Green }
        write-host "Done" -foregroundcolor green
        ""
    }

    if ($reply -eq '2') {
        ""
        $expandedName = Get-ChildItem ".\GroupExports\" | where{$_.Name -match "DistGroupExpanded_Export"}
        $expandedNameCount = $expandedName | measure | %{ $_.Count }
        if ($expandedNameCount -eq 1) {
            "Importing CSV file: $($expandedName.FullName)"
            $groups = Import-Csv $expandedName.FullName -ea 1
        } Else {
            "Error: Missing or multiple expanded csv files in GroupExports directory."
            Continue
        }
        "Group count found in expanded csv file: $($groups | measure | %{ $_.Count })"
        "The following DLs will be deleted:"
        $groups | ft PrimarySMTPAddress,Name,OrganizationalUnit -Autosize
        ""
        $ans = Read-Host "Proceed? (y/n)"
        if ($ans -ieq 'y') {
            $groups | %{ "Removing DL $($_.PrimarySMTPAddress)" ; Remove-DistributionGroup $_.PrimarySMTPAddress -Confirm:$true -BypassSecurityGroupManagerCheck -ea 1 }
        } Else {
            "Aborting.."
        }        
        write-host "Done" -foregroundcolor green
        ""
    }

    if ($reply -eq '3') {
        ""
        $expandedName = Get-ChildItem ".\GroupExports\" | where{$_.Name -match "DistGroupExpanded_Export"}
        $expandedNameCount = $expandedName | measure | %{ $_.Count }
        if ($expandedNameCount -eq 1) {
            "Importing CSV file: $($expandedName.FullName)"
            $groups = Import-Csv $expandedName.FullName -ea 1
        } Else {
            "Error: Missing or multiple expanded csv files in GroupExports directory."
            Continue
        }
        $groupCount = $groups | measure | %{ $_.Count }
        "Group count loaded through CSV search: $($groupCount)"

        foreach ($con in $groups) {
            $checkExists = Get-Recipient $con.PrimarySMTPAddress -ea 0
            if ($checkExists) {
                write-host "Error: Recipient exists in AD for $($con.PrimarySMTPAddress) with recipient type of $($checkExists.RecipientTypeDetails)" -foregroundcolor red
                Continue
            }
            write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Creating new mail contact for $($con.PrimarySMTPAddress)"
            New-MailContact -Name $con.Name -DisplayName $con.DisplayName -PrimarySMTPAddress $con.PrimarySMTPAddress -ExternalEmailAddress $con.RoutingAddress -OrganizationalUnit 'OU=BadAccounts,DC=tenethealth,DC=net' -ea 1 | out-null
            write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up proxyAddresses for $($con.PrimarySMTPAddress)"
            foreach ($proxy in ($con.EmailAddresses -split '\|')) {
                Try {
                    if ($proxy -cmatch "SMTP") {
                        Set-MailContact $con.PrimarySMTPAddress -EmailAddresses @{Add="$($proxy.ToLower())"} -ea 1    
                    } Else {
                        Set-MailContact $con.PrimarySMTPAddress -EmailAddresses @{Add="$($proxy)"} -ea 1
                    }
                } Catch {
                    write-host "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : ProxyError($($error[0].InvocationInfo.ScriptLineNumber)): $($con.PrimarySMTPAddress) : $($error[0].Exception.Message)" -ForegroundColor Red
                }
            }
            write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : Setting up x500 for $($con.PrimarySMTPAddress)"
            Set-MailContact $con.PrimarySMTPAddress -EmailAddresses @{Add="x500:$($con.LegacyExchangeDN)"} -ea 1
        }
        write-host "Done" -foregroundcolor green
        ""
    }

    if ($reply -eq '4') {
        ""
        $expandedName = Get-ChildItem ".\GroupExports\" | where{$_.Name -match "DistGroupExpanded_Export"}
        $expandedNameCount = $expandedName | measure | %{ $_.Count }
        if ($expandedNameCount -eq 1) {
            "Importing CSV file: $($expandedName.FullName)"
            $groups = Import-Csv $expandedName.FullName -ea 1
        } Else {
            "Error: Missing or multiple expanded csv files in GroupExports directory."
            Continue
        }
        $groupCount = $groups | measure | %{ $_.Count }
        "Group count loaded from CSV: $($groupCount)"
        $successfulAdds = 0
        $failedAdds = 0

        foreach ($con in $groups) {
            write-host "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : Processing MemberOf items for $($con.PrimarySMTPAddress)" -ForegroundColor Green
            foreach ($dlObject in ($con.MemberOf -split ';')) {
                $checkDLExists = $null
                if ($dlObject) {
                    $checkDLExists = Get-Recipient $dlObject -ea 0
                    if ($checkDLExists) {
                        if ($checkDLExists.RecipientType -match 'group') {
                            Try {
                                Add-DistributionGroupMember $dlObject -Member $con.PrimarySMTPAddress -ea 1
                                $successfulAdds += 1
                            } Catch {
                                write-host "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : DLMemOfError($($error[0].InvocationInfo.ScriptLineNumber)): $($con.PrimarySMTPAddress) : $($error[0].Exception.Message)" -ForegroundColor Red
                                $failedAdds += 1
                            }
                        } Else {
                            write-host "$(Get-Date -UFormat %y%m%d_%H%M%S%Z) : DLMemOfRecTypeError($($con.PrimarySMTPAddress)): Non-Group Recipient exists in AD for $($dlObject) with recipient type of $($checkDLExists.RecipientTypeDetails)" -foregroundcolor red
                        }
                    } Else {
                        "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : DLMemOfDLNotFoundError($($con.PrimarySMTPAddress)) : MemberOf item is not a valid recipient or does not exists"
                    }
                } Else { 
                    "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : DLMemOfNullError($($con.PrimarySMTPAddress)) : MemberOf item is empty or null"
                }
            }
        }
        write-host "Done" -foregroundcolor green
        ""
        "Successful additions: $($successfulAdds)"
        "Failed additions:     $($failedAdds)"
        ""
    }

    if ($reply -eq '5') {
        ""
        $expandedName = Get-ChildItem ".\GroupExports\" | where{$_.Name -match "DistGroupExpanded_Export"}
        $expandedNameCount = $expandedName | measure | %{ $_.Count }
        if ($expandedNameCount -eq 1) {
            "Importing CSV file: $($expandedName.FullName)"
            $groups = Import-Csv $expandedName.FullName -ea 1
        } Else {
            "Error: Missing or multiple expanded csv files in GroupExports directory."
            Continue
        }
        $groupCount = $groups | measure | %{ $_.Count }
        "Group count loaded from CSV: $($groupCount)"

        $moExport = @()

        foreach ($con in $groups) {
            #write-host "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : Getting MemberOf items for $($con.PrimarySMTPAddress)" -ForegroundColor Green
            $dlConObject = Get-MailContact $con.PrimarySMTPAddress -ea 0
            $adObj = Get-ADObject -Identity $dlConObject.DistinguishedName -Properties memberOf -ea 0
            if ($adObj) {
                $moCount = 0 ; $moCount = $adObj.memberOf | measure | %{ $_.Count }
                if ($moCount -ne 0) {
                    "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : Number of MemberOf items found for $($con.PrimarySMTPAddress): $($moCount)"
                    foreach ($moItem in $adObj.memberOf) {
                        $moRec = Get-Recipient $moItem -ea 0
                        $moObj = "" | select DLContact,MemberOf
                        $moObj.DLContact = $con.PrimarySMTPAddress
                        $moObj.MemberOf = $moRec.PrimarySMTPAddress
                        $moExport += $moObj
                    }
                } Else {
                    "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : No MemberOf items found for $($con.PrimarySMTPAddress)"
                }
            } Else {
                "$(Get-Date -UFormat %y%m%d_%H%M%S_%Z) : AD object for mail contact $($con.PrimarySMTPAddress) was not found"
            }
        }
        $moExport | Export-CSV ".\GroupExports\DLConMemOfCheck_$(Get-Date -UFormat %m%d%y_%H%M%S%Z).csv" -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        write-host "Done" -foregroundcolor green
        ""
    }

} while ($true)

Stop-Transcript
