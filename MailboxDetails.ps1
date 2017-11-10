Function Get-MailboxDetails {
[CmdletBinding()]
param(
	[parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
	[string]$Name
)
process {

#variables
$tracefile = ".\Function_GetMailboxDetails_TraceFile_$(get-date -format 'ddMMyyyy-HHmm').txt"
Start-Transcript -path $tracefile
$sessionlogfile = ".\Function_GetMailboxDetails_SessionLog_$(get-date -format 'ddMMyyyy-HHmm').txt"
$Global:ExportName = $Name

$Mbxes = Get-Mailbox -resultsize unlimited
[int]$validateinfototal = $Mbxes.Count
[int]$validateinfoprogress = 0

$allObjects = @()

 $Mbxes | %{
        $validateinfoprogress++
        Write-Progress -Activity "Checking mailbox, $validateinfoprogress of $validateinfototal..." -PercentComplete (($validateinfoprogress/$validateinfototal)*100)

        #Retrieve Info
        $email= $_.PrimarySMTPAddress.ToString()
        $userdetails =  Get-ADUser $_.SamAccountName -Properties co,GivenName,Office,roomNumber,department
        If ($_.UMEnabled -eq "True")
            {
            $mbxumdetails = Get-UMMailbox $_.SamAccountName
            }
        $Stats = Get-MailboxStatistics $email -Erroraction SilentlyContinue
        
        If ($userdetails.co -ne $null)
            {
            $Country = $userdetails.co.ToString()
            #ASML Country translation below
            If (($Country -eq "The Netherlands") -or ($Country -eq "Netherlands"))
                { $Country = "NL" }
            If (($Country -eq "USA") -or ($Country -eq "United States"))
                { $Country = "US" }
            If ($Country -eq "China")
                { $Country = "CN" }
            If ($Country -eq "Taiwan")
                { $Country = "TW" }
            If ($Country -eq "Germany")
                { $Country = "DE" }
            If ($Country -eq "Korea")
                { $Country = "KR" }
            If ($Country -eq "France")
                { $Country = "FR" }
            If ($Country -eq "Italy")
                { $Country = "IT" }
            If ($Country -eq "Belgium")
                { $Country = "BE" }
            If ($Country -eq "Japan")
                { $Country = "JP" }
            If ($Country -eq "Singapore")
                { $Country = "SG" }
            If ($Country -eq "Malaysia")
                { $Country = "MY" }
            If ($Country -eq "Ireland")
                { $Country = "IE" }
            If ($Country -eq "Israel")
                { $Country = "IL" }
            If ($Country -eq "United Kingdom")
                { $Country = "GB" }
            If ($Country -eq "Poland")
                { $Country = "PL" }
            }
        If ($userdetails.sn -ne $null)
            {
            $LastName = $userdetails.sn.ToString()
            }
        If ($userdetails.GivenName -ne $null)
            {
            $FirstName = $userdetails.GivenName.ToString()
            }
        If ($userdetails.Office -ne $null)
            {
            $Office = $userdetails.Office.ToString()
            }
        If ($userdetails.roomNumber -ne $null)
            {
            $RoomNumber = ($userdetails.roomNumber -join "^").ToString()
            }
        If ($userdetails.department -ne $null)
            {
            $Department = $userdetails.department.ToString()
            }
        If ($_.RecipientTypeDetails -eq "UserMailbox")
            { $MbxType = "USER" }
        If ($_.RecipientTypeDetails -eq "SharedMailbox")
            { $MbxType = "SHARED" }
        If ($_.RecipientTypeDetails -eq "RoomMailbox")
            { $MbxType = "ROOM" }
        If ($_.RecipientTypeDetails -eq "EquipmentMailbox")
            { $MbxType = "EQUIPMENT" }

        #Create the object
        $MailboxStats = New-Object -TypeName PSObject
        $MailboxStats | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $email
        $MailboxStats | Add-Member -MemberType NoteProperty -Name Location -Value $Country
        $MailboxStats | Add-Member -MemberType NoteProperty -Name Lastname -Value $LastName
        $MailboxStats | Add-Member -MemberType NoteProperty -Name Firstname -Value $FirstName
        $MailboxStats | Add-Member -MemberType NoteProperty -Name DisplayName -Value $($stats.DisplayName)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name Office -Value $Office
        $MailboxStats | Add-Member -MemberType NoteProperty -Name RoomNumber -Value $RoomNumber
        $MailboxStats | Add-Member -MemberType NoteProperty -Name Department -Value $Department
        $MailboxStats | Add-Member -MemberType NoteProperty -Name Type -Value $MbxType
        #$MailboxStats | Add-Member -MemberType NoteProperty -Name LicenseType -Value $($_.CustomAttribute13)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name UPN -Value $($_.UserPrincipalName)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name SamAccountName -Value $($_.SamAccountName)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name TotalItemSizeMB -Value $([math]::Round( ($stats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2))
        $MailboxStats | Add-Member -MemberType NoteProperty -Name FolderCount -Value $((get-mailboxfolderstatistics -id $_.Identity | select FolderId).count)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name TotalItemSize -Value $($stats.TotalItemSize)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name ItemCount -Value $($stats.ItemCount)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name LegacyExchangeDN -Value $($_.LegacyExchangeDN)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name ProxyAddresses -Value $($_.EmailAddresses -join "^")
        $MailboxStats | Add-Member -MemberType NoteProperty -Name LitigationHoldEnabled -Value $($_.LitigationHoldEnabled)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name LitigationHoldDate -Value $($_.LitigationHoldDate)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name LitigationHoldOwner -Value $($_.LitigationHoldOwner)
        $MailboxStats | Add-Member -MemberType NoteProperty -Name UMEnabled -Value $($_.UMEnabled)
        If ($_.UMEnabled -eq "True")
            {
            $MailboxStats | Add-Member -MemberType NoteProperty -Name UMDialPlan -Value $($mbxumdetails.UMDialPlan)
            $MailboxStats | Add-Member -MemberType NoteProperty -Name UMMailboxPolicy -Value $($mbxumdetails.UMMailboxPolicy)
            $MailboxStats | Add-Member -MemberType NoteProperty -Name Extensions -Value $($mbxumdetails.Extensions)
            }

        $allObjects += $MailboxStats
        }

[string]$MailboxDetailsFile = $exportpath + $Global:ExportName + "-MailboxDetails.csv"

If ($Global:ExportName -eq $null)
    {
    $MailboxDetailsFile = $exportpath + "MailboxDetailsExport.csv"
    }

If ((Test-Path $MailboxDetailsFile) -eq $false)
    {
    write-Host ("Exportfile not yet created, creating mailbox details exportfile: " + $MailboxDetailsFile) -Fore Yellow
    New-Item $MailboxDetailsFile -ItemType File | Out-Null
    }

$allObjects | Export-Csv $MailboxDetailsFile -NoTypeInformation -Delimiter ";"

   }
}