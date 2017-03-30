function Search-MailFrom {
    <#
    .SYNOPSIS
        Search a mailbox by "from address"
    .DESCRIPTION
        Use a search string with a trailing wildcard(*) to filter the mailbox
    #>
    Param(
        [string]$Identity
        , # Cannot start with a wildcard, only end with *
        [string]$From
        , [datetime]$Start = (Get-date).Date
        , [datetime]$End = $Start.addDays(1).Date
        , [switch]$Delete
        , [string]$ResultTarget = $env:username
    )
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd')) AND from:`"$from`"" -TargetMailBox $ResultTarget -TargetFolder "Search" -DeleteContent:$Delete
}

function Search-MailSubject {
    <#
    .SYNOPSIS
        Search a mailbox by subject
    .DESCRIPTION
        Use a search string with a trailing wildcard(*) to filter the mailbox
    #>
    Param(
        [string]$Identity
        , # Cannot start with a wildcard, only end with *
        [string]$Subject
        , [datetime]$Start = (Get-date).Date
        , [datetime]$End = $Start.addDays(1).Date
        , [switch]$Delete
        , [string]$ResultTarget = $env:username
    )
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd')) AND subject:`"$Subject`"" -TargetMailBox $ResultTarget -TargetFolder "Search" -DeleteContent:$Delete
}

function Search-MailDate {
    <#
    .SYNOPSIS
        Search a users mailbox by day
    .EXAMPLE
        Search-MailDate -Identity <searchTarget>

        RunspaceId       : xxxxxxxx-0000-1111-aaaa-1234abcd5678
        Identity         : Domain/Name/Users/Group/searchTarget
        DisplayName      : searchTarget
        TargetMailbox    : Domain/Name/Users/Group/resultTarget
        TargetPSTFile    :
        Success          : True
        TargetFolder     : \Search\searchTarget-19/10/2016 15:03:48
        ResultItemsCount : 29
        ResultItemsSize  : 18.83 MB (19,742,786 bytes)
    #>
    Param(
        [string]$Identity
        , [datetime]$Start = (Get-date).Date
        , [datetime]$End = $Start.addDays(1).Date
        , [switch]$Delete
        , [string]$ResultTarget = $env:username
    )
    "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd'))"
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd'))" -TargetMailBox $ResultTarget -TargetFolder "Search" -DeleteContent:$Delete
}

$script:session
function Import-MailServer {
    <#
    .SYNOPSIS
        Create the remote connection to the mail server as yourself.
    .DESCRIPTION
        Correctly open remote session to the exchange server which loads the exchange cmdlets.

        When you are finished remove the connection with Remove-MailServer
    .EXAMPLE
        Import-MailServer
        WARNING: The names of some imported commands from the module 'tmp_connection.hlk' include unapproved verbs that might
        make them less discoverable. To find the commands with unapproved verbs, run the Import-Module command again with the
        Verbose parameter. For a list of approved verbs, type Get-Verb.

        ModuleType Version    Name                                ExportedCommands
        ---------- -------    ----                                ----------------
        Script     1.0        tmp_connection.hlk                    {Add-ADPermission, Add-AvailabilityAddressSpace, Add-Conte...
        WARNING: Use 'Remove-Mailserver' to close the connection for other users to get on.
    #>
    Param(
        [string]
        $ComputerName = 'mailgate.birkdalehigh.co.uk'
    )
    $test = test-connection $ComputerName -count 1
    $host = [System.Net.Dns]::GetHostbyAddress($test.ProtocolAddress).HostName

    $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$host/PowerShell/" -Authentication Kerberos
    Import-PSSession $script:session
    Write-Warning "Use 'Remove-Mailserver' to close the connection for other users to get on."

}
function Get-MailServer {
    <#
    .SYNOPSIS
        Get the mailserver session
    .DESCRIPTION
        Use the existing session to import into to higher powershell scope to get access to all the exchange cmdlets
    .EXAMPLE
        Import-Mailserver; Import-PSSession (Get-MailServer)
    #>
    return $script:session
}
function Remove-MailServer {
    <#
    .SYNOPSIS
        Removes the imported remote connection to the mail server.
    .EXAMPLE
        Remove-MailServer
    #>
    Remove-PSSession $script:session
}

function Get-RecentFailedMessage {
    <#
    .SYNOPSIS
        Show failed messages sent to the exchange server in the last hour
    .DESCRIPTION
        Filter the message tracking log for failed messages not caught by Sophos i.e PmE12Transport.

        This filtering should be checked regularly as messages failing here will not show up anywhere else.
    .EXAMPLE
        Get-RecentFailedMessage -View

        Timestamp           EventId Source Sender            Recipients          MessageSubject
        ---------           ------- ------ ------            ----------          --------------
        07/02/2017 11:17:42 FAIL    SMTP   Sent@example.com  {demo@example.com}  Re: Example
    .EXAMPLE
        Get-RecentFailedMessage | select source,recipientstatus,sender

        Source RecipientStatus                           Sender           ClientIP
        ------ ---------------                           ------           ---------
        SMTP   {550 5.7.1 Sender ID (PRA) Not Permitted} sent@example.com 127.0.0.1
    .NOTES
        550 5.7.1 Sender ID (PRA) Not Permitted = Incorrect SPF record from domain and clientIP
        To fix this you need to exclude the domain by appending to the SenderIdConfig BypassedSenderDomains list
        Set-SenderIdConfig -BypassedSenderDomains ( (Get-SenderIdConfig).BypassedSenderDomains += "example.com" )
        Polite message the sending domain technical support their DNS is configured incorrectly.
    #>
    Param(
        # Friendly view that cannot be consumed
        [switch]
        $View
    )
    Process{
        $result = Get-MessageTrackingLog -Start (get-date).addMinutes(-60) | Where-Object {($_.eventid -eq 'fail') -and ($_.sourceContext -ne 'PmE12Transport')}
        if($view){
            Write-Output $result | format-table -AutoSize -property timestamp,source,sender,recipients,messagesubject
        } else {
            Write-Output $result
        }
    }
}

function Convert-DistributionGroupToSharedMailbox {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Recrete a group as a shared mailbox
    .DESCRIPTION
        A distribution gounp is a kind of address that exchanged cannot transform into a different type.

        This cmdlet takes a group, stores the address exchange uses to map the address lookup, removes the group and re-creates it as a shared mailbox. Then re-adds the address mapping.
    .EXAMPLE
        PS C:\> Get-DistributionGroup Finance | Convert-DistributionGroupToSharedMailbox
    .INPUTS
        Microsoft.Exchange.Data.Directory.Management.DistributionGroup
    .OUTPUTS
        Microsoft.Exchange.Data.Directory.Management.Mailbox
    .NOTES
        Order of operations.
        todo: test for exsiting distribution group
        todo: Check group membership to re-apply permissions on the mailbox
        Save LegacyExchangeDN attribute for outlook address book mappings
        Remove existing distribution group
        Create new shared mailbox in correct DN
        Add X500 of LegacyExchangeDN of distrubution group that was replaced
    #>
    Param(
        # Pipe Get-DistributionGroup to me.
        [Parameter(Mandatory)]
        [Microsoft.Exchange.Data.Directory.Management.DistributionGroup]
        $Group
    )
    Begin{
        Throw "Never tested. Not even once. Validate for yourself."
    }
    Process{
        $name = $Group.Name
        $Mapping = $Group.LegacyExchangeDN
        Remove-DistributionGroup -Identity $Name
        New-SharedMailbox -DisplayName $Name -Alias $Name | Set-Mailbox -EmailAddresses "X500:$Mapping"
    }
}
function New-SharedMailbox {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Create a Shared Mailbox with a default configuration
    .DESCRIPTION

    .EXAMPLE
        PS C:\> New-SharedMailbox -Name "AddressTitle" -DisplayName "Full Address Title" -Alias "address" -Users "personA", "PersonB"
        Name           Alias
        ----           -----
        AddressTitle   address
    .OUTPUTS
        Microsoft.Exchange.Data.Directory.Management.Mailbox
    .NOTES
        Order of operations;
        new shared mailbox in correct DN
        Assign mailbox send parameters
        Add persmissions for each user
            Check to add full send-from persmission from switch
        Add send on behalf permission
        Return mailbox
    #>
    Param(
        # Will show as the name for the contact and mailbox
        [Parameter(Mandatory)]
        [string]
        $Name

        , # The name that appears in the Exchange Management Console under Recipient Configuration
        [string]
        $DisplayName

        , #Emaill address excluding @example.com
        [Parameter(Mandatory)]
        [alias('EmailAddress')]
        [ValidatePattern('.*[^@].*')]
        [string]
        $Alias

        , # Users that initially use the mailbox
        [string[]]
        $Users

        , # Path to shared mailbox user account, not an AD OU Path
        [string]
        [ValidatePattern('[^=;]')]
        $OrganizationalUnit = "bhs.internal/BHS/Mail Users"

        , # mailbox database storage location
        [string]
        $Database = "BHS Staff Mailbox Database"

        , # Users will get permission to send "<user> On behald of <account>", directly from <account> or not at all.
        [ValidateSet('None', 'Behalf', 'From')]
        [string]
        $Send = 'Behalf'
    )
    Begin{
    }
    Process{
        New-MailBox -Name $Name -DisplayName $DisplayName -Alias $alias -OrganizationalUnit $OrganizationalUnit -Database $Database -UserPrincipalName "$Alias@bhs.internal" -Shared
        # By default sent items will only show in the senders account. This forces them into the shared sent items folder.
        Set-MailboxSentItemsConfiguration -Identity $Name -SendAsItemsCopiedTo 'SenderAndFrom' -SendOnBehalfOfItemsCopiedTo 'SenderAndFrom'
        $users | ForEach-Object {
            # Mailbox permissions allow the user to perform actions
            Add-MailBoxPermission $Name -AccessRights FullAccess -InheritanceType All -User $PSItem
            # AD permissions allow the users account (outlook) to find the mailbox
            Add-ADPermission $Name -ExtendedRights "Receive-As" -User $PSItem
            if($Send -eq 'From'){
                Add-ADPermission $Name -ExtendedRights "Send-As" -User $PSItem
            }
        }
        if($Send -eq 'Behalf'){
            # Send 'From' permissions will supersede this "On behalf" setting.
            # -GrantSendOnBehalfTo only replaces the list, you must append your own users list first.
            Set-Mailbox -Identity $Name -GrantSendOnBehalfTo $users
            Write-Warning "The permission for users to send 'From' will supersede this permission, check their AD-Permissions for 'Send-As'"
        }
    }
    End{
        Get-Mailbox -Identity $Name
    }
}

Export-ModuleMember -Function "Get-*", "Search-*", "New-*", "*-MailServer"
