function Search-MailFrom {
    <#
    .SYNOPSIS
        Search a mailbox by from address
    .DESCRIPTION
        Use a search string with a trailing wildcard to filter the mailbox
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
        Use a search string with a trailing wildcard to filter the mailbox
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
        Search-MailDate -Identity sdolan

        RunspaceId       : 90038754-0fdd-4524-ad84-7c11cacfb8a8
        Identity         : BHS.INTERNAL/BHS/Users/System Administrators/Mr S Dolan
        DisplayName      : Mr S Dolan
        TargetMailbox    : BHS.INTERNAL/BHS/Users/System Administrators/Mr J. Bennett
        TargetPSTFile    :
        Success          : True
        TargetFolder     : \Search\Mr S Dolan-19/10/2016 15:03:48
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
        Corectly enters a remote session to the exchange server and loads the requisite cmdlets.

        When you are finished remove the connection with Remove-MailServer
    .EXAMPLE
        Import-MailServer
        WARNING: The names of some imported commands from the module 'tmp_sy1nupvn.hlk' include unapproved verbs that might
        make them less discoverable. To find the commands with unapproved verbs, run the Import-Module command again with the
        Verbose parameter. For a list of approved verbs, type Get-Verb.

        ModuleType Version    Name                                ExportedCommands
        ---------- -------    ----                                ----------------
        Script     1.0        tmp_sy1nupvn.hlk                    {Add-ADPermission, Add-AvailabilityAddressSpace, Add-Conte...
        WARNING: Use 'Remove-Mailserver' to close the connection for other users to get on.
    #>
    $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://bhs-mail01.bhs.internal/PowerShell/ -Authentication Kerberos
    Import-PSSession $script:session
    Write-Warning "Use 'Remove-Mailserver' to close the connection for other users to get on."

}
function Get-MailServer {
    <#
    .SYNOPSIS
        Get the mailserver session
    .EXAMPLE
        Import-Mailserver; Import-PSSession (Get-MailServer)
    #>
    return $script:session
}
function Remove-Mailserver {
    <#
    .SYNOPSIS
        Removes the imported remote connection to the mail server..
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
        A distribution gounp is a kind of address that exchanged cannot trasnform into a different type.

        This cmdlet takes a group, stores the address exchange uses to map the address lookup, removes the group and re-creates it as a shared mailbox. Then re-adds the address mapping.
    .EXAMPLE
        PS C:\> Get-DistributionGroup Finance | Convert-DistributionGroupToSharedMailbox
    .INPUTS
        Microsoft.Exchange.Data.Directory.Management.DistributionGroup
    .OUTPUTS
        Microsoft.Exchange.Data.Directory.Management.Mailbox
    .NOTES
        Order of operations.
        todo: test for exsiting distrubtion group
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
        PS C:\> <example usage>
        Explanation of what the example does
    .OUTPUTS
        Microsoft.Exchange.Data.Directory.Management.Mailbox
    .NOTES
        Order of operations.
        todo: test for exsiting distrubtion group
        remove existing distribution group
        new shared mailbox in correct DN
        add X500 of legacyEchangeDN of distrubution group that was replaced for <reason>

    #>
    Param(
        #Will show as the name for the contact and mailbox
        [string]
        $DisplayName

        , #Emaill address excluding @example.com
        [string]
        $Alias

        , # Users that initially use the mailbox
        [string[]]
        $Users

        ,# Org unit path for shared mailboxes
        [string]
        $OrganizationalUnit = "bhs.internal/BHS/Mail Users"

        , # mailbox database storage location
        [string]
        $Database = "BHS Staff Mailbox Database"

        , # Send mail as the email account itself as opposed to "<user> On behald of <account>"
        [switch]
        $SendAs
    )
    Begin{
        Throw "Never tested. Not even once. Validate for yourself."
    }
    Process{
        New-MailBox -Identity $DisplayName -Alias $alias -OrganizationalUnit $OrgUnit -Database $Database -UserPrincipalName "$Alias@bhs.internal" -Shared
        # By default sent items will only show in the senders account. This forces them into the shared sent items folder.
        Set-MailboxSentItemsConfiguration -Identity $DisplayName -SendAsItemsCopiedTo 'SenderAndFrom' -SendOnBehalfOfItemsCopiedTo 'SenderAndFrom'
        $users | ForEach-Object {
            # Mailbox permissions allow the user to perform actions
            Add-MailBoxPermission $DisplayName -AccessRights FullAccess -InheritanceType All -User $PSItem
            # AD permissions allow the users account (outlook) to find the mailbox
            Add-ADPermission $DisplayName -ExtendedRights "Receive-As" -User $PSItem
            if($SendAs){
                Add-ADPermission $DisplayName -ExtendedRights "Send-As" -User $PSItem
            }
        }
        # SendAs permissions will supersede this "On behalf" setting.
        Set-Mailbox -Identity $DisplayName -GrantSendOnBehalfTo $users
    }
    End{
        Get-Mailbox -Identity $DisplayName
    }
}

Export-ModuleMember -Function "Get-*", "Search-*", "*-MailServer"
