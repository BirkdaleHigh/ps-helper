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
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd')) AND from:`"$from`"" -TargetMailBox $ResultTarget -TargetFolder "Search" -DeleteContent:$Delete -Force:$Delete
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
[bool]$script:import
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
        $ComputerName = 'mailgate.birkdalehigh.co.uk',

        [PSCredential]
        $Credential,

        [ValidateSet('Script','Global')]
        $Scope = 'Global'
    )
    $test = Test-Connection $ComputerName -count 1 -ErrorAction Stop
    $host = [System.Net.Dns]::GetHostbyAddress($test.ProtocolAddress).HostName

    $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$host/PowerShell/" -Authentication Kerberos -Credential:$Credential
    if($Scope -eq 'Global'){
        Import-module (Import-PSSession $script:session) -Global
    } else {
        if(-not $script:import){
            Import-PSSession $script:session
            $script:import = $true
        }
    }
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
    if ($script:session){
        return $script:session
    }
    else{
        Write-Error 'Cannot find open session, use Import-Mailserver'
    }
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

        This cmdlet wraps Get-MessageTrackingLog and requires the exchange cmdlets to be loaded. see `Import-MailServer`
    .EXAMPLE
        Get-RecentFailedMessage -FormatView
        Auto format the output to be easily read. Data not suitable for piping.

        Timestamp            Sender            Recipients          MessageSubject
        ---------            ------            ----------          --------------
        07/02/2017 11:17:42  Sent@example.com  {demo@example.com}  Re: Example
    .EXAMPLE
        Get-RecentFailedMessage -FormatView -IncludeSophos
        Auto format the output to be easily read. Includes the spam filter that normally has its own console.

        Timestamp            SourceContext    Sender             Recipients              MessageSubject
        ---------            -------------    ------             ----------              --------------
        29/09/2017 11:33:40  PmE12Transport   ABC@example.com    {internal@example.com}  Invoice
        29/09/2017 11:33:50  Sender Id Agent  DEF@example.com    {internal@example.com}  RE: Quick Quest

        Data not suitable for piping.
    .EXAMPLE
        Get-RecentFailedMessage | select SourceContext,recipientstatus,sender,ClientIP
        Make a table of any column headers, do not use a format switch as if cannot be used down the pipeline.

        SourceContext     RecipientStatus                           Sender           ClientIP
        -------------     ---------------                           ------           ---------
        Sender Id Agent   {550 5.7.1 Sender ID (PRA) Not Permitted} sent@example.com 127.0.0.1
    .EXAMPLE
        Get-RecentFailedMessage -IncludeSophos | measure
        Check the number of blocked messages in the last hour

        Count    : 231
    .EXAMPLE
        Get-RecentFailedMessage -recipients userA@example.com
        Gets messages that have faild in the last hour for a specific user account.

        EventId  Source   Sender                   Recipients           MessageSubject
        -------  ------   ------                   ----------           --------------
        FAIL     SMTP     SpamSender@example.com   {userA@example.com}  Invoice
    .EXAMPLE
        Get-RecentFailedMessage | select -first 1  | format-list *
        Show all the property : values of one message object.

        PSComputerName          : org-server.org.internal
        RunspaceId              : 4e7dfba1-dd6a-4076-84f2-9ed78b888854
        PSShowComputerName      : False
        Timestamp               : 29/09/2017 11:38:05
        ClientIp                : 222.127.163.110
        ClientHostname          : ORG-SERVER
        ServerIp                : 10.201.0.25
        ServerHostname          :
        SourceContext           : Sender Id Agent
        ConnectorId             : ORG-SERVER\Default ORG-SERVER
        Source                  : SMTP
        EventId                 : FAIL
        InternalMessageId       : 0
        MessageId               : <2a228e10-9af9-83bc-aa16-7b526ca3e832@example.com>
        Recipients              : {UserD@example.com}
        RecipientStatus         : {550 5.7.1 Sender ID (PRA) Domain Does Not Exist}
        TotalBytes              : 0
        RecipientCount          : 1
        RelatedRecipientAddress :
        Reference               :
        MessageSubject          : Invoice
        Sender                  : Spammer@example.com
        ReturnPath              : Spammer@example.com
        MessageInfo             :
        MessageLatency          :
        MessageLatencyType      : None
        EventData               :3

        Inspect every propery of a message. Usefull to find property names you might want to make a table off.
        i.e. Get-RecentFailedMessage | select ClientIp,RecipientCount,Sender,ReturnPath
    .NOTES
        SourceContext as PmE12Transport is the Sophos spam filter engine.
        SourceContext as Sender Id Agent is an exchange filter agent.
        RecipientStatus as "550 5.7.1 Sender ID (PRA) Not Permitted" is an Incorrect SPF record for sender domain and clientIP
        To fix this you need to exclude the domain by appending to the SenderIdConfig BypassedSenderDomains list
        Set-SenderIdConfig -BypassedSenderDomains ( (Get-SenderIdConfig).BypassedSenderDomains += "example.com" )
        Politely message the sending domain technical support their DNS is configured incorrectly.
    #>
    Param(
        # Past amound of hours to get results
        [parameter(position = 0)]
        [int]
        $PastHours = 1

        , # Filter for a specific recipient
        [parameter(position = 1)]
        [ValidatePattern('(?# User account includes domain suffix)^.*@.*$')]
        [Alias('Identity')]
        [string[]]
        $recipients

        , # Friendly view that cannot be consumed down the pipe
        [switch]
        $FormatView

        , # Include mail sent to the sophos agent
        [switch]
        $IncludeSophos
    )
    Process{
        $result = Get-MessageTrackingLog -Start (get-date).addHours(-$PastHours) -EventId "FAIL" -Recipients:$recipients
        if( -not $IncludeSophos){
            $filtered = $result | Where-Object sourceContext -ne 'PmE12Transport'
        }
        if($FormatView -and $IncludeSophos){
            Write-Output $result | format-table -AutoSize -property timestamp,sourceContext,sender,recipients,messagesubject
        } elseif ($FormatView){
            Write-Output $filtered | format-table -AutoSize -property timestamp,sender,recipients,messagesubject
        } elseif ($IncludeSophos){
            Write-Output $result
        } else {
            Write-Output $filtered
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
    <#
    .SYNOPSIS
        Create a Shared Mailbox with a default configuration
    .DESCRIPTION

    .EXAMPLE
        PS C:\> New-SharedMailbox -Name "AddressTitle" -DisplayName "Full Address Title" -Alias "address" -Users "personA", "PersonB"
        Name           Alias
        ----           -----
        AddressTitle   address

    .EXAMPLE
        PS C:\> New-SharedMailbox -Name reports -DisplayName Reports -Alias reports -Users OfficeAdmin,Receptionist
        Name       Alias      ServerName    ProhibitSendQuota
        ----       -----      ----------    -----------------
        reports    reports    exchange01    unlimited
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

        Permissions:
        There's levels have been discovered in testing and not researched. Enterprise Administrators would be the top level that can do anything to exchange.
        The ability to Set-MailboxSentItemsConfiguration requires "Enterprise Administrators" group membership with the credentials used for import-mailserver.
        Creating the mailbox and assigning permissions requires only "Organization Management" membership.
    #>
    [CmdletBinding()]
    Param(
        # Will show as the name for the contact and mailbox
        [Parameter(Mandatory)]
        [string]
        $Name

        , # The name that appears in the Exchange Management Console under Recipient Configuration
        [string]
        $DisplayName

        , # Emaill address excluding @example.com
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
        $mailbox = Get-Mailbox -Identity $Name
        # By default sent items will only show in the senders account. This forces them into the shared sent items folder.
        Set-MailboxSentItemsConfiguration -Identity $mailbox.Identity -SendAsItemsCopiedTo 'SenderAndFrom' -SendOnBehalfOfItemsCopiedTo 'SenderAndFrom'
        $users | ForEach-Object {
            # Mailbox permissions allow the user to perform actions
            Add-MailBoxPermission $mailbox.Identity -AccessRights FullAccess -InheritanceType All -User $PSItem
            # AD permissions allow the users account#new (outlook) to find the mailbox
            Add-ADPermission $mailbox.Identity -ExtendedRights "Receive-As" -User $PSItem
            if($Send -eq 'From'){
                Add-ADPermission $mailbox.Identity -ExtendedRights "Send-As" -User $PSItem
            }
        }
        if($Send -eq 'Behalf'){
            # Send 'From' permissions will supersede this "On behalf" setting.
            # -GrantSendOnBehalfTo only replaces the list, you must append your own users list first.
            Set-Mailbox -Identity $mailbox.Identity -GrantSendOnBehalfTo $users > $null
            Write-Warning "The permission for users to send 'From' will supersede this permission, check their AD-Permissions for 'Send-As'"
        }
    }
    End{
        Write-Output $mailbox
    }
}

Export-ModuleMember -Function "Get-*", "Search-*", "New-*", "*-MailServer"
