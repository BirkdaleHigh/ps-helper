function Search-MailFrom () {
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
    )
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd')) AND from:`"$from`"" -TargetMailBox $env:username -TargetFolder "Search" -DeleteContent:$Delete
}

function Search-MailSubject () {
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
    )
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd')) AND subject:`"$Subject`"" -TargetMailBox $env:username -TargetFolder "Search" -DeleteContent:$Delete
}

function Search-MailDate () {
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
    )
    "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd'))"
    Search-Mailbox -identity $Identity -SearchQuery "received:$($Start.toString('yyyy-MM-dd'))..$($End.toString('yyyy-MM-dd'))" -TargetMailBox $env:username -TargetFolder "Search" -DeleteContent:$Delete
}

$script:session
function Import-MailServer () {
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
function Get-MailServer (){
    <#
    .SYNOPSIS
        Get the mailserver session
    #>
    return $script:session
}
function Remove-Mailserver () {
    <#
    .SYNOPSIS
        Removes the imported remote connection to the mail server..
    .EXAMPLE
        Remove-MailServer
    #>
    Remove-PSSession $script:session
}
