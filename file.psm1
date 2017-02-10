function Disable-StudentAccess {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Deny student group permission to a file.
    .DESCRIPTION
        Add an NTFS Deny ACL to a given file for students to block their access.
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Disable-StudentAccess | Convertto-html | Out-File "Report.html"
        Generate a report that lists the files students have been blocked from using.
    #>
    Param(
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String[]]
        $Path
    )
    Begin {
        $Deny = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList 'BHS\W7Students', 'FullControl', 'Deny'
    }
    Process {
        $Path | Get-Acl | foreach-object {
            $psitem.SetAccessRule($Deny)
            try { $psitem | Set-acl }
            catch {
                Throw "Failed to set permission: $($psitem.path)"
            }
            Get-Item $psitem.Path | Write-Output
        }
    }
}

function Enable-StudentAccess {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Remove "Deny" permission added by "Enable-StudentAccess"
    .DESCRIPTION
        Remove the NTFS Deny ACL for a given file for students access.
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Enable-StudentAccess | Convertto-html | Out-File "Report.html"
        Generate a report that lists the files students have been blocked from using.
    #>
    Param(
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String[]]
        $Path
    )
    Begin {
        $Deny = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList 'BHS\W7Students', 'FullControl', 'Deny'
    }
    Process {
        $Path | Get-Acl | foreach-object {
            $psitem.RemoveAccessRule($Deny)
            try { $psitem | Set-acl }
            catch {
                Throw "Failed to set permission: $($psitem.path)"
            }
            Get-Item $psitem.Path | Write-Output
        }
    }
}

Export-ModuleMember -Function "*"
