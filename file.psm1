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
        $Deny = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList 'BHS\AllStudents', 'FullControl', 'Deny'
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
        Remove the NTFS Deny ACL for a given item for students access. Users might still not have file permission
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Enable-StudentAccess | Convertto-html | Out-File "Report.html"
        Generate a report that lists the item students have been blocked from using.
    #>
    Param(
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String[]]
        $Path
    )
    Begin {
        $Rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList 'BHS\AllStudents', 'FullControl', 'Deny'
    }
    Process {
        $Path | Get-Acl | foreach-object {
            $psitem.RemoveAccessRule($Rule) > $null
            try { $psitem | Set-acl }
            catch {
                Throw "Failed remove Deny permission: $($psitem.path)"
            }
            Get-Item $psitem.Path | Write-Output
        }
    }
    End {
        Write-Warning "Enable-StudentAccess will remove 'Deny' permission set by Disable-StudentAccess. See Add-StudentAccess"
    }
}

function Add-StudentAccess {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Add "Allow" permission to target path.
    .DESCRIPTION
        Add the NTFS Allow ACL for a given item for students access. Any deny rule will supercede this.
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Add-StudentAccess | Convertto-html | Out-File "Report.html"
        Generate a report that lists the items students have been allowed to aceess.
    #>
    Param(
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String[]]
        $Path
        , # keyword access permission levels
        [ValidateSet('FullControl','Modify','ReadAndExecute')]
        [string]
        $Access = 'ReadAndExecute'
        , # Sets the Applies to contition for child items
        [ValidateSet('All','ThisFolder')]
        [string]
        $Inherit
    )
    Begin {
        switch ($inherit) {
            'ThisFolder' { $Inherritance = @('ObjectInherit') }
            Default { $Inherritance = @('ContainerInherit', 'ObjectInherit') }
        }
        $FolderRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList 'BHS\AllStudents', $Access, $Inherritance, 'None', 'Allow'
        $FileRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList 'BHS\AllStudents', $Access, 'Allow'
    }
    Process {
        $Path | Get-Acl | foreach-object {
            if ( (Get-Item $Path) -is [System.IO.DirectoryInfo] ){
                $psitem.SetAccessRule($FolderRule)
            } else {
                $psitem.SetAccessRule($FileRule)
            }
            try { $psitem | Set-acl }
            catch {
                Throw "Failed to add permission: $($psitem.path)"
            }
            Get-Item $psitem.Path | Write-Output
        }
    }
}

function Remove-StudentAccess {
    [CmdletBinding()]
    <#
    .SYNOPSIS
        Remove "Allow" permission added by "Add-StudentAccess"
    .DESCRIPTION
        Remove the NTFS Allow ACL for a given item for students access.
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Remove-StudentAccess | Convertto-html | Out-File "Report.html"
        Generate a report that lists the items students have been blocked from using.
    #>
    Param(
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [String[]]
        $Path
    )
    Begin {
    }
    Process {
        $Path | Get-Acl | foreach-object {
            $Rule = $psitem.Access | Where-Object { ($_.IdentityReference -eq 'BHS\AllStudents') -and ($_.AccessControlType -eq  'Allow')}
            $psitem.RemoveAccessRule($Rule) > $null
            try { $psitem | Set-acl }
            catch {
                Throw "Failed to remove permission: $($psitem.path)"
            }
            Get-Item $psitem.Path | Write-Output
        }
    }
}

Export-ModuleMember -Function "*"
