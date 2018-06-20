function Disable-Access {
    <#
    .SYNOPSIS
        Deny student group permission to a file.
    .DESCRIPTION
        Add an NTFS Deny ACL to a given file for students to block their access.
        A deny rule supercedes any allow rule.
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Disable-Access | Convertto-html | Out-File "Report.html"
        Generate a report that lists the files students have been blocked from using.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline,
                   ValueFromPipelineByPropertyName,
                   Position = 0)]
        [String[]]
        $Path
        , # AD identity
        [Parameter(ValueFromPipelineByPropertyName,
                   Position = 1)]
        [string]
        $Identity = 'AllStudents'
    )
    Begin {
        $Deny = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $Identity, 'FullControl', 'Deny'
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

function Enable-Access {
    <#
    .SYNOPSIS
        Remove "Deny" permission added by "Deny-Access"
    .DESCRIPTION
        Remove the NTFS Deny ACL for a given item for students access.
        Does not test that users get access permission. See Add-Access
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Enable-Access | Convertto-html | Out-File "Report.html"
        Generate a report that lists the item students have been blocked from using.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline,
                   ValueFromPipelineByPropertyName,
                   Position = 0)]
        [String[]]
        $Path
        , # AD identity
        [Parameter(ValueFromPipelineByPropertyName,
                   Position = 1)]
        [string]
        $Identity = 'AllStudents'
    )
    Begin {
        $Rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $Identity, 'FullControl', 'Deny'
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
        Write-Warning "Enable-StudentAccess will remove 'Deny' permission set by Disable-StudentAccess. See Add-Access"
    }
}

function Add-Access {
    <#
    .SYNOPSIS
        Add "Allow" permission to target path.
    .DESCRIPTION
        Add the NTFS Allow ACL for a given item for students access.
        Any deny rule will supercede this. See Enable-Access.
    .PARAMETER Inherit
        Choose to edit a folder ACL or all files within the folder

        Set the "Applies to" contition for child items when All files

    .EXAMPLE
        Add-Access .\en-US\
        By default allow students to read permission to the folder, subfolder and files.
    .EXAMPLE
        Add-Access .\en-US\ -Inherit ThisFolder
        Allow students to read permission to this folder folder only.

        Useful to only allow the account to check further permission,
        this way you can add-access to files below this tree and access-based-enumeration will
        hide them from the user.
    .EXAMPLE
        Add-Access .\en-US\ -Identity '2016 Students'
        Specify the Active Directory identiy name to set access for;

            Directory: N:\Documents\src\ps-helper


        Mode                LastWriteTime         Length Name
        ----                -------------         ------ ----
        d-----       19/10/2016     15:33                en-US
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Add-Access | Convertto-html | Out-File "Report.html"
        Generate a report that lists the items students have been allowed to aceess.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline,
                   ValueFromPipelineByPropertyName,
                   Position = 0
                   )]
        [String[]]
        $Path
        , # AD identity to grant access
        [Parameter(ValueFromPipelineByPropertyName,
                   Position = 1)]
        [string[]]
        $Identity = 'AllStudents'
        , # keyword access permission levels
        [ValidateSet('FullControl','Modify','ReadAndExecute')]
        [string]
        $Access = 'ReadAndExecute'
        , #
        [ValidateSet('All','ThisFolder')]
        [string]
        $Inherit
    )
    Begin {
        switch ($inherit) {
            'ThisFolder' { $Inherritance = @('None') }
            Default { $Inherritance = @('ContainerInherit', 'ObjectInherit') }
        }
        $FolderRule = New-Object System.Collections.ArrayList
        $FileRule = New-Object System.Collections.ArrayList
        foreach($id in $Identity){
            $FolderRule.Add( (New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $id, $Access, $Inherritance, 'None', 'Allow') ) > $null
            $FileRule.Add( (New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $id, $Access, 'Allow') ) > $null
        }
    }
    Process {
        $Path | Get-Acl | foreach-object {
            $acl = $PSItem
            $item = Get-Item $PSItem.path
            try {
                if ( $item -is [System.IO.DirectoryInfo] ){
                    foreach($rule in $FolderRule){
                        $acl.SetAccessRule($rule)
                    }
                } else {
                    foreach($rule in $FileRule){
                        $acl.SetAccessRule($rule)
                    }
                }
                $acl | Set-acl
            }
            catch {
                Throw "Failed to add permission: $($acl.path)"
            }
            Write-Output $item
        }
    }
}

function Remove-Access {
    <#
    .SYNOPSIS
        Remove "Allow" permission added by "Add-StudentAccess"
    .DESCRIPTION
        Remove the NTFS Allow ACL for a given item for students access.
    .EXAMPLE
        PS C:\> \\<server\<share>\<Path> | Remove-StudentAccess | Convertto-html | Out-File "Report.html"
        Generate a report that lists the items students have been blocked from using.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline,ValueFromPipelineByPropertyName,
                   Position = 0)]
        [String[]]
        $Path
        , # AD identity to remove access
        [Parameter(ValueFromPipelineByPropertyName,
                   Position = 1)]
        [string]
        $Identity = 'AllStudents'
    )
    Begin {
    }
    Process {
        $Path | Get-Acl | foreach-object {
            $Rule = $psitem.Access | Where-Object { ($_.IdentityReference -like "*$Identity") -and ($_.AccessControlType -eq  'Allow')}
            $psitem.RemoveAccessRule($Rule) > $null
            try { $psitem | Set-acl }
            catch {
                Throw "Failed to remove permission: $($psitem.path)"
            }
            Get-Item $psitem.Path | Write-Output
        }
    }
}

function Show-Access {
    <#
    .SYNOPSIS
        Display the ACL information for given files or folders.
    .DESCRIPTION
        Wraps Get-ACL to select the access list while retaining the file path being inspected.
        Useful for Group-Object inspecting access controls.

        This function is intended for interactive use inspecting a network share, as such it breaks best practice of avoiding the use of format-* commands.
    .EXAMPLE
        Show-Access '.\INFO.lnk'

        Path                               IdentityReference            IsInherited AccessControlType            FileSystemRights
        ----                               -----------------            ----------- -----------------            ----------------
        \\ORG\Desktop\Start Menu\INFO.lnk  ORG\Office                         False             Allow ReadAndExecute, Synchronize
        \\ORG\Desktop\Start Menu\INFO.lnk  ORG\Support                         True             Allow                 FullControl
        \\ORG\Desktop\Start Menu\INFO.lnk  NT AUTHORITY\SYSTEM                 True             Allow                 FullControl
        \\ORG\Desktop\Start Menu\INFO.lnk  NT AUTHORITY\NETWORK SERVICE        True             Allow                 FullControl
        \\ORG\Desktop\Start Menu\INFO.lnk  BUILTIN\Administrator               True             Allow                 FullControl
        \\ORG\Desktop\Start Menu\INFO.lnk  ORG\Domain Admin                    True             Allow                 FullControl
        \\ORG\Desktop\Start Menu\INFO.lnk  ORG\Share Administrator             True             Allow                 FullControl

        Internal use of Format-Table to ease reading at the console.
    .EXAMPLE
        '.\.gitignore','.\asset.psm1' | show-access -NoFormat

        Path              : N:\Documents\helper\.gitignore
        IdentityReference : ORG\jbennett
        IsInherited       : True
        AccessControlType : Allow
        FileSystemRights  : FullControl

        Path              : N:\Documents\helper\.gitignore
        IdentityReference : BUILTIN\Administrator
        IsInherited       : True
        AccessControlType : Allow
        FileSystemRights  : FullControl

        Use the NoFormat switch to bypass internal use of Format-Table for piping objects to further functions.
    .NOTES
        Should refactor the format switch into a custom native cmdlet format document.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0,Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [string]
        $Path

        , # Do not automatically pipe output to Format-Table
        [switch]$NoFormat = $false
    )
    Process {
        $output = $path |
            get-acl |
            Foreach {
                $item = $psItem | select-Object -Expandproperty Access
                $item | Add-Member -MemberType NoteProperty -Name Path -Value $psItem.path
                Write-Output $Item
            } |
            select @{
                    name='Path'
                    expression = { $_.path.replace('Microsoft.PowerShell.Core\FileSystem::','') }
                },
                'IdentityReference',
                'IsInherited',
                'AccessControlType',
                'FileSystemRights'

        if($NoFormat){
            Write-Output $output
        } else{
            $output | Format-Table | Write-Output
        }
    }
}

Register-ArgumentCompleter -CommandName 'Add-Access','Remove-Access','Enable-Access','Disable-Access' -ParameterName 'Identity' -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    [System.Collections.ArrayList]$preset = @(
        'AllStudents'
        'AllCAStudents'
        'AllStaff'
        'Office'
        "'Site Management'"
        "'Student Teachers'"
        "'Teaching Staff'"
        'Govenors'
        "'Exam Candidate'"
    )

    [int]$year = (get-date).year
    for( $i = $year; $i -ge ($year -6); $i -= 1){
        $preset.add("'$i Students'") > $null
    }

    $preset |
        ForEach-Object {
            [System.Management.Automation.CompletionResult]::new($psitem, $psitem, 'ParameterValue', ("AD Name: " + $psitem))
        }
}

Export-ModuleMember -Function "*"
