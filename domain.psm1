Function Get-LocalLogonProfile {
    [CmdletBinding()]
    Param(
        [switch]
        $IncludeDefault

        ,
        [string[]]
        $ComputerName
    )
    Begin {
        $DefaultSID = @(
            'S-1-5-18'
            'S-1-5-19'
            'S-1-5-20'
            'S-1-5-21-3316553321-2298761035-1846811995-500'
        )
    }
    Process {
        $list = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-ChildItem 'hklm:SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' |
                Get-ItemProperty -Name ProfileImagePath
        } |
            Select-Object @(
            @{
                name       = 'Name';
                Expression = { split-path -leaf $_.ProfileImagePath}
            }
            @{
                name       = 'SID';
                Expression = { split-path -leaf $_.pspath}
            }
            @{
                name       = 'Path';
                Expression = { $_.pspath }
            }
        )
        if ($IncludeDefault) {
            Write-Output $list
        }
        else {
            Write-Output ($list | Where-Object Sid -NotIn $DefaultSID)
        }
    }
}

function New-ComputerList {
    <#
    .Synopsis
       Create a computer list given a room name and computer number
    .DESCRIPTION
       Will create a computer list given a room name and computer number without verifying if these names exist in the network.

       Pipeline output is an object with properties computerName
    .EXAMPLE
       New-ComputerList -Room 50 -Computer (8..11)

        ComputerName
        ------------
        50PC08
        50PC09
        50PC10
        50PC11
    .EXAMPLE
        New-ComputerList 40 -Computer 2,3,4,10 | Invoke-Command { ... }

        Pipe the computer name into a remote execution command.
    .EXAMPLE
        New-ComputerList o16 (1..10) -unc g

        ComputerName
        ------------
        \\o16PC01\g$\
        \\o16PC02\g$\
        \\o16PC03\g$\
        \\o16PC04\g$\
        \\o16PC05\g$\
        \\o16PC06\g$\
        \\o16PC07\g$\
        \\o16PC08\g$\
        \\o16PC09\g$\
        \\o16PC10\g$\
    .EXAMPLE
        1..4 | New-ComputerList -Room A,B -Exclude 3
        Pipe the range 1,2,3,4 into the Computer property and get 2 room names worth A and B. Skip PC 3 in both rooms.

        ComputerName
        ------------
        APC01
        BPC01
        APC02
        BPC02
        APC04
        BPC04
    .EXAMPLE
        New-ComputerList -Computer (1..4) -Room A,B -Exclude 3
        Expand a range to 1,2,3,4 for the Computer property and get 2 room names worth A and B. Skip PC 3 in both rooms.

        ComputerName
        ------------
        APC01
        BPC01
        APC02
        BPC02
        APC04
        BPC04
    #>
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([System.Management.Automation.PSCustomObject])]
    Param
    (
        # Room number
        [Parameter(Mandatory,
                   ValueFromPipeline,
                   ValueFromPipelineByPropertyName,
                   Position = 0)]
        [string[]]
        $Room

        , # Start PC number, default is 1. leading zero is unnecessary
        [Parameter(Position = 1,
                   ValueFromPipeline,
                   ValueFromPipelineByPropertyName)]
        [uint16[]]
        $Computer = 1

        , # Exclude specific numbers
        [Parameter(Position = 2,
                   ValueFromPipelineByPropertyName)]
        [uint16[]]
        $Exclude

        , # Drive letter in UNC path
        [Parameter(Position = 3,
                   ValueFromPipelineByPropertyName)]
        [ValidateSet('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z')]
        [string]
        $Drive = 'c'
    )

    Begin {
        # The local fully qualified domain name
        $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().name
    }

    Process {
        # So for each room we've been given
        $room.forEach({
            $room = $psitem
            Write-Verbose "Making list for Room: $room"

            # Make a computer for each of the computers specified, ignoring "Exclude" list
            $computer.where({$psitem -notin $Exclude}).foreach({
                Write-Verbose "adding computer: $psitem to the room list"
                # Add this string to our list of computernames with a fully qualified domain name
                $computerName = "$room`PC$( $psitem.toString('00') )"

                [psCustomObject]@{
                    "ComputerName" = $computerName
                    "FQDN" = "$computerName.$domain"
                    "Path" = "\\$computerName\$Drive$\"
                }
            })
        })
    }
}

Export-ModuleMember -Function "*"
