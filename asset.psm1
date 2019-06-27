function Get-MonitorInformation {
    <#
    .SYNOPSIS
        Gets attached monitor product information for inventory.
    .DESCRIPTION
        Find the monitor manufacture date, name and serial number for a target computer.
    .EXAMPLE
        "itspc04" | Get-MonitorInformation

        ComputerName Name Serial       YearOfManufacture
        ------------ ---- ------       -----------------
        ITSPC04      AOC  D3EXXXXXXXXX              2014
        ITSPC04      AOC  D3EXXXXXXXXX              2014
    .NOTES
        Todo: Detect trailing zeros in serial number for correct length.
    #>
    [cmdletBinding()]
    param(
        # Target computer
        [parameter(ValueFromPipeline,
            ValueFromPipelineByPropertyName)]
        [string[]]
        $ComputerName = 'localhost'
    )
    Process {
        Get-CimInstance -Class wmiMonitorID -Namespace "root\wmi" -ComputerName $ComputerName |
            Select-Object @{
            name       = 'ComputerName'
            expression = { $psitem.PSComputerName }
        },
        @{  name       = 'Name'
            expression = { [System.Text.Encoding]::ASCII.GetString($psitem.ManufacturerName) }
        },
        @{  name       = 'Serial'
            expression = { [System.Text.Encoding]::ASCII.GetString($psitem.SerialNumberID) }
        },
        'YearOfManufacture'

    }
}

function Get-Memory {
    [CmdletBinding()]
    [OutputType([psobject])]
    Param
    (
        # Target computer names
        [Parameter(ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $ComputerName = 'localhost'
    )

    Begin {
        # Inconsistent notation to increase readability.
        $ListOfMemory = @{
            20 = "DDR"
            21 = "DDR 2"
            24 = "DDR 3"
            26 = "DDR 4"
            0  = "Unknown"
        }

        $Speed = @{
            Name       = "Speed (MHz)"
            Expression = {
                $_.Speed
            }
        }
        $Capacity = @{
            Name       = "Capacity (MB)"
            Expression = {
                [Math]::Round( $_.Capacity / 1mb , 0)
            }
        }
        $SourceComputer = @{
            Name       = "ComputerName"
            Expression = {
                $_.PSComputerName
            }
        }
        $MemoryType = @{
            Name       = "Memory Type"
            Expression = {
                $ListOfMemory.Item([int]$_.SMBIOSMemoryType)
            }
        }

    }
    Process {
        Get-CimInstance -ClassName 'win32_PhysicalMemory' -ComputerName $ComputerName |
            select-Object @(
                $SourceComputer
                'Manufacturer'
                'PartNumber'
                $Speed
                $Capacity
                'DeviceLocator'
                $MemoryType
            ) |
            where-Object { $_.DeviceLocator -notmatch "SYSTEM ROM" } |
            Write-Output
    }
}

Export-ModuleMember -Function "Get-MonitorInformation", "Get-Memory"
