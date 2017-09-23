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
    Process{
        Get-WmiObject -Class wmiMonitorID -Namespace "root\wmi" -ComputerName $ComputerName |
            Select-Object @{
                name = 'ComputerName'
                expression = { $psitem.__SERVER }
            },
            @{  name = 'Name'
                expression = { [System.Text.Encoding]::ASCII.GetString($psitem.ManufacturerName) }
            },
            @{  name = 'Serial'
                expression = { [System.Text.Encoding]::ASCII.GetString($psitem.SerialNumberID) }
            },
            YearOfManufacture

    }
}

Export-ModuleMember -Function "Get-MonitorInformation"
