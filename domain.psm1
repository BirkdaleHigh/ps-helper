Function Get-LocalLogonProfile{
    [CmdletBinding()]
    Param(
        [switch]
        $IncludeDefault

        ,
        [string[]]
        $ComputerName
    )
    Begin{
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
                    name = 'Name';
                    Expression = { split-path -leaf $_.ProfileImagePath}
                }
                @{
                    name = 'SID';
                    Expression = { split-path -leaf $_.pspath}
                }
                @{
                    name = 'Path';
                    Expression = { $_.pspath }
                }
            )
        if($IncludeDefault){
            Write-Output $list
        } else {
            Write-Output ($list | Where-Object Sid -NotIn $DefaultSID)
        }
    }
}

Export-ModuleMember -Function "*"
