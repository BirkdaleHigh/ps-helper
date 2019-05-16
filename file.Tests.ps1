Import-Module './file.psm1' -Force

$TestDirectoryPath = Join-Path $Env:Temp "Peseter-Add-Access"

Describe "Add-Access" {
    BeforeEach {
        New-Item -ItemType Directory -Path $TestDirectoryPath
    }
    AfterEach {
        Remove-Item -Force -Path $TestDirectoryPath
    }
    It "Adds the correct identity to the ACL"{
        $testItem = Get-Item $TestDirectoryPath
        $testID = "{0}\{1}" -f $env:computername, 'Guest'
        $before = ($testItem | get-acl).Access | where-object IdentityReference -eq $TestID
        $before.count | should -BeExactly 0
        $testItem |
            Add-Access -Identity $testID

        $after = ($testItem | get-acl).Access | where-object IdentityReference -eq $TestID
        $after.count | should -BeExactly 1
    }
    It "Add Read and Execute file permission"{
        $testItem = Get-Item $TestDirectoryPath
        $testID = "{0}\{1}" -f $env:computername, 'Guest'
        $testItem |
            Add-Access -Identity $testID -Access 'ReadAndExecute'

        $acl = ($testItem | get-acl).Access | where-object IdentityReference -eq $TestID
        $acl.FileSystemRights -match 'Write'   | should -BeFalse
        $acl.FileSystemRights -match 'Execute' | should -BeTrue
        $acl.FileSystemRights -match 'Read'    | should -BeTrue
    }
    It "Add write, Read and Execute file permission"{
        $testItem = Get-Item $TestDirectoryPath
        $testID = "{0}\{1}" -f $env:computername, 'Guest'
        $testItem |
            Add-Access -Identity $testID -Access 'Write'

        $acl = ($testItem | get-acl).Access | where-object IdentityReference -eq $TestID
        $acl.FileSystemRights -match 'Write'   | should -BeTrue
        $acl.FileSystemRights -match 'Execute' | should -BeTrue
        $acl.FileSystemRights -match 'Read'    | should -BeTrue
    }
    It "Fail to set an incorrect permission"{
        $testItem = Get-Item $TestDirectoryPath
        $testID = "{0}\{1}" -f $env:computername, 'Guest'
        {$testItem | Add-Access -Identity $testID -Access 'Potato'} | Should -Throw
    }
}