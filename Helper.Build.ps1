Task InstallDependencies {
    $Name = @(
        'Pester'
        'PSScriptAnalyzer'
        'PSDeploy'
    )
    foreach($dep in $Name){
        if(-not (Get-module -ListAvailable $dep)){
            Install-Module $dep -Force
        }
    }
}

Task Analyze {
    $params = @{
        Path = "$BuildRoot\"
        Severity = @('Error', 'Warning')
        Recurse = $true
        Verbose = $false
        ExcludeRule = 'PSUseDeclaredVarsMoreThanAssignments', 'PSAvoidUsingCmdletAliases'
    }
    $results = Invoke-ScriptAnalyzer @params

    if ($results) {
        $results | Format-Table
        throw "One or more PSScriptAnalyzer errors/warnings where found."
    }
}

task Test {
    $invokePesterParams = @{
        Strict = $true
        PassThru = $true
        Verbose = $false
        EnableExit = $false
    }

    # Publish Test Results as NUnitXml
    $testResults = Invoke-Pester @invokePesterParams;

    $numberFails = $testResults.FailedCount
    assert($numberFails -eq 0) ('Failed "{0}" unit tests.' -f $numberFails)
}

Task Clean {
    $buildfolder = ".\dist"
    if(test-path $buildfolder){
        Remove-Item -Recurse -Force $buildfolder
    }
}

Task Build {
    $ModuleName = "Helper"

    $path = new-item -ItemType Directory '.\dist' -Force
    $artifact = join-path $path "$ModuleName.zip"
    Compress-Archive -Path .\*.psm1 -DestinationPath $artifact

    Compress-Archive -Update -Path ".\*.psd1"  -DestinationPath $artifact
    Compress-Archive -Update -Path ".\LICENSE"  -DestinationPath $artifact
    Compress-Archive -Update -Path ".\README.md"  -DestinationPath $artifact
    Compress-Archive -Update -Path ".\en-US"  -DestinationPath $artifact
}

Task . InstallDependencies, Analyze, Test, Clean, Build
