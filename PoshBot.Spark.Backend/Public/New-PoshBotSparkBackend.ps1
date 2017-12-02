function New-PoshBotSparkBackend {
    <#
    .SYNOPSIS
        Create a new instance of a Spark backend
    .DESCRIPTION
        Create a new instance of a Spark backend
    .PARAMETER Configuration
        The hashtable containing backend-specific properties on how to create the Spark backend instance.
    .EXAMPLE
        PS C:\> $backendConfig = @{Name = 'SparkBackend'; Token = '<SPARK-API-TOKEN>'}
        PS C:\> $backend = New-PoshBotSparkBackend -Configuration $backendConfig

        Create a Spark backend using the specified API token
    .INPUTS
        Hashtable
    .OUTPUTS
        SparkBackend
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='*')]
    [cmdletbinding()]
    param(
        [parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('BackendConfiguration')]
        [hashtable[]]$Configuration
    )

    process {
        foreach($item in $Configuration) {
            if(-not $item.Token) {
                throw 'Configuration is missing [Token] parameter'
            } else {
                Write-Verbose 'Creating new Spark backend instance'
                $backend = [SparkBackend]::new($item.Token)
                if($item.Name) {
                    $backend.Name = $item.Name
                }
                $backend
            }
        }
    }
}

Export-ModuleMember -Function 'New-PoshBotSparkBackend'