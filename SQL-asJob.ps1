Function RebuildPaths ($ServerInstance){        
$SP_RebuildPaths = @"
    USE [IODC_Central]
    GO
    DECLARE @RC int
    DECLARE @Include_Drawings bit
    SET @Include_Drawings = 0
    EXECUTE @RC = [dbo].[Job_Cache_Partition_Paths_Set] @Include_Drawings
"@

    $job = "RebuildPaths"
    Get-Job $job | Remove-Job -Force -ErrorAction SilentlyContinue

    $SqlParams = @{
        ServerInstance = $ServerInstance
        Database = "IODC_Central"
        Query = $SP_RebuildPaths
        OutputSqlErrors = $true
        ErrorVariable = "SQL_Error"
        QueryTimeout = 300
    }

    $ImportParams = @{
        Name = "SQLPS"
        ErrorAction = "SilentlyContinue"
        DisableNameChecking = $true
    }
    Start-Job -ScriptBlock{
        param($ImportParams, $SqlParams)
        Import-Module @ImportParams
        Invoke-Sqlcmd @SqlParams 
    } -Name $job -ArgumentList ($ImportParams, $SqlParams) #-ComputerName $env:COMPUTERNAME
}
