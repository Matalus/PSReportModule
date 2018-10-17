        #Launch Script for Modular Email Reports
        $ErrorActionPreference = "Stop"
        
        #Define working Dir
        $RunDir = split-path -parent $MyInvocation.MyCommand.Definition

        #Define Log Dir
        $log = "$RunDir\lastrun.log"

        #Start Logging
        Try{Stop-Transcript -ErrorAction SilentlyContinue}Catch{}
        Try{Start-Transcript $log -ErrorAction continue -Force}Catch{}

        #Load Params from JSON config
        $Params = (Get-Content "$RunDir\config.json") -join "`n" | ConvertFrom-Json
        #Uncomment below line to run sample
        #$Params = (Get-Content "$RunDir\sample_config.json") -join "`n" | ConvertFrom-Json

        Set-Location $RunDir

        #Append RunDir to params object for report Module use
        $Params | Add-Member -NotePropertyName "RunDir" -NotePropertyValue $RunDir

        #Loads Dependent Modules
        #Pre-requistes include SQLPS module - installed with SSMS
        #ZipFiles 
        #Modular Email Reports
        Import-Module SQLPS -DisableNameChecking -ErrorAction SilentlyContinue
        Import-Module "$RunDir\Modules\ZipFile\ZipFile.psm1" -ErrorAction SilentlyContinue
        Remove-Module ModularEmailReports -ErrorAction SilentlyContinue
        Import-Module "$RunDir\Modules\ModularEmailReports\ModularEmailReports.psm1" #-ErrorAction SilentlyContinue

        #Run Report Module with params
        ModularEmailReports -config $Params

        #Stop logging
        Try{Stop-Transcript -ErrorAction SilentlyContinue}Catch{}



