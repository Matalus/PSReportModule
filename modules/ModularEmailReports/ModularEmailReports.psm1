#ModuleEmailReports.psm1 - Provides Flexible SQL reporting with modular options
#Matt Hamende - RunSmart 2017
#
#Parameter info below in params
#
#Loads the SQL PowerShell Module that will connect to SQL to run the query

Function ModularEmailReports {
    [CmdletBinding()]
    param (
        #Inputfile = UNC path to query location I.E. C:\queries\SQL_Query.sql
        [parameter(Mandatory = $false, ValueFromPipeline = $true)]
        $config,

        #EmailAddress = Email Address to send query results to I.E bsmith@io.com
        [parameter(Mandatory = $false)]
        $emailaddress,

        #From = Email Address to send as
        [parameter(Mandatory = $false)]
        $From,

        #Subject = Subject to be displayed on email
        [parameter(Mandatory = $false)]
        $Subject,
    
        #ServerInstance = SQL Instance - I.E. sqlserver.domain.com
        [parameter(Mandatory = $false)]
        $ServerInstance,
	
        #SMTPServer = SMTP SERVER to send query results to I.E mail.outlook.com
        [parameter(Mandatory = $false)]
        $SMTP,

        #Database =  DB name I.E. User_DB
        [parameter(Mandatory = $false)]
        $database
    )
    BEGIN {}

    PROCESS {
        Import-Module SQLPS -DisableNameChecking -ErrorAction SilentlyContinue
        Import-Module ZipFile -ErrorAction SilentlyContinue
    
        Write-Host "@
    ServerInstance : $($config.ServerInstance)
    Database       : $($config.Database)
    Recipients     : $($config.Recipients.toString())
    RowPreview     : $($config.RowPreview)
    Reports        : $($config.Reports.ReportName.toString())
    RunDir         : $($config.RunDir)
@"



        #Remove-Variable * -Force -ErrorAction SilentlyContinue
        Function Log($message) {
            "$(Get-Date -Format u) | $message"
        }
        

        #logs current directory - needed to ensure attachment can be located.
        $DIR = $config.RunDir
        write-host $DIR
        Set-Location $DIR
        $database = $config.Database
        $ServerInstance = $config.ServerInstance

        #Deletes temp files
        Get-ChildItem "$DIR\Temp" | Remove-Item -Force -ErrorAction Continue

        $ReportHead = @"
        <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
        <html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
        <head>
"@

        IF ($config.ReportHeader.Length -ge 1) {
            $ReportHead += $config.ReportHeader
        }

        $ReportHead += @"
        </head>
        <meta charset="UTF-8">
        <body>
        <table class='documenttitle'>
        <tbody>
        <tr class='documenttitle'>
        <td class='documenttitle'>$($config.Title)</td>
        </tr>
        </tbody>
        </table>        
        <div class="resultsets">
        (Results sets larger than $($config.RowPreview) rows will appear in zipped attachment)
        </div>
        <br>
"@
        IF ($config.deeplink.Length -ge 1) {
            $ReportHead += $config.deeplink
        }

        $Summary = @()
        $ReportCount = 0
        ForEach ($Report in $Config.Reports) {
            $ReportCount++            
            $Attempt = 0
            Do {
                $Attempt++ 
                $SQLError = $null               
                Write-Host "Running Report: $($Report.ReportName)"
                Try {
                    [array]$SQL = ExecuteQuery -Config $config -Report $Report
                    if ($SQL.Count -ge 1){
                        "Success! - Returned $($SQL.Count) rows"
                        $rowcount = $SQL.Count
                    }
                }
                Catch {
                    [array]$SQLError += $_ | Select-Object @{N="Attempts";E={$Attempt}},@{N="DateTime";E={$(Get-Date -Format u)}},@{N="Error";E={$_.Exception.Message}}
                    "Query Failed - $($SQLError.Error)"
                    "Attempt: $($Attempt) Retrying..."
                    $rowcount = $SQLError.Count                    
                }
            }While ($Attempt -le 3 -and $SQLError)
            
            IF($SQLError){
                $SQLError | Select-Object -First 10 * | Format-Table
                
            } ELSE {
                $SQL | Select-Object -First 10 * | Format-Table
            }

            $ReportStats = [pscustomobject]@{
                "Report Name"   = $Report.ReportName
                "Description"   = $Report.Description
                "Rows Returned" = $rowcount
            }
            $Summary += $ReportStats

            if ($rowcount -lt 1) {
                Write-Host "Returned $($rowcount) rows - breaking loop"
                Continue
            }


            #Mail Variables
            #SMTP server to send email
	
            $Report | Add-Member -NotePropertyName "rowcount" -NotePropertyValue $rowcount
        
            $BODY = $null
            $BODY += @"
        <table class="reportname">
        <tbody>
        <tr>
        <td class="reportname">$($Report.ReportName)</td>
        </tr>
        </tbody>
        </table>
        <table class="rowsreturned">
        <tbody>
        <tr class="rowsreturned">
        <!--[if !mso]><!-->
        <td class="rowsreturned" style="width: 120px;">
        <button id='table$($ReportCount)hide' onclick="jsHide('table$($ReportCount)','jsHideTable'); toggleText(this)"><span>Hide Rows</span></button>      
        </td>
        <!--<![endif]-->
        <td class="rowsreturned">Total Rows Returned : $rowcount $(IF($rowcount -ge $config.rowpreview){" - results exceed row preview of $($config.rowpreview) please see attachment for full results."})</td>
        </tr>
        </tbody>
        </table>
"@

            IF ($Report.Description.length -ge 1) {
                $BODY += @"
            <table class="descrip">
            <tbody>
            <tr>
            <td class="descrip">$($Report.Description)</td>
            </tr>
            </tbody>
            </table>
"@
            }



            $html = $Report.Header  
            #Converts SQL results to HTML Table
            IF($SQLError){
            
                $BODY += $SQLError | Select-Object * -first $($config.rowpreview) |
                convertto-html -fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd |
                out-string | Add-HTMLTableAttribute -AttributeName "id" -Value "table$($ReportCount)" |
                Add-HTMLTableAttribute -AttributeName "class" -Value "reporttable"
            } ELSE {
                $BODY += $SQL | Select-Object * -first $($config.rowpreview) |
                convertto-html -fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd |
                out-string | Add-HTMLTableAttribute -AttributeName "id" -Value "table$($ReportCount)" |
                Add-HTMLTableAttribute -AttributeName "class" -Value "reporttable"
            }
                

            #HTML Replace
            ForEach ($tag in $Report.ReplaceTags) {
                IF ($BODY -match $tag.oldChar) {
                    Write-Host @"       
                    Replacing Tag:
                old: $($tag.oldChar)
                new: $($tag.newChar)
"@
                    $BODY = $BODY.Replace($tag.oldChar, $tag.newChar)
                    #$BODY = $BODY.Replace('"', '\"')
                    $BODY | Out-File .\reporthtml.html -Force
                }
                ELSE {
                    $ErrString = "Unable to locate tag : $($tag.oldchar) : if this tag is unused consider removing it from config.reports.ReplaceTags `n"
                    Write-Error $ErrString   -ErrorAction Continue
                }
                
            }

            #Write-Host $Config.Reports

            
            $attach = ("$($Report.QueryPath).csv").Replace(".sql", "")
            $SQL | Select-Object * | Export-Csv -Path "$DIR\Temp\$attach" -NoTypeInformation

            $ReportBody += $BODY
            $ReportBody += "<BR>"        
    
        
        } #END MAIN REPORT LOOP


        #Write-Host "--------------------------------------------------------------------"
        #Write-Host $ReportBody
        #Write-Host "--------------------------------------------------------------------"
        $ReportSummary = @"
    <table class="reportname" style="width: 100%; text-align: left;">
    <tbody>
    <tr>
    <td class="reportname">Report Summary</td>
    </tr>
    </tbody>
    </table>
    <table class="rowsreturned">
    <tbody>
    <tr class="rowsreturned">
    <td class="rowsreturned">Total Reports : $($Summary.count)</td>
    </tr>
    </tbody>
    </table>
    <table class="descrip">
    <tbody>
    <tr>
    <td class="descrip">Summary of all reports in rollup - reports with no results won't appear in details</td>
    </tr>
    </tbody>
    </table>
"@
        $ReportSummary += $Summary | Select-Object * -first 500 |
            convertto-html -fragment | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd |
            out-string | Add-HTMLTableAttribute -AttributeName "ID" -Value "table0"
        $ReportSummary += "<BR>"
        $ReportEndBody = @"
        <script type="text/javascript">
        $($config.JavaScript)
        </script>
        </body>
        </html>
"@
        IF ($Config.Summary -ne $false -or $Config.Summary -is $null) {
            $FinalBody = $ReportHead + $ReportSummary + $ReportBody + $ReportEndBody
        }
        ELSE {
            $FinalBody = $ReportHead + $ReportBody + $ReportEndBody
        }
        $FinalBody | Out-File "$DIR\temp\Interactive_Report.html" -Force

        #Exports Query Results and creates attachment to send
        Set-Location $DIR

        IF (-not (Test-Path "$DIR\Temp")) {New-Item "$DIR\Temp" -ItemType Directory}




            
        $zippedname = "Report.zip"
        $ZIP = "$DIR\$zippedname"
        
        "Attach"
        $attach
        "DEST"
        "$DIR\$zippedname"
        

        Remove-Item -Path $ZIP -Force -ErrorAction Continue
        #$ZIP = New-ZipFile -InputObject $attach -ZipFilePath "$DIR\$zippedname" -Compression Optimal -ErrorAction SilentlyContinue
        ZipFiles -sourcedir "$DIR\Temp" -zipFullPath $ZIP

        #ForEach ($To in $config.Recipients) {
        #Array Splat to apply params to Send Email
        $mailsplat = @{
            To          = $config.Recipients
            From        = $config.From
            Subject     = $config.Subject
            Body        = $FinalBody
            BodyAsHtml  = $true
            SmtpServer  = $config.SMTPServer
            Attachments = $ZIP
        }
    
        #Sends Email

        Get-ChildItem $ZIP -ErrorAction SilentlyContinue
        "$(Get-Date -Format u) | Sending Report to $($mailsplat.to)"
        ""
        $Error.Clear()
        Try {
            $mailsend = Send-MailMessage @mailsplat          
        }
        Catch {$Error}
        #}

    }
    END {}
    #End Function
}  

function ZipFiles( [string]$sourcedir = "", [string]$zipFullPath = "" ) {
    $boolAlreadyCompleted = Test-Path $zipFullPath
    if ($boolAlreadyCompleted) {return $false}
    Add-Type -Assembly System.IO.Compression.FileSystem
    $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
    [System.IO.Compression.ZipFile]::CreateFromDirectory($sourcedir,
        $zipFullPath, $compressionLevel, $false)
    return Test-Path $zipFullPath
}

function ExecuteQuery($Config, $Report) {
    #Array Splat to apply params to SQL Invoke
        
    Write-Host "$DIR\queries\$($Report.QueryPath)"
    $splat = @{
        ServerInstance  = $config.ServerInstance
        Database        = $Config.Database
        InputFile       = "$DIR\queries\$($Report.QueryPath)"
        QueryTimeout    = 120
        OutputSqlErrors = $true
    }

    #Executes Query and Stores results in variable object
    $SQL = Invoke-Sqlcmd @splat
        
    $SQL = $SQL | Select-Object * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors -ErrorAction SilentlyContinue

    $SQL
};

Function Set-AlternatingRows {
	
    [CmdletBinding()]
   	Param(
       	[Parameter(Mandatory, ValueFromPipeline)]
        [string]$Line,
       
   	    [Parameter(Mandatory)]
       	[string]$CSSEvenClass,
       
        [Parameter(Mandatory)]
   	    [string]$CSSOddClass
   	)
    Begin {
        $ClassName = $CSSEvenClass
    }
    Process {
        If ($Line.Contains("<tr><td>")) {
            $Line = $Line.Replace("<tr>", "<tr class=""$ClassName"">")
            If ($ClassName -eq $CSSEvenClass) {
                $ClassName = $CSSOddClass
            }
            Else {
                $ClassName = $CSSEvenClass
            }
        }
        Return $Line
    }
}

Function Add-HTMLTableAttribute {
    Param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]
        $HTML,

        [Parameter(Mandatory = $true)]
        [string]
        $AttributeName,

        [Parameter(Mandatory = $true)]
        [string]
        $Value

    )

    $xml = [xml]$HTML
    $attr = $xml.CreateAttribute($AttributeName)
    $attr.Value = $Value
    $xml.table.Attributes.Append($attr) | Out-Null
    Return ($xml.OuterXML | out-string)
}
