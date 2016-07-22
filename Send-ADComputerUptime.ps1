function Send-ADComputerUptime
{
  
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter(ValueFromPipeline=$true)]$Objects,
    [Parameter(Mandatory=$True)][string[]]$Recipient,
    [Parameter(Mandatory=$True)][string]$Sender,
    [Parameter(Mandatory=$True)][string]$EmailServer,
    [string]$SqlConnectionString,
    [switch]$AsAttachment = $false
  )
  
  
  
  Begin{
    #Import-Module -Name C:\Scripts\Modules\SQLReporting
    Write-Verbose 'Begin Block'
    Write-Verbose 'Initializing object arrays'
    $FormattedObjects = @()
    if (-not $Objects) { $Objects = Get-ReportData -TypeName Report.ADCUptime -ConnectionString $SqlConnectionString }
  }
  
  Process{
    foreach ($Object in $Objects)
    {
      $FormattedObjects += Select-Object -InputObject $Object -Property ComputerName,@{l='DaysUptime';e={ (New-TimeSpan -Start $_.LastBootUptime -End $_.DateCaptured).Days }},DateCaptured
    }
  }
  
  End{

    $FormattedObjects = $FormattedObjects | Sort-Object -Property DaysUptime -Descending  | Where-Object -FilterScript {$_.DateCaptured -ge (Get-Date).AddDays(-1)} | Select-Object -Property ComputerName,DaysUptime
    
    # CSS - Doesn't format well with Windows version of Outlook due to Word being used as rendering engine
    $css = '<style>
      table { width:98%; }
      td { text-align:center; padding:5px; }
      th { background-color:blue; color:white; }
      h3 { text-align:center }
      h6 { text-align:center }
    </style>'
    
    Write-Verbose 'End Block'
    Write-Verbose 'Building HTML report'
    
    $ReportDate = (Get-Date).ToShortDateString()
    $ObjectsHTML = $FormattedObjects | ConvertTo-Html -Fragment -PreContent "<h3>Uptime Statistics as of $ReportDate</h3>" | Out-String
    $FooterHtml = ConvertTo-Html -Fragment -PostContent "<h6>This report was run from:  $env:COMPUTERNAME on $(Get-Date)</h6>" | Out-String
    
    Write-Verbose "Sending Email:
      Recipient   : $Recipient
      Sender      : $Sender
    EmailServer : $EmailServer"
    
    if ($AsAttachment){
      $Report = ConvertTo-Html -Body "$ObjectsHTML $FooterHtml $css" | Out-File $env:TMP\uptime.html
      Write-Verbose "$Report"
      Send-MailMessage -to $Recipient -From $Sender -Subject 'Uptime Report' -Body 'Please find the attached uptime report.' -Attachments $env:TMP\uptime.html -SmtpServer $EmailServer
    }
    else{
      $Report = ConvertTo-Html -Body "$ObjectsHTML $FooterHtml $css" | Out-String
      Write-Verbose "$Report"
      Send-MailMessage -to $Recipient -From $Sender -Subject 'Uptime Report' -BodyAsHtml $Report -SmtpServer $EmailServer
    }
    
    
  }
  
}
