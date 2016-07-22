function Get-ADComputerUptime {
  [CmdletBinding()]
  
  param(
    [Parameter(Mandatory=$True)][string[]]$OrganizationalUnits,
    [switch]$NoSql=$false,
    [string]$SqlConnectionString
  )
  
  $Computers = @()
  
  foreach ($OU in $OrganizationalUnits)
  {
    $Computers += (Get-ADComputer -SearchBase $OU -Filter *).Name
  } # End foreach loop
  
  foreach ($Computer in $Computers)
  {
    $obj = Get-WmiObject -ComputerName $Computer -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
    if ($obj)
    {
      
      $prop = @{
                'ComputerName' = $obj.PSComputerName
                'LastBootUptime' = $obj.ConvertToDateTime($obj.LastBootUptime)
                'DateCaptured' = (Get-Date)
            }
      $object = New-Object -TypeName PSObject -Property $prop
      $object.PSObject.TypeNames.Insert(0,'Report.ADCUptime')
      if (-not $NoSql) { Write-Output $object  | Save-ReportData -ConnectionString $SqlConnectionString } # End if statement
      Write-Output $object
    } # End if statement
    else { Write-Warning "Unable to connect to computer $Computer"} # End else statement
    
  } # End foreach loop
  
}