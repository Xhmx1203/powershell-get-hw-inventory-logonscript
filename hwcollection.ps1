[CmdletBinding()]

Param (

    #[parameter(ValueFromPipeline=$True)]
   # [string]$ComputerName="localhost"

)

Begin
{
    #Initialize
    $path ="\\192.168.10.10\Com$\D_IT\04_HWInfo_Collection\HWInventory\hardwareinfo-collection.csv"
    $ospp_path="\\192.168.10.10\Com$\D_IT\04_HWInfo_Collection\bin\Office16\ospp.vbs"

}

Process
{

    #---------------------------------------------------------------------
    # Process each ComputerName
    #---------------------------------------------------------------------

    if (!($PSCmdlet.MyInvocation.BoundParameters[“Verbose”].IsPresent))
    {
  #    Write-Host "Processing $ComputerName"
    }

   # Write-Verbose "=====> Processing $ComputerName <====="

    #$htmlreport = @()
    $CollectInfo = New-Object PSobject
   # $htmlfile = "$($ComputerName).html"
   # $spacer = "<br />"

    #---------------------------------------------------------------------
    # Do 10 pings and calculate the fastest response time
    # Not using the response time in the report yet so it might be
    # removed later.
    #---------------------------------------------------------------------
    $officeDstatus = "cscript.exe $ospp_path /dstatus"
    $officeVersion =($officeDstatus | Select-String -Pattern "LICENSE NAME").ToString().Split(":")[1].Trim()
    $officeLast5Keys = ($officeDstatus | Select-String -Pattern "Last 5 characters of installed product key").ToString().Split(":")[1].Trim()
    


        #---------------------------------------------------------------------
        # Collect computer system information and convert to HTML fragment
        #---------------------------------------------------------------------
    
  #      Write-Verbose "Collecting computer system information"
        $CollectInfo | Add-Member NoteProperty -Name 'DateTime' -Value (get-date).toString('u')
        try
        {
            $csinfo = Get-WmiObject Win32_ComputerSystem
            
            $physicalmemory = [math]::round($csinfo.TotalPhysicalMemory/1GB).ToString()
            $CollectInfo | Add-Member NoteProperty -Name 'Name' -Value  $csinfo.Name
            $CollectInfo | Add-Member NoteProperty -Name 'Manufacturer' -Value $csinfo.Manufacturer
            $CollectInfo | Add-Member NoteProperty -Name 'Model' -Value $csinfo.Model
            $CollectInfo | Add-Member NoteProperty -Name 'Physical Processors' -Value $csinfo.NumberOfProcessors
            $CollectInfo | Add-Member NoteProperty -Name 'Total Physical Memory (Gb)' -Value $physicalmemory
            $CollectInfo | Add-Member NoteProperty -Name 'DnsHostName' -Value $csinfo.DnsHostName
            $CollectInfo | Add-Member NoteProperty -Name 'Domain' -Value $csinfo.Domain
            $CollectInfo | Add-Member NoteProperty -Name "UserName" -Value ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).toString()
            $CollectInfo | Add-Member NoteProperty -Name "WindowsKey" -Value (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
            $CollectInfo | Add-Member NoteProperty -Name "OfficeVersion" -Value  $officeVersion
            $CollectInfo | Add-Member NoteProperty -Name "OfficeLast5Key" -Value  $officeLast5Keys
        }catch
        {
   #         Write-Warning $_.Exception.Message
            #$CollectInfo #+= "<p>An error was encountered. $($_.Exception.Message)</p>"
            #$CollectInfo += $spacer
        }



        #---------------------------------------------------------------------
        # Collect operating system information and convert to HTML fragment
        #---------------------------------------------------------------------
    
    #    Write-Verbose "Collecting operating system information"
 
        try
        {
            $osinfo = Get-WmiObject Win32_OperatingSystem
            $Date1= [datetime]::ParseExact($osinfo.InstallDate.SubString(0,8),"yyyyMMdd",$null)
            $CollectInfo | Add-Member NoteProperty -Name 'Operating System' -Value $osinfo.Caption
            $CollectInfo | Add-Member NoteProperty -Name 'OSArchitecture' -Value $osinfo.OSArchitecture
            $CollectInfo | Add-Member NoteProperty -Name 'Version' -Value $osinfo.Version
            $CollectInfo | Add-Member NoteProperty -Name 'Install Date' -Value  $Date1
           
        }
        catch
        {
     #       Write-Warning $_.Exception.Message
            #$CollectInfo # += "<p>An error was encountered. $($_.Exception.Message)</p>"
            #$CollectInfo += $spacer
        }


        #---------------------------------------------------------------------
        # Collect BIOS information and convert to HTML fragment
        #---------------------------------------------------------------------
      #  Write-Verbose "Collecting BIOS information"

        try
        {
            $biosinfo = Get-WmiObject Win32_Bios
            $disk = get-disk -Number 0 | select Manufacturer, Model,SerialNumber,@{L="Size";E={($_.Size/1GB) -as [int] }}

            $CollectInfo | Add-Member NoteProperty -Name 'Serial Number' -Value $biosinfo.SerialNumber  
            $CollectInfo | Add-Member NoteProperty -Name 'Disk-Manufacturer' -Value $disk.Manufacturer
            $CollectInfo | Add-Member NoteProperty -Name 'Disk-Model' -Value $disk.Model
            $CollectInfo | Add-Member NoteProperty -Name 'Disk-SerialNumber' -Value $disk.SerialNumber
            $CollectInfo | Add-Member NoteProperty -Name 'Disk-Size' -Value $disk.Size
        }
        catch
        {
       #     Write-Warning $_.Exception.Message
           # $CollectInfo #+= "<p>An error was encountered. $($_.Exception.Message)</p>"
            #$CollectInfo += $spacer
        }
         

        #---------------------------------------------------------------------
        # Collect logical disk information and convert to HTML fragment
        #---------------------------------------------------------------------

              
        $CollectInfo | Export-Csv -Append -NoTypeInformation -Encoding UTF8 -Path $path 
      #New-Object psobject -Property $CollectInfo| Export-Csv .\info.csv -Encoding Default
    }

End
{
    #Wrap it up
   # Write-Verbose "=====> Finished <====="
}
