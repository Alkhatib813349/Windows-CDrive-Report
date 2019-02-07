# -----------------------
# Define Global Variables
# -----------------------
$Global:Folder = $env:USERPROFILE+"\Documents\WindowsVM-Disk1-Reports" 
$Global:VCName = $null
$Global:Creds = $null

#*****************
# Get VC from User
#*****************
Function Get-VCenter {
    [CmdletBinding()]
    Param()
    #Prompt User for vCenter
    Write-Host "Enter the FQHN of the vCenter to Get Hosting Listing From: " -ForegroundColor "Yellow" -NoNewline
    $Global:VCName = Read-Host 
}
#*******************
# EndFunction Get-VC
#*******************

#*************************************************
# Check for Folder Structure if not present create
#*************************************************
Function Verify-Folders {
    [CmdletBinding()]
    Param()
    "Building Local folder structure" 
    If (!(Test-Path $Global:Folder)) {
        New-Item $Global:Folder -type Directory
        }
    "Folder Structure built" 
}
#***************************
# EndFunction Verify-Folders
#***************************

#*******************
# Connect to vCenter
#*******************
Function Connect-VC {
    [CmdletBinding()]
    Param()
    "Connecting to $Global:VCName"
    Connect-VIServer $Global:VCName -Credential $Global:Creds -WarningAction SilentlyContinue
}
#***********************
# EndFunction Connect-VC
#***********************

#*******************
# Disconnect vCenter
#*******************
Function Disconnect-VC {
    [CmdletBinding()]
    Param()
    "Disconnecting $Global:VCName"
    Disconnect-VIServer -Server $Global:VCName -Confirm:$false
}
#**************************
# EndFunction Disconnect-VC
#**************************


#*********************
# Clean Up after Run
#*********************
Function Clean-Up {
    [CmdletBinding()]
    Param()
    $Global:Folder = $null
    $Global:HostList = $null
    $Global:VCName = $null
    $Global:Creds = $null
}
#*********************
# EndFunction Clean-Up
#*********************

#********************************
# Function Get-WindowsDisk1Report
#********************************
Function Get-WindowsDisk1Report{
    [CmdletBinding()]
    Param()
    $results = @()
    Write-Host "Gathering List of Windows VMs"
    Write-Host "Be patient this may take some time ..."
    $WindowsVMList = get-view -ViewType VirtualMachine | Where {$_.Guest.GuestFullName -like "*Windows*"}
    Write-Host "List of Windows VMs Generated"
    $Count = 1
    Write-Host "Generating Disk Data..."
    Foreach ($vm in $WindowsVMList){
        $result = "" | Select vmName,vmPowerState,vmOS,DiskPath,CapacityGB,FreeSpaceGB,PercentFree
        $result.vmName =$vm.Name
        $result.vmPowerState = $vm.guest.gueststate
        $result.vmOS = $vm.guest.guestFullName
        $HDListing = $vm.Guest.Disk
        $HDCount = 1
        Foreach ($HD in $HDListing){
            If ($HD.DiskPath -eq "C:\"){
                $result.DiskPath = $HD.DiskPath
                $result.CapacityGB = [math]::Round($HD.Capacity / 1GB)
                $FreeSpace = ($HD.FreeSpace / 1GB)
                $result.FreeSpaceGB = [math]::Round($FreeSpace,2)
                $result.PercentFree = "{0:N0}" -f [math]::Round($HD.FreeSpace / $HD.Capacity * 100)  
            }
            $HDCount++
        }
        $results += $result
        $Count++
    }
    $results | Export-CSV -Path $Global:Folder\$Global:VCname-WindowsOSDisk-Report-$(Get-Date -Format yyyy-MM-dd).csv -NoTypeInformation
}
#***********************************
# EndFunction Get-WindowsDisk1Report
#***********************************

#**************************
# Function Convert-To-Excel
#**************************
Function Convert-To-Excel {
    [CmdletBinding()]
    Param()
   "Converting List from $Global:VCname to Excel"
    $workingdir = $Global:Folder+ "\*.csv"
    $csv = dir -path $workingdir

    foreach($inputCSV in $csv){
        $outputXLSX = $inputCSV.DirectoryName + "\" + $inputCSV.Basename + ".xlsx"
        ### Create a new Excel Workbook with one empty sheet
        $excel = New-Object -ComObject excel.application 
        $excel.DisplayAlerts = $False
        $workbook = $excel.Workbooks.Add(1)
        $worksheet = $workbook.worksheets.Item(1)

        ### Build the QueryTables.Add command
        ### QueryTables does the same as when clicking "Data » From Text" in Excel
        $TxtConnector = ("TEXT;" + $inputCSV)
        $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
        $query = $worksheet.QueryTables.item($Connector.name)


        ### Set the delimiter (, or ;) according to your regional settings
        ### $Excel.Application.International(3) = ,
        ### $Excel.Application.International(5) = ;
        $query.TextFileOtherDelimiter = $Excel.Application.International(5)

        ### Set the format to delimited and text for every column
        ### A trick to create an array of 2s is used with the preceding comma
        $query.TextFileParseType  = 1
        $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
        $query.AdjustColumnWidth = 1

        ### Execute & delete the import query
        $query.Refresh()
        $query.Delete()

        ### Get Size of Worksheet
        $objRange = $worksheet.UsedRange.Cells 
        $xRow = $objRange.SpecialCells(11).ow
        $xCol = $objRange.SpecialCells(11).column

        ### Format First Row
        $RangeToFormat = $worksheet.Range("1:1")
        $RangeToFormat.Style = 'Accent1'

        ### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
        $Workbook.SaveAs($outputXLSX,51)
        $excel.Quit()
    }
    ## To exclude an item, use the '-exclude' parameter (wildcards if needed)
    remove-item -path $workingdir 

}
#*****************************
# EndFunction Convert-To-Excel
#*****************************

#***************
# Execute Script
#***************

# Get Start Time
$startDTM = (Get-Date)

#CLS
$ErrorActionPreference="SilentlyContinue"

"=========================================================="
" "
Write-Host "Get CIHS credentials" -ForegroundColor Yellow
$Global:Creds = Get-Credential -Credential $null

Get-VCenter
Connect-VC
"----------------------------------------------------------"
Verify-Folders
"----------------------------------------------------------"
Get-WindowsDisk1Report
"----------------------------------------------------------"
Convert-To-Excel
"----------------------------------------------------------"
Disconnect-VC
"Open Explorer to $Global:Folder"
Invoke-Item $Global:Folder
Clean-Up

# Get End Time
$endDTM = (Get-Date)

# Echo Time elapsed
"Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"