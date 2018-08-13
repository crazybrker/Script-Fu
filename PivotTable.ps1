########################################################################
# Just right-click and run with powershell
########################################################################

########################################################################
# ACAS to Retina and pivot table Generator
# Simple script, modify it as you need
########################################################################

#### MODIFY THIS!!! ####
# If you want to try NETBIOS before DNS set this to 1 (one) #
# If you want to try DNS before NETBIOS set this to 0 (zero) #
$NBNameBeforeDNS = 1;

### Maybe modify this ###
#If you want to filter out the -B-, -T- from the first sheet, this is used.
$filterIAV= "-[BT]-|N/A" #REGEX style, seperate with pipes

### No more mods below this point ###


#Setups
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

#Open Prompt
$openFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog1.DefaultExt = ".csv"
$openFileDialog1.FileName = ""
$openFileDialog1.ShowHelp = $True
$OpenFileDialog1.filter = "Scan Files (*.CSV)| *.csv"
echo "Select a scan file in the other window"
$status = $openFileDialog1.ShowDialog()

$LASTEXITCODE
if ($status -ne "OK") {break}

echo "Opening file, please wait..."
echo  $openFileDialog1.FileName

#Initialize
$outsheet = @()

$csv = Import-Csv $openFileDialog1.FileName

#Format ACAS to Retina output
$csv | %{

#IAVA and Netbios Cleanup
$IAV = $_.'Cross References' -split ',' -split '#' -match '-[BAT]-' | Out-String -Stream
$_.'NetBIOS Name' = $($_.'NetBIOS Name' -split '\\')[-1]

if ($_.'DNS Name' -match '\.') 
{$_.'DNS Name' = $($_.'DNS Name' -split '\.')[0] }

#Depending on preference....
if ($NBNameBeforeDNS){
    #Using NB for name frist, if thats blank then use DNS
    $name = $_.'NetBIOS Name'
    if (!$name){$name = $_.'DNS Name'}
}else{
    #Or else do the opposite
    $name = $_.'DNS Name'
    if (!$name){$name = $_.'NetBIOS Name'}
}




$entry = New-Object System.Object
$entry | Add-Member -NotePropertyName "IP" -NotePropertyValue $_.'IP Address'
$entry | Add-Member -NotePropertyName "NetBIOSName" -NotePropertyValue $name
$entry | Add-Member -NotePropertyName "MAC Address" -NotePropertyValue $_.'MAC Address'
$entry | Add-Member -NotePropertyName "IAV" -NotePropertyValue $IAV
$entry | Add-Member -NotePropertyName "Name" -NotePropertyValue $_.'Plugin Name'
$entry | Add-Member -NotePropertyName "Repository" -NotePropertyValue $_.'Repository'
$outsheet += $entry
}

#Output to retina CSV
$RetinaCSV = $openFileDialog1.FileName -replace ".csv" , " Retina.csv"
$outsheet | Sort-Object -Unique "IP" , "IAV" , "Name" | Export-Csv -NoTypeInformation -Path $RetinaCSV


echo "Preparing Pivot table"


        #### Setup Excel and workbook with 4 sheets 
        $excel = New-Object -ComObject excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $True
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.sheets.add()
    
        $csv = $excel.Workbooks.Open($RetinaCSV)


        #### Copy/Paste Import from scan.csv
        $items = $csv.worksheets.item(1)
        $items.UsedRange.select() > $null
        $items.UsedRange.copy() > $null

        $workbook.Activate()

        $RAW = $workbook.Sheets.Item(4)
        $RAW.Activate()
        $RAW.Paste()
        $RAW.Name = "RAW"

        $usedarea = [char](64 + $RAW.UsedRange.Columns.Count) + $RAW.UsedRange.Rows.Count

        #############################
        ### IAVA Pivot table ###
        #############################

        $PivotTable = $Workbook.PivotCaches().Create(1,"RAW!R1C1:R$($RAW.UsedRange.Rows.Count)C$($RAW.UsedRange.Columns.Count)",3)
        $PivotTable.CreatePivotTable("Sheet4!R4C1","Pivot") > $null
        $sheet = $workbook.Sheets.Item(1)
        $sheet.Activate()
        $sheet.Name = "By System IAVAs Only"

        $PivotFields = $Sheet.PivotTables("Pivot").PivotFields("NetBIOSName")
        $PivotFields.Orientation=1
        $PivotFields.Position=1

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("IP")
        $PivotFields.Orientation=1
        $PivotFields.Position=2

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("IAV")
        $PivotFields.Orientation=1
        $PivotFields.Position=3
        ### Filter out B patches and N/A
        $PivotFields.VisibleItems() | ?{$_.Name -match "$filterIAV"} | %{$PivotFields.PivotItems("$($_.Name)").visible = 0}

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("Name")
        $PivotFields.Orientation= 1
        $PivotFields.Position=4
        $PivotFields.Orientation= 4 #Add to values too

        #### Collapse NetBIOSName
        $PivotFields = $Sheet.PivotTables("Pivot").PivotFields("NetBIOSName")
        $PivotFields.ShowDetail = 0
        $PivotFields.AutoSort(2,"Count of Name")

        ############################
        ### All Vuln Pivot table ###
        ############################

        $PivotTable = $Workbook.PivotCaches().Create(1,"RAW!R1C1:R$($RAW.UsedRange.Rows.Count)C$($RAW.UsedRange.Columns.Count)",3)
        $PivotTable.CreatePivotTable("Sheet1!R4C1","Pivot") > $null
        $sheet = $workbook.Sheets.Item(2)
        $sheet.Activate()
        $sheet.Name = "By System ALL Vuln"

        $PivotFields = $Sheet.PivotTables("Pivot").PivotFields("NetBIOSName")
        $PivotFields.Orientation=1
        $PivotFields.Position=1

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("IP")
        $PivotFields.Orientation=1
        $PivotFields.Position=2

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("IAV")
        $PivotFields.Orientation=1
        $PivotFields.Position=3

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("Name")
        $PivotFields.Orientation= 1
        $PivotFields.Position=4
        $PivotFields.Orientation= 4 #Add to values too

        #### Collapse NetBIOSName
        $PivotFields = $Sheet.PivotTables("Pivot").PivotFields("NetBIOSName")
        $PivotFields.ShowDetail = 0
        $PivotFields.AutoSort(2,"Count of Name")

        ####################################
        ### By Vulnerability Pivot table ###
        ####################################

        $PivotTable = $Workbook.PivotCaches().Create(1,"RAW!R1C1:R$($RAW.UsedRange.Rows.Count)C$($RAW.UsedRange.Columns.Count)",3)
        $PivotTable.CreatePivotTable("Sheet2!R1C1","Pivot") > $null
        $sheet = $workbook.Sheets.Item(3)
        $sheet.Activate()
        $sheet.Name = "By IAVA"

        $PivotFields = $Sheet.PivotTables("Pivot").PivotFields("IAV")
        $PivotFields.Orientation=1
        $PivotFields.Position=1
        ### Filter for only CAT 1
        #$PivotFields.VisibleItems() | ?{$_.Name -notlike "Category I"} | %{$PivotFields.PivotItems("$($_.Name)").visible = 0}


        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("Name")
        $PivotFields.Orientation=1
        $PivotFields.Position=2

        $PivotFields = $Sheet.pivottables("Pivot").PivotFields("IP")
        $PivotFields.Orientation=1
        $PivotFields.Position=3
        $PivotFields.Orientation= 4 #Add to values too

        #### Collapse NetBIOSName
        $PivotFields = $Sheet.PivotTables("Pivot").PivotFields("IAV")
        $PivotFields.ShowDetail = 0
        $PivotFields.AutoSort(2,"IAV")


        #### Activate sheet 1
        $sheet = $workbook.Sheets.Item(1)
        $sheet.Activate()

        #### Close temp
        $csv.close()

        #Copy IAVA systems to clipboard
        Import-Csv $scanlist | ?{$_.IAV -match "-A-" } | Select-Object IP | Set-Clipboard


