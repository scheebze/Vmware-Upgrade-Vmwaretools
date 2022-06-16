## --- Import VMWare Module --- ##
#Import-Module VMware.PowerCLI

Stop-Transcript | out-null
$logspath = $PSScriptRoot + "\" + ("{0:yyyyMMdd}_UpgradeVmwareToolslog.log" -f (get-date))  
Start-Transcript -path $logspath -append

#Region Functions
#Import CSV Option Page
function FormCSVOption-closeForm(){
    $Form_CSVOption.close()
  }

function FormCSVOption-ManuallyChoose(){
    $Form_CSVOption_ManuallyChooseButton.text = "In Development"
}

function FormCSVOption-ImportCSV(){
    $Form_CSVOption.controls.clear()

    $Form_CSVOption.controls.AddRange(@(
        $Form_ImportCSV_closeButton,
        $Form_ImportCSV_Label_Instructions,
        $Form_ImportCSV_Result,
        #$Form_ImportCSV_SnapshotButton,
        $Form_ImportCSV_DownloadSampleButton,
        $Form_ImportCSV_BrowseButton,
        $Form_ImportCSV_TextBox_Browse,
        $Form_ImportCSV_ExportResultsButton,
        $Form_ImportCSV_Label_Username,
        $Form_ImportCSV_TextBox_UserName,
        $Form_ImportCSV_Label_Password,
        $Form_ImportCSV_TextBox_Password,
        $Form_ImportCSV_GetVMsButton))

}

#import CSV Page

function FormImportCSV-DownloadSample(){
    $exportpath = $PSScriptRoot + "\SampleCSV.csv"

    $sample = @()
    $Row1 = New-Object PSObject 
    $Row1 | Add-Member -MemberType NoteProperty -Name "ServerName" -Value "Server1"
    $Row1 | Add-Member -MemberType NoteProperty -Name "VCenter" -Value "Vcenter1.domain.local"
    $Row2 = New-Object PSObject 
    $Row2 | Add-Member -MemberType NoteProperty -Name "ServerName" -Value "Server2"
    $Row2 | Add-Member -MemberType NoteProperty -Name "VCenter" -Value "Vcenter2.domain.local"
    $Row3 = New-Object PSObject 
    $Row3 | Add-Member -MemberType NoteProperty -Name "ServerName" -Value "Server3"
    $Row3 | Add-Member -MemberType NoteProperty -Name "VCenter" -Value "Vcenter2.domain.local"
    $Row4 = New-Object PSObject 
    $Row4 | Add-Member -MemberType NoteProperty -Name "ServerName" -Value "Server4"
    $Row4 | Add-Member -MemberType NoteProperty -Name "VCenter" -Value "Vcenter1.domain.local"
    $sample += $Row1,$Row2, $Row3,$Row4

    $sample | export-csv -NoTypeInformation $exportpath
    
    $Form_ImportCSV_Label_DownloadSample.text = "$Exportpath"
    $Form_CSVOption.controls.Add($Form_ImportCSV_Label_DownloadSample)


}

function FormImportCSV-Browse(){
    $openfiledialog = New-Object System.Windows.Forms.OpenFileDialog
    $openfiledialog.Filter = "CSV Files (*.csv)|*.csv"
    if($openfiledialog.ShowDialog() -eq "OK"){
        $Form_ImportCSV_TextBox_Browse.text  = $openfiledialog.FileName
    }
}


function FormImportCSV-GetVMs(){
    $Form_CSVOption.controls.remove($Form_ImportCSV_Label_UsernameandPasswordError) | out-null
    if($Form_ImportCSV_TextBox_Browse.text -ne ""){
        #identify if username and password is filled out
        if($Form_ImportCSV_TextBox_UserName.text -ne "" -and $Form_ImportCSV_TextBox_Password.text -ne ""){
            
            #Import selected CSV File
            $Form_ImportCSV_Result.text = "Importing CSV file: $($Form_ImportCSV_TextBox_Browse.text)"
            try{
                $CSVContent = import-csv $Form_ImportCSV_TextBox_Browse.text -erroraction stop
                $Form_ImportCSV_Result.text += "`r`nImported $($csvcontent.count) items"
            }
            catch{
                $Form_ImportCSV_Result.text += "`r`nFailed to import CSV."
            }
            
            
            #Create Filter Search for getting VMs from Vcenter
            $importvmsearch = $csvcontent.servername
    
            #Set VCenters Array
            $VCenters = $csvcontent | select -Unique -ExpandProperty Vcenter 
            $Form_ImportCSV_Result.text += "`r`nIdentified $($Vcenters.count) VCenters"
                
            $username = $Form_ImportCSV_TextBox_UserName.text

            $password = $Form_ImportCSV_TextBox_Password.text
    
            ## -- connect to vcenter -- ##
            foreach($Vcenter in $VCenters){
                $Form_ImportCSV_Result.text += "`r`nConnecting to $vcenter"
                try{
                    connect-viserver -server $vcenter -username $username -password $password -force -erroraction stop
                }Catch{
                    $Form_ImportCSV_Result.text += "`r`n     Could not connect to $Vcenter."
                    $Form_ImportCSV_Result.text += "`r`n     Please check Vcenter Name is correct and the username and password used to login are correct."
                }
            }

            if($global:defaultviserver){
                $Form_ImportCSV_Result.text += "`r`n"
                $Form_ImportCSV_Result.text += "`r`nGetting VMs from imported list..."
                #Get vms filtered on imported CSV. This is technically more efficient than pulling each vm individually... maybe...idk. 
                $vms = get-vm | where {$_.name -in $importvmsearch}

                $Form_ImportCSV_Result.text += "`r`nFound $($vms.count) of $($importvmsearch.count)"


                #Build Columns for list view
                $Form_ImportCSV_ResultListview.columns.add("VMName") | out-null
                $Form_ImportCSV_ResultListview.columns.add("VCenter") | out-null
                $Form_ImportCSV_ResultListview.columns.add("PowerState") | out-null
                $Form_ImportCSV_ResultListview.columns.add("IPAddress") | out-null
                $Form_ImportCSV_ResultListview.columns.add("Operating_System") | out-null
                $Form_ImportCSV_ResultListview.columns.add("ToolsRunningStatus") | out-null
                $Form_ImportCSV_ResultListview.columns.add("VmwareTools_Version") | out-null
                $Form_ImportCSV_ResultListview.columns.add("VmwareTools_InstallType") | out-null
                $Form_ImportCSV_ResultListview.columns.add("VmwareTools_Status") | out-null
                $Form_ImportCSV_ResultListview.columns.add("VmwareTools_VersionStatus") | out-null
                $Form_ImportCSV_ResultListview.columns.add("VmwareTools_VersionStatus2") | out-null

                #loop through each VM
                foreach ($vm in $vms){
                    #Empty Values
                    $Vcenter = @()
                    $VMname = @()
                    $PowerState = @()
                    $IPAddress = @()
                    $Operating_System = @()
                    $ToolsRunningStatus = @()
                    $VmwareTools_Version = @()
                    $VmwareTools_InstallType = @()
                    $VmwareTools_Status = @()
                    $VmwareTools_VersionStatus = @()
                    $VmwareTools_VersionStatus2 = @()
                    
                    #Set Variables
                    $trash1,$vcenter,$trash2 = ($vm.guest.vmuid -split "@") -split ":"
                    $VMname = ($VM.Name).tostring()
                    $PowerState = ($vm.powerstate).tostring()
                    $IPAddress = ($vm.guest.extensiondata.IpAddress).tostring()
                    $Operating_System = ($vm.guest.extensiondata.GuestFullName).tostring()
                    $ToolsRunningStatus = ($vm.guest.extensiondata.ToolsRunningStatus).tostring()
                    $VmwareTools_Version = ($vm.guest.extensiondata.ToolsVersion).tostring()
                    $VmwareTools_InstallType = ($vm.guest.extensiondata.ToolsInstallType).tostring()
                    $VmwareTools_Status = ($vm.guest.extensiondata.ToolsStatus).tostring()
                    $VmwareTools_VersionStatus = ($vm.guest.extensiondata.ToolsVersionStatus).tostring()
                    $VmwareTools_VersionStatus2 = ($vm.guest.extensiondata.ToolsVersionStatus2).tostring()

                    #Handle Null Values
                    if(!$VMname){$VMname = "Null"}
                    if(!$PowerState){$PowerState = "Null"}
                    if(!$IPAddress){$IPAddress = "Null"}
                    if(!$Operating_System){$Operating_System = "Null"}
                    if(!$ToolsRunningStatus){$ToolsRunningStatus = "Null"}
                    if(!$VmwareTools_Version){$VmwareTools_Version = "Null"}
                    if(!$VmwareTools_InstallType){$VmwareTools_InstallType = "Null"}
                    if(!$VmwareTools_Status){$VmwareTools_Status = "Null"}
                    if(!$VmwareTools_VersionStatus){$VmwareTools_VersionStatus = "Null"}
                    if(!$VmwareTools_VersionStatus2){$VmwareTools_VersionStatus2 = "Null"}

                    #Add Listviewitems
                    $Form_ImportCSV_ResultListview_Listviewitem = New-object System.Windows.Forms.ListViewItem($VMName)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($vcenter)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($PowerState)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($IPAddress)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($Operating_System)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($ToolsRunningStatus)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($VmwareTools_Version)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($VmwareTools_InstallType)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($VmwareTools_Status)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($VmwareTools_VersionStatus)
                    $Form_ImportCSV_ResultListview_Listviewitem.subitems.add($VmwareTools_VersionStatus2)

                    $Form_ImportCSV_ResultListview.items.add($Form_ImportCSV_ResultListview_Listviewitem) | out-null
                    

                }

                $Form_ImportCSV_ResultListview.AutoResizeColumns("HeaderSize")

                $Form_CSVOption.controls.remove($Form_ImportCSV_Result) | out-null
                $Form_CSVOption.controls.AddRange($Form_ImportCSV_ResultListview,$Form_ImportCSV_SnapshotButton) | out-null

            }
            #>
        }
        if($Form_ImportCSV_TextBox_UserName.text -eq "" -or $Form_ImportCSV_TextBox_Password.text -eq ""){
            $Form_CSVOption.controls.Add($Form_ImportCSV_Label_UsernameandPasswordError)
        }  

    }
}

function FormImportCSV-Snapshot(){

    }
     
function FormImportCSV-ExportResults(){
    $Today = ((Get-Date).ToString('MMddyyyy'))
    $exportpath = $PSScriptRoot + "\$Today" + "_VmwareToolsUpgrade.log"
    if($Form_ImportCSV_Result.text -ne "*" -or $Form_ImportCSV_Result.text -eq "No data to export."){
        $Form_ImportCSV_Result.text = "No data to export."
    }
    if($Form_ImportCSV_Result.text -like "*" -and $Form_ImportCSV_Result.text -ne "No data to export."){
        $Form_ImportCSV_Result.text | Out-File $exportpath -Append
    }

    # Check to see if the column headers on the CSV is the correct format
    
}

#endRegion Functions

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Region Import CSV Option
#Build Form
$Form_CSVOption                                = New-Object system.Windows.Forms.Form
$Form_CSVOption.Size                           = New-Object System.Drawing.Size(800,400)
$Form_CSVOption.text                           = "Upgrade VMware Tools"
$Form_CSVOption.TopMost                        = $True
$Form_CSVOption.StartPosition                  = 'CenterScreen'
$Form_CSVOption.FormBorderStyle                = "FixedDialog"
$Form_CSVOption.MaximizeBox                    = $false
$Form_CSVOption.MinimizeBox                    = $True
$Form_CSVOption.ControlBox                     = $True

#Import CSV Option Page
#Label
$Form_CSVOption_Label_Notice            = New-Object system.Windows.Forms.Label
$Form_CSVOption_Label_Notice.text       = "Notice: Manually Choose Server is currently in development."
$Form_CSVOption_Label_Notice.width      = 400
$Form_CSVOption_Label_Notice.height     = 60
$Form_CSVOption_Label_Notice.location   = New-Object System.Drawing.Point(200,125)
$Form_CSVOption_Label_Notice.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',14)

#Buttons
$Form_CSVOption_CloseButton                     = New-Object system.Windows.Forms.Button
$Form_CSVOption_CloseButton.text                = "Close"
$Form_CSVOption_CloseButton.width               = 100
$Form_CSVOption_CloseButton.height              = 40
$Form_CSVOption_CloseButton.location            = New-Object System.Drawing.Point(30,300)
$Form_CSVOption_CloseButton.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_CSVOption_CloseButton.Add_Click({FormCSVOption-closeForm})

$Form_CSVOption_ImportCSVButton                 = New-Object system.Windows.Forms.Button
$Form_CSVOption_ImportCSVButton.text            = "Import from CSV"
$Form_CSVOption_ImportCSVButton.width           = 150
$Form_CSVOption_ImportCSVButton.height          = 60
$Form_CSVOption_ImportCSVButton.location        = New-Object System.Drawing.Point(30,30)
$Form_CSVOption_ImportCSVButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_CSVOption_ImportCSVButton.Add_Click({FormCSVOption-ImportCSV})

$Form_CSVOption_ManuallyChooseButton            = New-Object system.Windows.Forms.Button
$Form_CSVOption_ManuallyChooseButton.text       = "Manually Choose Server"
$Form_CSVOption_ManuallyChooseButton.width      = 150
$Form_CSVOption_ManuallyChooseButton.height     = 60
$Form_CSVOption_ManuallyChooseButton.location   = New-Object System.Drawing.Point(30,120)
$Form_CSVOption_ManuallyChooseButton.Font       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_CSVOption_ManuallyChooseButton.Add_Click({FormCSVOption-ManuallyChoose})

#Import CSV Page
#Label
$Form_ImportCSV_Label_Instructions              = New-Object system.Windows.Forms.Label
$Form_ImportCSV_Label_Instructions.text         = "Please select a CSV from your local machine."
$Form_ImportCSV_Label_Instructions.width        = 300
$Form_ImportCSV_Label_Instructions.height       = 15
$Form_ImportCSV_Label_Instructions.TextAlign    = "MiddleLeft"
$Form_ImportCSV_Label_Instructions.location     = New-Object System.Drawing.Point(30,20)
$Form_ImportCSV_Label_Instructions.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ImportCSV_Label_DownloadSample              = New-Object system.Windows.Forms.Label
$Form_ImportCSV_Label_DownloadSample.text         = ""
$Form_ImportCSV_Label_DownloadSample.width        = 300
$Form_ImportCSV_Label_DownloadSample.height       = 15
$Form_ImportCSV_Label_DownloadSample.TextAlign    = "MiddleLeft"
$Form_ImportCSV_Label_DownloadSample.location     = New-Object System.Drawing.Point(320,310)
$Form_ImportCSV_Label_DownloadSample.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ImportCSV_Label_Username              = New-Object system.Windows.Forms.Label
$Form_ImportCSV_Label_Username.text         = "Username:"
$Form_ImportCSV_Label_Username.width        = 75
$Form_ImportCSV_Label_Username.height       = 30
$Form_ImportCSV_Label_Username.TextAlign    = "MiddleLeft"
$Form_ImportCSV_Label_Username.location     = New-Object System.Drawing.Point(160,100)
$Form_ImportCSV_Label_Username.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ImportCSV_Label_Password              = New-Object system.Windows.Forms.Label
$Form_ImportCSV_Label_Password.text         = "Password:"
$Form_ImportCSV_Label_Password.width        = 70
$Form_ImportCSV_Label_Password.height       = 30
$Form_ImportCSV_Label_Password.TextAlign    = "MiddleLeft"
$Form_ImportCSV_Label_Password.location     = New-Object System.Drawing.Point(370,100)
$Form_ImportCSV_Label_Password.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ImportCSV_Label_UsernameandPasswordError              = New-Object system.Windows.Forms.Label
$Form_ImportCSV_Label_UsernameandPasswordError.text         = "Error! Please provide username and password!"
$Form_ImportCSV_Label_UsernameandPasswordError.width        = 400
$Form_ImportCSV_Label_UsernameandPasswordError.height       = 30
$Form_ImportCSV_Label_UsernameandPasswordError.TextAlign    = "MiddleLeft"
$Form_ImportCSV_Label_UsernameandPasswordError.location     = New-Object System.Drawing.Point(160,70)
$Form_ImportCSV_Label_UsernameandPasswordError.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_Label_UsernameandPasswordError.ForeColor    = "Red" 

#Result Box
$Form_ImportCSV_Result                          = New-Object system.Windows.Forms.TextBox
$Form_ImportCSV_Result.width                    = 720
$Form_ImportCSV_Result.height                   = 150
$Form_ImportCSV_Result.location                 = New-Object System.Drawing.Point(30,140)
$Form_ImportCSV_Result.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_Result.multiline                = $True

$Form_ImportCSV_ResultListview                          = New-Object system.Windows.Forms.listview
$Form_ImportCSV_ResultListview.width                    = 720
$Form_ImportCSV_ResultListview.height                   = 150
$Form_ImportCSV_ResultListview.location                 = New-Object System.Drawing.Point(30,140)
$Form_ImportCSV_ResultListview.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_ResultListview.view                     = "Details"
$Form_ImportCSV_ResultListview.FullRowSelect            = $True
$Form_ImportCSV_ResultListview.MultiSelect              = $True
$Form_ImportCSV_ResultListview.AllowColumnReorder       = $True
$Form_ImportCSV_ResultListview.GridLines                = $True 
$Form_ImportCSV_ResultListview.Sorting                  = "ascending"

#Text Box
$Form_ImportCSV_TextBox_Browse              = New-Object system.Windows.Forms.TextBox
$Form_ImportCSV_TextBox_Browse.text         = ""
$Form_ImportCSV_TextBox_Browse.width        = 400
$Form_ImportCSV_TextBox_Browse.height       = 50
$Form_ImportCSV_TextBox_Browse.location     = New-Object System.Drawing.Point(150,47)
$Form_ImportCSV_TextBox_Browse.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ImportCSV_TextBox_UserName              = New-Object system.Windows.Forms.TextBox
$Form_ImportCSV_TextBox_UserName.text         = ""
$Form_ImportCSV_TextBox_UserName.width        = 120
$Form_ImportCSV_TextBox_UserName.height       = 30
$Form_ImportCSV_TextBox_UserName.location     = New-Object System.Drawing.Point(235,100)
$Form_ImportCSV_TextBox_UserName.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ImportCSV_TextBox_Password              = New-Object system.Windows.Forms.TextBox
$Form_ImportCSV_TextBox_Password.text         = ""
$Form_ImportCSV_TextBox_Password.width        = 120
$Form_ImportCSV_TextBox_Password.height       = 30
$Form_ImportCSV_TextBox_Password.location     = New-Object System.Drawing.Point(440,100)
$Form_ImportCSV_TextBox_Password.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Buttons
$Form_ImportCSV_closeButton                     = New-Object system.Windows.Forms.Button
$Form_ImportCSV_closeButton.text                = "Close"
$Form_ImportCSV_closeButton.width               = 100
$Form_ImportCSV_closeButton.height              = 40
$Form_ImportCSV_closeButton.location            = New-Object System.Drawing.Point(30,300)
$Form_ImportCSV_closeButton.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_closeButton.Add_Click({FormCSVOption-closeForm})

$Form_ImportCSV_DownloadSampleButton                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_DownloadSampleButton.text            = "Download Sample CSV"
$Form_ImportCSV_DownloadSampleButton.width           = 150
$Form_ImportCSV_DownloadSampleButton.height          = 40
$Form_ImportCSV_DownloadSampleButton.location        = New-Object System.Drawing.Point(160,300)
$Form_ImportCSV_DownloadSampleButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_DownloadSampleButton.Add_Click({FormImportCSV-DownloadSample})

$Form_ImportCSV_BrowseButton                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_BrowseButton.text            = "Browse..."
$Form_ImportCSV_BrowseButton.width           = 100
$Form_ImportCSV_BrowseButton.height          = 40
$Form_ImportCSV_BrowseButton.location        = New-Object System.Drawing.Point(30,40)
$Form_ImportCSV_BrowseButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_BrowseButton.Add_Click({FormImportCSV-Browse})

$Form_ImportCSV_GetVMsButton                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_GetVMsButton.text            = "Import CSV"
$Form_ImportCSV_GetVMsButton.width           = 120
$Form_ImportCSV_GetVMsButton.height          = 40
$Form_ImportCSV_GetVMsButton.location        = New-Object System.Drawing.Point(30,90)
$Form_ImportCSV_GetVMsButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_GetVMsButton.Add_Click({FormImportCSV-GetVMs})

$Form_ImportCSV_SnapshotButton                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_SnapshotButton.text            = "Upgrade VMware Tools"
$Form_ImportCSV_SnapshotButton.width           = 120
$Form_ImportCSV_SnapshotButton.height          = 40
$Form_ImportCSV_SnapshotButton.location        = New-Object System.Drawing.Point(630,90)
$Form_ImportCSV_SnapshotButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_SnapshotButton.Add_Click({FormImportCSV-Snapshot})

$Form_ImportCSV_ExportResultsButton                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_ExportResultsButton.text            = "Export Results"
$Form_ImportCSV_ExportResultsButton.width           = 100
$Form_ImportCSV_ExportResultsButton.height          = 40
$Form_ImportCSV_ExportResultsButton.location        = New-Object System.Drawing.Point(650,300)
$Form_ImportCSV_ExportResultsButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_ExportResultsButton.Add_Click({FormImportCSV-ExportResults})

#Add items to form
$Form_CSVOption.controls.AddRange(@(
    $Form_CSVOption_Label_Notice,
    $Form_CSVOption_CloseButton,
    $Form_CSVOption_ImportCSVButton,
    $Form_CSVOption_ManuallyChooseButton))

[void]$Form_CSVOption.ShowDialog()

#End Region Import CSV Option 


<#
## -- Declare Vcenter Servers
$vcenters = @("vcenter1.domain.local","vcenter2.domain.local")

## -- Get server list of VMs -- ##
$vms = get-content C:\temp\vms.txt

## -- connect to vcenter -- ##
foreach($Vcenter in $Vcenters){
    Write-host "Connecting to $Vcenter" -ForegroundColor Yellow
    connect-viserver -server $vcenter -credential $creds -force
}

#Run Through Each VM and take snapshot and upgrade VM
foreach ($vm in $vms){
    write-host "Working on $vm" -ForegroundColor Yellow

    #Take Snapshot 
    Write-host "Taking Snapshot"
    get-vm "*$vm*" | new-snapshot -Name "Prior to Upgrading Vmware Tools" -Description "Snapshot taken prior to upgrading vmware tools $today." 

    #Update Vmware Tools
    Write-host "Updating Vmware Tools"
    Update-Tools "*$vm*"  -NoReboot -RunAsync
}
#>

Stop-Transcript | out-null
