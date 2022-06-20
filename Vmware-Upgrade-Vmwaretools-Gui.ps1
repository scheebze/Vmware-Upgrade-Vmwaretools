## --- Import VMWare Module --- ##
#Install-Module VMware.PowerCLI -confirm:$False 
#Import-Module VMware.PowerCLI

Stop-Transcript | out-null
$logspath = $PSScriptRoot + "\" + ("{0:yyyyMMdd}_UpgradeVmwareToolslog.log" -f (get-date))  
Start-Transcript -path $logspath -append

#FeatureToggles
$global:Feature_ManuallyChoose = $true

#Region Functions
#Import CSV Option Page
function FormGlobal-closeForm(){
    $Form_CSVOption.close()
  }

function FormCSVOption-ManuallyChoose(){

    if($Feature_ManuallyChoose -eq $False){
        $Form_CSVOption.controls.clear()
    }
    
    if($Feature_ManuallyChoose -eq $True){
        $Form_CSVOption.controls.AddRange(@(
            $Form_Global_closeButton,
            #$Form_ImportCSV_Label_Instructions,
            $Form_Global_Result,
            #$Form_ImportCSV_DownloadSampleButton,
            #$Form_ImportCSV_BrowseButton,
            #$Form_ImportCSV_TextBox_Browse,
            $Form_ImportCSV_Label_Username,
            $Form_ImportCSV_TextBox_UserName,
            $Form_ImportCSV_Label_Password,
            $Form_ImportCSV_TextBox_Password,
            $Form_ImportCSV_GetVMsButton))
    }

    $global:PageSelector = "ManuallyChoose"
    
}

function FormCSVOption-ImportCSV(){
    $Form_CSVOption.controls.clear()

    $Form_CSVOption.controls.AddRange(@(
        $Form_Global_closeButton,
        $Form_ImportCSV_Label_Instructions,
        $Form_Global_Result,
        $Form_ImportCSV_DownloadSampleButton,
        $Form_ImportCSV_BrowseButton,
        $Form_ImportCSV_TextBox_Browse,
        $Form_ImportCSV_Label_Username,
        $Form_ImportCSV_TextBox_UserName,
        $Form_ImportCSV_Label_Password,
        $Form_ImportCSV_TextBox_Password,
        $Form_ImportCSV_GetVMsButton))

    $global:PageSelector = "ImportCSV"

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
    #Remove items if present
    $Form_CSVOption.controls.remove($Form_ImportCSV_Label_UsernameandPasswordError) | out-null
    $Form_CSVOption.controls.remove($Form_ImportCSV_GetVMResultListview) | out-null
    $Form_ImportCSV_GetVMResultListview.clear()
    
    #Add Result box
    $Form_Global_Result.text = ""
    $Form_CSVOption.controls.add($Form_Global_Result) | out-null

    #If there is something in the browse text box do stuff
    if($Form_ImportCSV_TextBox_Browse.text -ne ""){
        #identify if username and password is filled out
        if($Form_ImportCSV_TextBox_UserName.text -ne "" -and $Form_ImportCSV_TextBox_Password.text -ne ""){
            
            #Import selected CSV File
            $Form_Global_Result.text = "Importing CSV file: $($Form_ImportCSV_TextBox_Browse.text)"
            try{
                $CSVContent = import-csv $Form_ImportCSV_TextBox_Browse.text -erroraction stop
                $Form_Global_Result.text += "`r`nImported $($csvcontent.count) items"
            }
            catch{
                $Form_Global_Result.text += "`r`nFailed to import CSV."
            }
            
            
            #Create Filter Search for getting VMs from Vcenter
            $importvmsearch = $csvcontent.servername
    
            #Set VCenters Array
            $VCenters = $csvcontent | select -Unique -ExpandProperty Vcenter 
            $Form_Global_Result.text += "`r`nIdentified $($Vcenters.count) VCenters"
            
            #Username Pulled from text box
            $username = $Form_ImportCSV_TextBox_UserName.text

            #Password Pulled from text box
            $password = $Form_ImportCSV_TextBox_Password.text
    
            ## -- connect to vcenter -- ##
            foreach($Vcenter in $VCenters){
                $Form_Global_Result.text += "`r`nConnecting to $vcenter"
                try{
                    connect-viserver -server $vcenter -username $username -password $password -force -erroraction stop
                }Catch{
                    $Form_Global_Result.text += "`r`n     Could not connect to $Vcenter."
                    $Form_Global_Result.text += "`r`n     Please check Vcenter Name is correct and the username and password used to login are correct."
                }
            }

            #If clause checks to see if a vcenter was able to be connected to using the global default vi server cmdlet
            if($global:defaultviserver){
                $Form_Global_Result.text += "`r`n"
                $Form_Global_Result.text += "`r`nGetting VMs from imported list..."
                #Get vms filtered on imported CSV. This is technically more efficient than pulling each vm individually... maybe...idk. 
                $vms = get-vm | where {$_.name -in $importvmsearch}

                $Form_Global_Result.text += "`r`nFound $($vms.count) of $($importvmsearch.count)"


                #Build Columns for list view
                $Form_ImportCSV_GetVMResultListview.columns.add("VMName") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VCenter") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("PowerState") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("IPAddress") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("Operating_System") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("ToolsRunningStatus") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VmwareTools_Version") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VmwareTools_InstallType") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VmwareTools_Status") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VmwareTools_VersionStatus") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VmwareTools_VersionStatus2") | out-null
                $Form_ImportCSV_GetVMResultListview.columns.add("VmwareTools_GuestFamily") | out-null

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
                    $VmwareTools_GuestFamily = @()
                    
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
                    $VmwareTools_GuestFamily = ($vm.guest.extensiondata.GuestFamily).tostring()

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
                    if(!$VmwareTools_GuestFamily){$VmwareTools_GuestFamily = "Null"}

                    #Add Listviewitems
                    $Form_ImportCSV_GetVMResultListview_Listviewitem = New-object System.Windows.Forms.ListViewItem($VMName)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($vcenter)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($PowerState)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($IPAddress)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($Operating_System)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($ToolsRunningStatus)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($VmwareTools_Version)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($VmwareTools_InstallType)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($VmwareTools_Status)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($VmwareTools_VersionStatus)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($VmwareTools_VersionStatus2)
                    $Form_ImportCSV_GetVMResultListview_Listviewitem.subitems.add($VmwareTools_GuestFamily)

                    $Form_ImportCSV_GetVMResultListview.items.add($Form_ImportCSV_GetVMResultListview_Listviewitem) | out-null
                    

                }

                $Form_ImportCSV_GetVMResultListview.AutoResizeColumns("HeaderSize")

                $Form_CSVOption.controls.remove($Form_Global_Result) | out-null
                $Form_CSVOption.controls.AddRange(@($Form_ImportCSV_GetVMResultListview,$Form_ImportCSV_UpgradeTools)) | out-null

            }

            #If clause checks to see if no vcenter was connected to. A bit of error handling...Kinda.
            if(!$global:defaultviserver){
                $Form_Global_Result.text += "`r`n "
                $Form_Global_Result.text += "`r`nDid not connect to any vcenter. Please try again. Double check the username and password used is correct. Make sure you are able to connect to the vcenter from your network."
            }
            #>
        }
        #if there is not username or password display error
        if($Form_ImportCSV_TextBox_UserName.text -eq "" -or $Form_ImportCSV_TextBox_Password.text -eq ""){
            $Form_CSVOption.controls.Add($Form_ImportCSV_Label_UsernameandPasswordError)
        }  

    }
}

function FormImportCSV-UpgradeTools(){
    $Form_CSVOption.controls.remove($Form_ImportCSV_GetVMResultListview)
    $Form_CSVOption.controls.Add($Form_ImportCSV_Label_UpgradeTracker)
    $Form_CSVOption.controls.Add($Form_ImportCSV_UpgradeVMwareToolsResultListview)

    #Build Columns for list view
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("VMName") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("VCenter") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("PowerState") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("VmwareTools_GuestFamily") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("SnapshotResult") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("SnapshotMessage") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("VmwareToolsUpgradeResult") | out-null
    $Form_ImportCSV_UpgradeVMwareToolsResultListview.columns.add("VmwareToolsUpgradeMessage") | out-null

    $SelectedVMs = @($Form_ImportCSV_GetVMResultListview.SelectedIndices)

    #Identify Indexes for properties
    $VcenterIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "Vcenter"}).index
    $VMnameIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VMName"}).index
    $PowerStateIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "PowerState"}).index
    $IPAddressIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "IPAddress"}).index
    $Operating_SystemIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "Operating_System"}).index
    $ToolsRunningStatusIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "ToolsRunningStatus"}).index
    $VmwareTools_VersionIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VmwareTools_Version"}).index
    $VmwareTools_InstallTypeIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VmwareTools_InstallType"}).index
    $VmwareTools_StatusIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VmwareTools_Status"}).index
    $VmwareTools_VersionStatusIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VmwareTools_VersionStatus"}).index
    $VmwareTools_VersionStatus2Index = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VmwareTools_VersionStatus2"}).index
    $VmwareTools_GuestFamilyIndex = ($Form_ImportCSV_GetVMResultListview.Columns | where {$_.text -eq "VmwareTools_GuestFamily"}).index

    #Counters
    $a = 1
    $b = $SelectedVMs.count

    $global:data = @()

    $SelectedVMs | foreach {
        
        #Clear Arrays for loop
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
        $VmwareTools_GuestFamily = @()
        $SnapshotResult = @()
        $SnapshotMessage = @()
        $VmwareToolsUpgradeResult = @()
        $VmwareToolsUpgradeMessage = @()

        #Set Property array for selected property in index
        $VMName = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VMNameIndex]).text
        $Vcenter = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VCenterIndex]).text
        $PowerState = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$PowerStateIndex]).text
        $IPAddress = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$IPAddressIndex]).text
        $Operating_System = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$Operating_SystemIndex]).text
        $ToolsRunningStatus = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$ToolsRunningStatusIndex]).text
        $VmwareTools_Version = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VmwareTools_VersionIndex]).text
        $VmwareTools_InstallType = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VmwareTools_InstallTypeIndex]).text
        $VmwareTools_Status = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VmwareTools_StatusIndex]).text
        $VmwareTools_VersionStatus = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VmwareTools_VersionStatusIndex]).text
        $VmwareTools_VersionStatus2 = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VmwareTools_VersionStatus2Index]).text
        $VmwareTools_GuestFamily = ($Form_ImportCSV_GetVMResultListview.items[$_].subitems[$VmwareTools_GuestFamilyIndex]).text

        $Form_ImportCSV_Label_UpgradeTracker.text = "Working on item $a of $b selected items."
        $Form_ImportCSV_Label_UpgradeTracker.text += "`r`nVM:$VMName on $vcenter."


        #Tools don't need upgrade = Skip
        if($PowerState -eq "Poweredon" -and $VmwareTools_VersionStatus -ne "guestToolsNeedUpgrade" -and $VmwareTools_GuestFamily -ne "windowsGuest"){
            $SnapshotResult = "Skipped"
            $SnapshotMessage = "Tools don't need upgrade"
            $VmwareToolsUpgradeResult = "Skipped"
            $VmwareToolsUpgradeMessage = "Tools don't need upgrade"
        }

        #Tools Status Handling: Tools not installed = Skip
        if($PowerState -eq "Poweredon" -and $VmwareTools_Status -eq "toolsNotInstalled" -and $VmwareTools_GuestFamily -ne "windowsGuest"){
            $SnapshotResult = "Skipped"
            $SnapshotMessage = "Tools not installed. Manually install Tools"
            $VmwareToolsUpgradeResult = "Skipped"
            $VmwareToolsUpgradeMessage = "Tools not installed. Manually install Tools"
        }

        #Guest Family Handling: not windows = skip
        if($PowerState -eq "Poweredon" -and $VmwareTools_GuestFamily -ne "windowsGuest"){
            $SnapshotResult = "Skipped"
            $SnapshotMessage = "Not a windows VM"
            $VmwareToolsUpgradeResult = "Skipped"
            $VmwareToolsUpgradeMessage = "Not a windows VM"

        }

        #Powerstate Handling: Off = Skip
        if($PowerState -ne "Poweredon"){
            $SnapshotResult = "Skipped"
            $SnapshotMessage = "VM is offline"
            $VmwareToolsUpgradeResult = "Skipped"
            $VmwareToolsUpgradeMessage = "VM is offline. Turn on VM and try again"

        }

        #Tools need Upgrade
        if($PowerState -eq "Poweredon" -and $VmwareTools_VersionStatus -eq "guestToolsNeedUpgrade" -and $VmwareTools_GuestFamily -eq "windowsGuest"){
            $Form_ImportCSV_Label_UpgradeTracker.text = "Initiating Snapshot on $vmname"
            
            #Take Snapshot of VM
            $task1 = get-vm "*$vmname*" -server $vcenter | new-snapshot -Name "Prior to Upgrading Vmware Tools" -Description "Snapshot taken prior to upgrading vmware tools." 
            $task1 | Wait-Task
            $task1 = Get-Task -Id $task1.Id
            if ($task1.State -eq "Success") {
                $SnapshotResult = "Completed"
                $SnapshotMessage = "Snapshot Succeeded"
                $Form_ImportCSV_Label_UpgradeTracker.text = "`r`nSnapshot Succeeded"

                #Upgrade VMware Tools
                $Form_ImportCSV_Label_UpgradeTracker.text = "Initiating VMware tools upgrade on $vmname"
                $task2 = Update-Tools "*$vmname*" -server $vcenter -NoReboot -RunAsync
                $task2 | Wait-Task
                $task2 = Get-Task -Id $task2.Id
                if ($task2.State -eq "Success") {
                    $VmwareToolsUpgradeResult = "Completed"
                    $VmwareToolsUpgradeMessage = "Upgrade VMTools successfully"
                    $Form_ImportCSV_Label_UpgradeTracker.text = "`r`nVMware Tools Upgrade Succeeded"
                } else {
                    $VmwareToolsUpgradeResult = "Failed"
                    $VmwareToolsUpgradeMessage = $task.ExtensionData.Info.Error.LocalizedMessage
                    $Form_ImportCSV_Label_UpgradeTracker.text = "`r`nVMware Tools Upgrade Failed"
                }

            } else {
                $SnapshotResult = "Failed"
                $SnapshotMessage = $task1.ExtensionData.Info.Error.LocalizedMessage
                $VmwareToolsUpgradeResult = "Aborted"
                $VmwareToolsUpgradeMessage = "Snapshot Failed. Aborted VMware tools upgrade"
                $Form_ImportCSV_Label_UpgradeTracker.text = "Failed to take snapshot for VM: $vmname"
                $Form_ImportCSV_Label_UpgradeTracker.text = "`r`nAborting VMware Tools Upgrade"
            }
        }

        if(!$SnapshotResult){$SnapshotResult = "Failed"}
        if(!$SnapshotMessage){$SnapshotMessage = "Value Ended Null"}
        if(!$VmwareToolsUpgradeResult){$VmwareToolsUpgradeResult = "Failed"}
        if(!$VmwareToolsUpgradeMessage){$VmwareToolsUpgradeMessage = "Value Ended Null"}


        #Add Listviewitems
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem = New-object System.Windows.Forms.ListViewItem($VMName)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($vcenter)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($PowerState)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($VmwareTools_GuestFamily)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($SnapshotResult)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($SnapshotMessage)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($VmwareToolsUpgradeResult)
        $Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem.subitems.add($VmwareToolsUpgradeMessage)

        $Form_ImportCSV_UpgradeVMwareToolsResultListview.items.add($Form_ImportCSV_UpgradeVMwareToolsResultListview_Listviewitem) | out-null

        #Build Export Report
        $row = New-Object PSObject
        $row | Add-Member -MemberType NoteProperty -Name "VMName" -Value $VMName
        $row | Add-Member -MemberType NoteProperty -Name "VCenter" -Value $vcenter
        $row | Add-Member -MemberType NoteProperty -Name "PowerState" -Value $PowerState
        $row | Add-Member -MemberType NoteProperty -Name "VmwareTools_GuestFamily" -Value $VmwareTools_GuestFamily
        $row | Add-Member -MemberType NoteProperty -Name "SnapshotResult" -Value $SnapshotResult
        $row | Add-Member -MemberType NoteProperty -Name "SnapshotMessage" -Value $SnapshotMessage
        $row | Add-Member -MemberType NoteProperty -Name "VmwareToolsUpgradeResult" -Value $VmwareToolsUpgradeResult
        $row | Add-Member -MemberType NoteProperty -Name "VmwareToolsUpgradeMessage" -Value $VmwareToolsUpgradeMessage
        $global:data += $row


        $a++
    }
    $Form_ImportCSV_Label_UpgradeTracker.text = "Completed Tasks for $b VMs."

    $Form_CSVOption.controls.remove($Form_ImportCSV_UpgradeTools)
    $Form_CSVOption.controls.Add($Form_ImportCSV_ExportResultsButton)

}

function FormCSV-ExportResults(){
    $Today = ((Get-Date).ToString('MMddyyyy'))
    $exportpath = $PSScriptRoot + "\$Today" + "_VmwareToolsUpgrade.csv"
    if($data){
        $data | export-csv -NoTypeInformation $exportpath
        $Form_ImportCSV_Label_UpgradeTracker.text = "Exported to $Exportpath"
    }else{
        $Form_ImportCSV_Label_UpgradeTracker.ForeColor    = "Red" 
        $Form_ImportCSV_Label_UpgradeTracker.text = "No Data to export."
    }
        
}

#Manually Choose Page



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

#Global Items
#Label

#Result Box

#Text Box

#Buttons

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
$Form_CSVOption_CloseButton.Add_Click({FormGlobal-closeForm})

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

$Form_ImportCSV_Label_UpgradeTracker              = New-Object system.Windows.Forms.Label
$Form_ImportCSV_Label_UpgradeTracker.text         = ""
$Form_ImportCSV_Label_UpgradeTracker.width        = 280
$Form_ImportCSV_Label_UpgradeTracker.height       = 30
$Form_ImportCSV_Label_UpgradeTracker.TextAlign    = "MiddleLeft"
$Form_ImportCSV_Label_UpgradeTracker.location     = New-Object System.Drawing.Point(320,300)
$Form_ImportCSV_Label_UpgradeTracker.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_Label_UpgradeTracker.ForeColor    = "Blue" 

#Result Box
$Form_Global_Result                          = New-Object system.Windows.Forms.TextBox
$Form_Global_Result.width                    = 720
$Form_Global_Result.height                   = 150
$Form_Global_Result.location                 = New-Object System.Drawing.Point(30,140)
$Form_Global_Result.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_Global_Result.multiline                = $True

$Form_ImportCSV_GetVMResultListview                          = New-Object system.Windows.Forms.listview
$Form_ImportCSV_GetVMResultListview.width                    = 720
$Form_ImportCSV_GetVMResultListview.height                   = 150
$Form_ImportCSV_GetVMResultListview.location                 = New-Object System.Drawing.Point(30,140)
$Form_ImportCSV_GetVMResultListview.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_GetVMResultListview.view                     = "Details"
$Form_ImportCSV_GetVMResultListview.FullRowSelect            = $True
$Form_ImportCSV_GetVMResultListview.MultiSelect              = $True
$Form_ImportCSV_GetVMResultListview.AllowColumnReorder       = $True
$Form_ImportCSV_GetVMResultListview.GridLines                = $True 
$Form_ImportCSV_GetVMResultListview.Sorting                  = "ascending"

$Form_ImportCSV_UpgradeVMwareToolsResultListview                          = New-Object system.Windows.Forms.listview
$Form_ImportCSV_UpgradeVMwareToolsResultListview.width                    = 720
$Form_ImportCSV_UpgradeVMwareToolsResultListview.height                   = 150
$Form_ImportCSV_UpgradeVMwareToolsResultListview.location                 = New-Object System.Drawing.Point(30,140)
$Form_ImportCSV_UpgradeVMwareToolsResultListview.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_UpgradeVMwareToolsResultListview.view                     = "Details"
$Form_ImportCSV_UpgradeVMwareToolsResultListview.FullRowSelect            = $True
$Form_ImportCSV_UpgradeVMwareToolsResultListview.MultiSelect              = $True
$Form_ImportCSV_UpgradeVMwareToolsResultListview.AllowColumnReorder       = $True
$Form_ImportCSV_UpgradeVMwareToolsResultListview.GridLines                = $True 
$Form_ImportCSV_UpgradeVMwareToolsResultListview.Sorting                  = "ascending"

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
$Form_Global_closeButton                     = New-Object system.Windows.Forms.Button
$Form_Global_closeButton.text                = "Close"
$Form_Global_closeButton.width               = 100
$Form_Global_closeButton.height              = 40
$Form_Global_closeButton.location            = New-Object System.Drawing.Point(30,300)
$Form_Global_closeButton.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_Global_closeButton.Add_Click({FormGlobal-closeForm})

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

$Form_ImportCSV_UpgradeTools                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_UpgradeTools.text            = "Upgrade VMware Tools"
$Form_ImportCSV_UpgradeTools.width           = 120
$Form_ImportCSV_UpgradeTools.height          = 40
$Form_ImportCSV_UpgradeTools.location        = New-Object System.Drawing.Point(630,300)
$Form_ImportCSV_UpgradeTools.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_UpgradeTools.Add_Click({FormImportCSV-UpgradeTools})

$Form_ImportCSV_ExportResultsButton                 = New-Object system.Windows.Forms.Button
$Form_ImportCSV_ExportResultsButton.text            = "Export Results"
$Form_ImportCSV_ExportResultsButton.width           = 100
$Form_ImportCSV_ExportResultsButton.height          = 40
$Form_ImportCSV_ExportResultsButton.location        = New-Object System.Drawing.Point(650,300)
$Form_ImportCSV_ExportResultsButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ImportCSV_ExportResultsButton.Add_Click({FormCSV-ExportResults})

#Manually Choose Page
#Label
$Form_ManuallyChoose_Label_Instructions              = New-Object system.Windows.Forms.Label
$Form_ManuallyChoose_Label_Instructions.text         = "Please Provide Vcenter Server, Username, and Password."
$Form_ManuallyChoose_Label_Instructions.width        = 300
$Form_ManuallyChoose_Label_Instructions.height       = 15
$Form_ManuallyChoose_Label_Instructions.TextAlign    = "MiddleLeft"
$Form_ManuallyChoose_Label_Instructions.location     = New-Object System.Drawing.Point(30,20)
$Form_ManuallyChoose_Label_Instructions.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ManuallyChoose_Label_Username              = New-Object system.Windows.Forms.Label
$Form_ManuallyChoose_Label_Username.text         = "Username:"
$Form_ManuallyChoose_Label_Username.width        = 75
$Form_ManuallyChoose_Label_Username.height       = 30
$Form_ManuallyChoose_Label_Username.TextAlign    = "MiddleLeft"
$Form_ManuallyChoose_Label_Username.location     = New-Object System.Drawing.Point(160,100)
$Form_ManuallyChoose_Label_Username.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ManuallyChoose_Label_Password              = New-Object system.Windows.Forms.Label
$Form_ManuallyChoose_Label_Password.text         = "Password:"
$Form_ManuallyChoose_Label_Password.width        = 70
$Form_ManuallyChoose_Label_Password.height       = 30
$Form_ManuallyChoose_Label_Password.TextAlign    = "MiddleLeft"
$Form_ManuallyChoose_Label_Password.location     = New-Object System.Drawing.Point(370,100)
$Form_ManuallyChoose_Label_Password.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ManuallyChoose_Label_UsernameandPasswordError              = New-Object system.Windows.Forms.Label
$Form_ManuallyChoose_Label_UsernameandPasswordError.text         = "Error! Please provide VCenter Server, username, and password!"
$Form_ManuallyChoose_Label_UsernameandPasswordError.width        = 400
$Form_ManuallyChoose_Label_UsernameandPasswordError.height       = 30
$Form_ManuallyChoose_Label_UsernameandPasswordError.TextAlign    = "MiddleLeft"
$Form_ManuallyChoose_Label_UsernameandPasswordError.location     = New-Object System.Drawing.Point(160,70)
$Form_ManuallyChoose_Label_UsernameandPasswordError.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ManuallyChoose_Label_UsernameandPasswordError.ForeColor    = "Red" 

#Result Box

#Text Box
$Form_ManuallyChoose_TextBox_UserName              = New-Object system.Windows.Forms.TextBox
$Form_ManuallyChoose_TextBox_UserName.text         = ""
$Form_ManuallyChoose_TextBox_UserName.width        = 120
$Form_ManuallyChoose_TextBox_UserName.height       = 30
$Form_ManuallyChoose_TextBox_UserName.location     = New-Object System.Drawing.Point(235,100)
$Form_ManuallyChoose_TextBox_UserName.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form_ManuallyChoose_TextBox_Password              = New-Object system.Windows.Forms.TextBox
$Form_ManuallyChoose_TextBox_Password.text         = ""
$Form_ManuallyChoose_TextBox_Password.width        = 120
$Form_ManuallyChoose_TextBox_Password.height       = 30
$Form_ManuallyChoose_TextBox_Password.location     = New-Object System.Drawing.Point(440,100)
$Form_ManuallyChoose_TextBox_Password.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Buttons
$Form_ManuallyChoose_GetVMsButton                 = New-Object system.Windows.Forms.Button
$Form_ManuallyChoose_GetVMsButton.text            = "Get-VMs"
$Form_ManuallyChoose_GetVMsButton.width           = 120
$Form_ManuallyChoose_GetVMsButton.height          = 40
$Form_ManuallyChoose_GetVMsButton.location        = New-Object System.Drawing.Point(30,90)
$Form_ManuallyChoose_GetVMsButton.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Form_ManuallyChoose_GetVMsButton.Add_Click({FormManuallyChoose-GetVMs})

#Add items to form
$Form_CSVOption.controls.AddRange(@(
    $Form_CSVOption_Label_Notice,
    $Form_Global_closeButton,
    $Form_CSVOption_ImportCSVButton,
    $Form_CSVOption_ManuallyChooseButton))

[void]$Form_CSVOption.ShowDialog()

#End Region Import CSV Option 

Disconnect-VIServer -Server $global:DefaultVIServers -confirm:$False -Force |out-null

Stop-Transcript | out-null
