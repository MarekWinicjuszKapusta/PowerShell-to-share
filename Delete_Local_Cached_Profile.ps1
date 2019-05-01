#Script made by Marek Winicjusz Kapusta for internal use of Fujitsu Technology Solutions.

#ver 1.3 Added Creator label
#ver 1.4 Added instruction, moved creator label a little.

#Load required assemblies  
[void] [system.reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
[void] [Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
#For MsgBoxInput
Add-Type -AssemblyName PresentationFramework

#List of variables

$asset = "Hostname"

#Drawing form and controls
#--------------------------------------------------------------------------------  
$form_HelloWorld = New-Object System.Windows.Forms.Form  
    $form_HelloWorld.Text = "Delete local cached profile"  
    $form_HelloWorld.size = New-Object System.Drawing.Size(1264,800)  
    $form_HelloWorld.FormBorderStyle = "FixedDialog"  
    $form_HelloWorld.TopMost = $false  
    $form_HelloWorld.MaximizeBox = $false  
    $form_HelloWorld.MinimizeBox = $true  
    $form_HelloWorld.ControlBox = $true  
    $form_HelloWorld.StartPosition = "CenterScreen"  
    $form_HelloWorld.Font = "Segoe UI"
    $form_HelloWorld.BackColor = "lightblue"

#Listview
#--------------------------------------------------------------------------------
$Global:ListView = New-Object System.Windows.Forms.ListView
    $ListView.Location = New-Object System.Drawing.Size(256,8)
    $ListView.Size = New-Object System.Drawing.Size(988,748)
    $ListView.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor
    [System.Windows.Forms.AnchorStyles]::Right -bor
    [System.Windows.Forms.AnchorStyles]::Top -bor
    [System.Windows.Forms.AnchorStyles]::Left
    $ListView.View = "Details"
    $ListView.FullRowSelect = $true
    $ListView.MultiSelect = $true
    $ListView.Sorting = "Ascending"
    $ListView.AllowColumnReorder = $true
    $ListView.GridLines = $true
    #$ListView.Add_ColumnClick({SortListView $_.Column})
    $ListView.Columns.Add("Name")  | Out-Null
    $ListView.Columns.Add("Loaded") | Out-Null
    $ListView.Columns.Add("Last Write Time")  | Out-Null
    $ListView.Columns.Add("Security ID") | Out-Null
    $form_HelloWorld.controls.Add($ListView)
	[void]$ListView.AutoResizeColumns(1)
	
#ContextMenu for Listview
#-----------------------------------------------------------------------------------------------

$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip

    #Right click menu "Delete Profile"
	$contextMenuStrip1.Items.Add("Delete Profile", $ContextMenuStripItemImages).add_Click({
    #start delete profile
    $form_HelloWorld.Text = "Loading..."
    $form_HelloWorld.ForeColor = "red" 
    $Selected = @($ListView.SelectedItems)
    $Selected.text |
    ForEach-Object {
    
    $profile_to_delete = Get-WmiObject -ComputerName $asset Win32_UserProfile -Filter "LocalPath like '%$_%'"
    
        If ( $profile_to_delete.loaded -eq $true )       
           {
           $wshell = New-Object -ComObject Wscript.Shell
           $wshell.Popup("Operation Failed: $_ is logged in")
           } 
          
        Elseif ( $profile_to_delete.loaded -eq $false )
           {
		   $msgBoxInput =  [System.Windows.MessageBox]::Show("I made sure that no data will be lost for $_ in the process of deleting local cached profile.",'Warning','YesNo','Error')

  				switch  ($msgBoxInput) {

  					'Yes' {

  					$profile_to_delete | Remove-WmiObject
                    If($? -eq $false)
                        {
                        $wshell = New-Object -ComObject Wscript.Shell
                        $wshell.Popup("Operation Failed: Unable to delete local cached profile for $_")
                        }
					#removing any leftovers
					remove-item -Path "//$asset/c$/users/$_" -Force -Recurse
                    #end of "yes"
  					}

                #end of the switch
  				}
           
           }

        Else
           {
           $wshell = New-Object -ComObject Wscript.Shell
           $wshell.Popup("Operation Failed: there is no local cached profile associated with '$_' folder in '//$asset/users' ")
           }

    }
    #end of deleting profiles
    #------------------------
    #start load

        #clear content of columns
        $ListView.Items.Clear()

        #get folders from C:/users
        $child = get-childitem -Path "//$asset/c$/users"
        $child |
        ForEach-Object{
        #a----------
        #get name of the folder from c:/users.
        $a = $_.name
        #b----------
        #find out if user is logged in. Get that info from local cached profile.
        $hook = Get-WmiObject -ComputerName "$asset" Win32_UserProfile -Filter "LocalPath like '%$a%'"
        $hook_loaded = $hook.loaded
        If($hook_loaded -eq $true)
        { $b = "True" }
        elseif($hook_loaded -eq $false) 
        { $b = "False" }
        else {$b = " " }
        #c----------
        #check when was the last time that folder was modified.
        $c = $_.lastwritetime | out-string
        #d----------
        #find security ID of local cached profile.
        $d = $hook.sid

        #creating Powershell custom object to list informations mentioned above on ListView
        $custom = [pscustomobject]@{
              Name=$a
              Loaded=$b
              LastWriteTime=$c
              SID=$d
        
           
        }
        #Creating items based on PSCustom Object
        $ListViewItem = New-Object System.Windows.Forms.ListViewItem($custom.name)
        $ListViewItem.Subitems.Add($custom.loaded) | Out-Null
        $ListViewItem.Subitems.Add($custom.lastwritetime) | Out-Null
        $ListViewItem.Subitems.Add($custom.SID) | Out-Null
        $ListView.Items.Add($ListViewItem) | Out-Null
    }  
    #end of load   

    $form_HelloWorld.Text = "Delete local cached profile"
    $form_HelloWorld.ForeColor = "black"  
    })

#end of Delete Profile right-click option
#---------------------------------------------------------------------------------------------------------------
#start of Delete Folder right-click option

    #Right click menu "Delete Folder"
	$contextMenuStrip1.Items.Add("Delete Folder", $ContextMenuStripItemImages).add_Click({
    #start delete folder
    $form_HelloWorld.Text = "Loading..."
    $form_HelloWorld.ForeColor = "red" 
    $Selected = @($ListView.SelectedItems)
    $Selected.text |
    ForEach-Object {
    
    $profile_to_delete = Get-WmiObject -ComputerName $asset Win32_UserProfile -Filter "LocalPath like '%$_%'"
    
        If ( $profile_to_delete.loaded -eq $true )       
           {
           $wshell = New-Object -ComObject Wscript.Shell
           $wshell.Popup("Operation Failed: There is still local cached profile associated with '$_' and user is logged in.")
           } 
          
        Elseif ( $profile_to_delete.loaded -eq $false )
           {
           $wshell = New-Object -ComObject Wscript.Shell
           $wshell.Popup("Operation Failed: There is still local cached profile associated with '$_'.")
           }

        Else
           {
           remove-item -Path "//$asset/c$/users/$_" -Force -Recurse
           If($? -eq $false)
                {
                $wshell = New-Object -ComObject Wscript.Shell
                $wshell.Popup("Operation Failed: Unable to delete '$_' folder.")
                }
           
           }

    }
    #end of deleting folders
    #------------------------
    #start load

        #clear content of columns
        $ListView.Items.Clear()

        #get folders from C:/users
        $child = get-childitem -Path "//$asset/c$/users"
        $child |
        ForEach-Object{
        #a----------
        #get name of the folder from c:/users.
        $a = $_.name
        #b----------
        #find out if user is logged in. Get that info from local cached profile.
        $hook = Get-WmiObject -ComputerName "$asset" Win32_UserProfile -Filter "LocalPath like '%$a%'"
        $hook_loaded = $hook.loaded
        If($hook_loaded -eq $true)
        { $b = "True" }
        elseif($hook_loaded -eq $false) 
        { $b = "False" }
        else {$b = " " }
        #c----------
        #check when was the last time that folder was modified.
        $c = $_.lastwritetime | out-string
        #d----------
        #find security ID of local cached profile.
        $d = $hook.sid

        #creating Powershell custom object to list informations mentioned above on ListView
        $custom = [pscustomobject]@{
              Name=$a
              Loaded=$b
              LastWriteTime=$c
              SID=$d
        
           
        }
        #Creating items based on PSCustom Object
        $ListViewItem = New-Object System.Windows.Forms.ListViewItem($custom.name)
        $ListViewItem.Subitems.Add($custom.loaded) | Out-Null
        $ListViewItem.Subitems.Add($custom.lastwritetime) | Out-Null
        $ListViewItem.Subitems.Add($custom.SID) | Out-Null
        $ListView.Items.Add($ListViewItem) | Out-Null
    }  
    #end of load   

    $form_HelloWorld.Text = "Delete local cached profile"
    $form_HelloWorld.ForeColor = "black"  
    })

    #end of options
    #--------------------------------------------------------------------------------
	$ListView.ContextMenuStrip = $contextMenuStrip1

#start of buttons & labels
#-------------------------------------------------------------------------------------


#Label Primary Asset  
$label_asset = New-Object System.Windows.Forms.Label  
    $label_asset.Location = New-Object System.Drawing.Size(8,8)  
    $label_asset.Size = New-Object System.Drawing.Size(116,32)  
    $label_asset.TextAlign = "MiddleCenter"  
    $label_asset.Text = "$asset"
    $form_HelloWorld.Controls.Add($label_asset)  

#Button to set Primary Asset
$button_asset = New-Object System.Windows.Forms.Button  
    $button_asset.Location = New-Object System.Drawing.Size(124,8)  
    $button_asset.Size = New-Object System.Drawing.Size(58,32)  
    $button_asset.TextAlign = "MiddleCenter"  
    $button_asset.Text = "Set"
    $button_asset.BackColor = "white"
    $button_asset.Add_Click({ $script:asset = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter asset tag", "Asset Tag")
    $script:asset1 = Get-ADPrincipalGroupMembership (Get-ADComputer $asset).DistinguishedName | Get-ADGroup -Properties name,description
    IF($?)
        {
        $label_asset.Text = "$asset"
        $label_asset.foreColor = "black"
        }
    else
        {
        $label_asset.Text = "$asset not found"
        $label_asset.foreColor = "red"
        }
    
    })  
    $form_HelloWorld.Controls.add($button_asset)

#Button to load local cached profiles
$button_loada = New-Object System.Windows.Forms.Button  
    $button_loada.Location = New-Object System.Drawing.Size(188,8)  
    $button_loada.Size = New-Object System.Drawing.Size(58,32)  
    $button_loada.TextAlign = "MiddleCenter"  
    $button_loada.Text = "Load" 
    $button_loada.BackColor = "white"
    $button_loada.Add_Click({
        
        #clear content of columns
        $ListView.Items.Clear()

        #get folders from C:/users
        $child = get-childitem -Path "//$asset/c$/users"
        $child |
        ForEach-Object{
        #a----------
        #get name of the folder from c:/users.
        $a = $_.name
        #b----------
        #find out if user is logged in. Get that info from local cached profile.
        $hook = Get-WmiObject -ComputerName "$asset" Win32_UserProfile -Filter "LocalPath like '%$a%'"
        $hook_loaded = $hook.loaded
        If($hook_loaded -eq $true)
        { $b = "True" }
        elseif($hook_loaded -eq $false) 
        { $b = "False" }
        else {$b = " " }
        #c----------
        #check when was the last time that folder was modified.
        $c = $_.lastwritetime | out-string
        #d----------
        #find security ID of local cached profile.
        $d = $hook.sid

        #creating Powershell custom object to list informations mentioned above on ListView
        $custom = [pscustomobject]@{
              Name=$a
              Loaded=$b
              LastWriteTime=$c
              SID=$d
        
           
        }
        #Creating items based on PSCustom Object
        $ListViewItem = New-Object System.Windows.Forms.ListViewItem($custom.name)
        $ListViewItem.Subitems.Add($custom.loaded) | Out-Null
        $ListViewItem.Subitems.Add($custom.lastwritetime) | Out-Null
        $ListViewItem.Subitems.Add($custom.SID) | Out-Null
        $ListView.Items.Add($ListViewItem) | Out-Null
    }         
     
    })  
    $form_HelloWorld.Controls.add($button_loada)

#Instruction
#-------------------------------------------------------------------------

#Label Instruction
$label_instruction = New-Object System.Windows.Forms.Label  
    $label_instruction.Location = New-Object System.Drawing.Size(8,44)  
    $label_instruction.Size = New-Object System.Drawing.Size(240,672)  
    $label_instruction.TextAlign = "MiddleCenter"  
    $label_instruction.BorderStyle = "fixed3d"
    $label_instruction.Text =
    "Instruction:
    1. Press 'Set'  button, and enter asset tag.
    If font changed to red, asset was not found in Active Directory.
    You can also enter IPv4, but font will also change to red.
    
    2. Press 'Load' to fill ListView window with profiles found on computer.
    
    3. Hover pointer over username of the profile that you want to delete, and perform right click mouse operation.
    
    4. Select 'Delete Profile'
    Profile can be deleted only if 'Loaded' is set to 'False', that means user was fully signed out.
    If user 'Loaded' status is set to 'True' even after sign out, restart machine, and then after user gets to logon screen, press 'Load' button once again.
    "  
    $form_HelloWorld.Controls.Add($label_instruction)

#creator label, last row
#-------------------------------------------------------------------------

#Label creator
$label_creator = New-Object System.Windows.Forms.Label  
    $label_creator.Location = New-Object System.Drawing.Size(8,724)  
    $label_creator.Size = New-Object System.Drawing.Size(240,32)  
    $label_creator.TextAlign = "MiddleCenter"  
    $label_creator.BorderStyle = "fixed3d"
    $label_creator.Text = "Creator: Marek Kapusta"  
    $form_HelloWorld.Controls.Add($label_creator)


#The end
#-------------------------------------------------------------------------

#Show form  
$form_HelloWorld.Add_shown({$form_HelloWorld.Activate()})  
[void] $form_HelloWorld.ShowDialog() 
