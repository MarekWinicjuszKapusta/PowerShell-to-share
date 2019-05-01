#Script made by Marek Winicjusz Kapusta for internal use of Fujitsu Technology Solutions.
#i am totally not satisfied with this script.

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
    $form_HelloWorld.Text = "Application Manager"  
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
    $ListView.Columns.Add("Version") | Out-Null
    $ListView.Columns.Add("Vendor")  | Out-Null
    $ListView.Columns.Add("Install Date") | Out-Null
    $form_HelloWorld.controls.Add($ListView)
	[void]$ListView.AutoResizeColumns(1)
	
#ContextMenu for Listview
#-----------------------------------------------------------------------------------------------

$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip

    #Right click menu "Uninstall"
	$contextMenuStrip1.Items.Add("Uninstall", $ContextMenuStripItemImages).add_Click({
    
    #start Uninstall
    #-------------------------------------------------------------------------------------------------
    $form_HelloWorld.Text = "Loading..."
    $form_HelloWorld.ForeColor = "red" 
    $Selected = @($ListView.SelectedItems)
    $Selected.text |
    ForEach-Object {
    $app_name = $_
    $msgBoxInput =  [System.Windows.MessageBox]::Show("Do you want to uninstall '$app_name' on '$asset' ?",'Warning','YesNo','Error')

  	switch  ($msgBoxInput) {

  		'Yes' {

                $WMI = Get-WmiObject win32_product -ComputerName $asset -Filter "name = '$app_name'"
                $WMI.Uninstall()
                If($? -eq $false)
                {

                $wshell = New-Object -ComObject Wscript.Shell
                $error_msg = $error[0].Exception.Message
                $wshell.Popup("Uninstallation of $app_name Failed: $error_msg")

                }
    
                else
                {

                $wshell = New-Object -ComObject Wscript.Shell
                $wshell.Popup("Uninstallation of $app_name succeeded.")

                #Start Load
                #-------------------------------------------------------------------------------------------------

                #$form_HelloWorld.Text = "Loading..."
                #$form_HelloWorld.ForeColor = "red" 
        
                #clear content of columns
                $ListView.Items.Clear()

                $applications = Get-WmiObject win32_product -ComputerName $asset | Select Name,Version,Vendor,InstallDate
                $applications |
                ForEach-Object{

                If ($_.name -eq $null)
                {

                }
                else{
                $ListViewItem = New-Object System.Windows.Forms.ListViewItem($_.Name)
                $ListViewItem.Subitems.Add($_.Version) | Out-Null
                $ListViewItem.Subitems.Add($_.Vendor) | Out-Null
                $ListViewItem.Subitems.Add($_.InstallDate) | Out-Null
                $ListView.Items.Add($ListViewItem) | Out-Null
                }
                
                } 
     
                

                #End Load
                #-------------------------------------------------------------------------------------------------
                }
             
                
             #end of yes
             }
        #end of switch
        }

    }
    $form_HelloWorld.Text = "Application Manager"
    $form_HelloWorld.ForeColor = "black" 

    #End Uninstall
    #-------------------------------------------------------------------------------------------------
    })

    #start Out-GridView
    #-------------------------------------------------------------------------------------------------
    $contextMenuStrip1.Items.Add("OutGridView", $ContextMenuStripItemImages).add_Click({

            $Selected = @($ListView.Items) 
            $Selected |
            ForEach-Object{
            $a = $_.text
            $b = $_.subitems[1].text
            $c = $_.subitems[2].text
            $d = $_.subitems[3].text
            [pscustomobject]@{
              Name=$a
              Version=$b
              Vendor=$c
              InstallDate=$d
              }

    } | Out-GridView -Title "Installed applications on $asset"
    
    #end Out-GridView
    #-------------------------------------------------------------------------------------------------
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

#Button to load primary asset tag groups
$button_loada = New-Object System.Windows.Forms.Button  
    $button_loada.Location = New-Object System.Drawing.Size(188,8)  
    $button_loada.Size = New-Object System.Drawing.Size(58,32)  
    $button_loada.TextAlign = "MiddleCenter"  
    $button_loada.Text = "Load" 
    $button_loada.BackColor = "white"
    $button_loada.Add_Click({

        #Start Load
        #-------------------------------------------------------------------------------------------------

        $form_HelloWorld.Text = "Loading..."
        $form_HelloWorld.ForeColor = "red" 
        
        #clear content of columns
        $ListView.Items.Clear()

        $applications = Get-WmiObject win32_product -ComputerName $asset | Select Name,Version,Vendor,InstallDate
        $applications |
        ForEach-Object{

        If ($_.name -eq $null)
        {

        }
        else{
            $ListViewItem = New-Object System.Windows.Forms.ListViewItem($_.Name)
            $ListViewItem.Subitems.Add($_.Version) | Out-Null
            $ListViewItem.Subitems.Add($_.Vendor) | Out-Null
            $ListViewItem.Subitems.Add($_.InstallDate) | Out-Null
            $ListView.Items.Add($ListViewItem) | Out-Null
            }
        } 
     
     $form_HelloWorld.Text = "Application Manager"
     $form_HelloWorld.ForeColor = "black" 

     #End Load
     #-------------------------------------------------------------------------------------------------

    })  
    $form_HelloWorld.Controls.add($button_loada)


#The end
#-------------------------------------------------------------------------

#Show form  
$form_HelloWorld.Add_shown({$form_HelloWorld.Activate()})  
[void] $form_HelloWorld.ShowDialog() 
