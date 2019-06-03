#Script made by Marek Winicjusz Kapusta
#Main purpose of this script is to compare Active Directory objects and find the differences
#Script is easy to use because of introduction of Graphical User Interface

#ver 1.1 added Out-Gridview button
#ver 1.2 added GPRSoP
#ver 1.3 added popup windows prompt if operations fails for RSoP
#ver 1.4 added add / remove group button in row 9
#ver 1.5 moved primary, seconadry, asset1, asset2 variables to load buttons
#ver 1.6 changed positions of administrator functions, added add/remobe group function to primary asset
#ver 1.7 Added promp if groups were added successfully, same for RSoP.
#ver 1.8 Added Row 12, Members of the group
#ver 1.9 Added right click menu with copy, export, out-gridview
#ver 2.0 Added creator label
#ver 2.1 moved creator label a little



#Load required assemblies  
[void] [system.reflection.assembly]::LoadWithPartialName("System.Windows.Forms")  
[void] [Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')  

#List of variables
$user = "Primary User"
$equivalent = "Equivalent User" 
$item
$path
$asset = "Primary Computer"
$easset = "Equivalent Computer" 
[int]$counter = 1
$gridviewtitle = "GridView"

#Drawing form and controls
#--------------------------------------------------------------------------------  
$form_HelloWorld = New-Object System.Windows.Forms.Form  
    $form_HelloWorld.Text = "AD Compare"  
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
    $ListView.Sorting = "None"
    $ListView.AllowColumnReorder = $true
    $ListView.GridLines = $true
    $ListView.Add_ColumnClick({SortListView $_.Column})
    $ListView.Columns.Add("Group Name")  | Out-Null
    $ListView.Columns.Add("Description") | Out-Null
    $form_HelloWorld.controls.Add($ListView)
	[void]$ListView.AutoResizeColumns(1)
	
#ContextMenu for Listview
#-----------------------------------------------------------------------------------------------

$contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip

    #Right click menu "Copy"
	$contextMenuStrip1.Items.Add("Copy", $ContextMenuStripItemImages).add_Click({
 
    $Selected = @($ListView.SelectedItems)
    $Selected.text |
    ForEach-Object {
    If($counter -eq 1)
        {$clip = $clip + $_
        [int]$counter = $counter + 1
        }
    else {$clip = $clip + ", " + $_}
        }

        $clip | clip

    })

    ##Right click menu "Export"
    $contextMenuStrip1.Items.Add("Export", $ContextMenuStripItemImages).add_Click({ 
    
        $script:path = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter location and name of the file. Example: C:\Users\username\Desktop\export.csv", "please enter location")
        $Selected = @($ListView.SelectedItems) 
        $Selected |
        ForEach-Object{
        $a = $_.text
        $b = $_.subitems[1].text
        [pscustomobject]@{
              Name=$a
              Description=$b
              
              }

        } | Export-Csv -Path "$path" -NoTypeInformation -Append

    })


    #Right click menu "Out-GridView"
    $contextMenuStrip1.Items.Add("Out-GridView", $ContextMenuStripItemImages).add_Click({ 
    
        $Selected = @($ListView.Items) 
        $Selected |
        ForEach-Object{
        $a = $_.text
        $b = $_.subitems[1].text
        [pscustomobject]@{
              Name=$a
              Description=$b
              
              }

    } | Out-GridView -Title $gridviewtitle
    
    })

    #end of options
    #--------------------------------------------------------------------------------
	$ListView.ContextMenuStrip = $contextMenuStrip1

#Row 1
#-------------------------------------------------------------------------

#Label Primary User  
$label_primary_user = New-Object System.Windows.Forms.Label  
    $label_primary_user.Location = New-Object System.Drawing.Size(8,8)  
    $label_primary_user.Size = New-Object System.Drawing.Size(116,32)  
    $label_primary_user.TextAlign = "MiddleCenter"  
    $label_primary_user.Text = "$user" 
    $form_HelloWorld.Controls.Add($label_primary_user)

#Button to set Primary User  
$button_primary_user = New-Object System.Windows.Forms.Button  
    $button_primary_user.Location = New-Object System.Drawing.Size(124,8)  
    $button_primary_user.Size = New-Object System.Drawing.Size(58,32)  
    $button_primary_user.TextAlign = "MiddleCenter"  
    $button_primary_user.Text = "Set"  
    $button_primary_user.backcolor = "white"
    $button_primary_user.Add_Click({ $script:user = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter username", "Primary User")   
    $script:primary = Get-ADPrincipalGroupMembership $user | Get-ADGroup -Properties name,description
    If($?)
        {
        $label_primary_user.Text = "$user"
        $label_primary_user.foreColor = "black"
        }
    else
        {
        $label_primary_user.Text = "$user not found"
        $label_primary_user.foreColor = "red"
        }
    
    })  
    $form_HelloWorld.Controls.add($button_primary_user)

#Button to Load Primary user groups  
$button_load_primary_user = New-Object System.Windows.Forms.Button  
    $button_load_primary_user.Location = New-Object System.Drawing.Size(188,8)  
    $button_load_primary_user.Size = New-Object System.Drawing.Size(58,32)  
    $button_load_primary_user.TextAlign = "MiddleCenter"  
    $button_load_primary_user.Text = "Load"  
    $button_load_primary_user.backcolor = "white"
    $button_load_primary_user.Add_Click({ 
        
        $ListView.Items.Clear()
        $script:primary = Get-ADPrincipalGroupMembership $user | Get-ADGroup -Properties name,description | Sort-Object 
        $primary | Sort-Object | ForEach-Object {$item = New-Object System.Windows.Forms.ListViewItem($_.name)
        $ListView.Items.Add($item)
        $item.SubItems.Add($_.description)

        }
    $script:gridviewtitle = "$user is member of:"   
    })  
    $form_HelloWorld.Controls.add($button_load_primary_user)

#Row 2
#-------------------------------------------------------------------------

#Label equivalent 
$label_equivalent_user = New-Object System.Windows.Forms.Label  
    $label_equivalent_user.Location = New-Object System.Drawing.Size(8,48)  
    $label_equivalent_user.Size = New-Object System.Drawing.Size(116,32)  
    $label_equivalent_user.TextAlign = "MiddleCenter"  
    $label_equivalent_user.Text = "$equivalent"  
    $form_HelloWorld.Controls.Add($label_equivalent_user)

#Button to set equivalent  
$button_equivalent = New-Object System.Windows.Forms.Button  
    $button_equivalent.Location = New-Object System.Drawing.Size(124,48)  
    $button_equivalent.Size = New-Object System.Drawing.Size(58,32)  
    $button_equivalent.TextAlign = "MiddleCenter"  
    $button_equivalent.Text = "Set"
    $button_equivalent.backcolor = "white" 
    $button_equivalent.Add_Click({ $script:equivalent = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter username", "Equivalent User")  
    $script:secondary = Get-ADPrincipalGroupMembership $equivalent | Get-ADGroup -Properties name,description
        If($?)
        {
        $label_equivalent_user.Text = "$equivalent"
        $label_equivalent_user.ForeColor = "black"
        }
    else
        {
        $label_equivalent_user.Text = "$equivalent not found"
        $label_equivalent_user.ForeColor = "red"
        }  
    
    })  
    $form_HelloWorld.Controls.add($button_equivalent)

#Button to Load Equivalent
$button_loade = New-Object System.Windows.Forms.Button  
    $button_loade.Location = New-Object System.Drawing.Size(188,48)  
    $button_loade.Size = New-Object System.Drawing.Size(58,32)  
    $button_loade.TextAlign = "MiddleCenter"  
    $button_loade.Text = "Load"
    $button_loade.BackColor = "white"
    $button_loade.Add_Click({

        $ListView.Items.Clear()
        $script:secondary = Get-ADPrincipalGroupMembership $equivalent | Get-ADGroup -Properties name,description | Sort-Object
        $secondary | Sort-Object | ForEach-Object {$item = New-Object System.Windows.Forms.ListViewItem($_.name)
        $ListView.Items.Add($item)
        $item.SubItems.Add($_.description) 

        }
    $script:gridviewtitle = "$equivalent is member of:"    
    })  
    $form_HelloWorld.Controls.add($button_loade)

#Row 3
#-------------------------------------------------------------------------

#Label Primary Asset  
$label_asset = New-Object System.Windows.Forms.Label  
    $label_asset.Location = New-Object System.Drawing.Size(8,88)  
    $label_asset.Size = New-Object System.Drawing.Size(116,32)  
    $label_asset.TextAlign = "MiddleCenter"  
    $label_asset.Text = "$asset"
    $form_HelloWorld.Controls.Add($label_asset)  

#Button to set Primary Asset
$button_asset = New-Object System.Windows.Forms.Button  
    $button_asset.Location = New-Object System.Drawing.Size(124,88)  
    $button_asset.Size = New-Object System.Drawing.Size(58,32)  
    $button_asset.TextAlign = "MiddleCenter"  
    $button_asset.Text = "Set"
    $button_asset.BackColor = "white"
    $button_asset.Add_Click({ $script:asset = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter asset tag", "Primary Asset Tag")
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
    $button_loada.Location = New-Object System.Drawing.Size(188,88)  
    $button_loada.Size = New-Object System.Drawing.Size(58,32)  
    $button_loada.TextAlign = "MiddleCenter"  
    $button_loada.Text = "Load" 
    $button_loada.BackColor = "white"
    $button_loada.Add_Click({
    
        $ListView.Items.Clear()
        $script:asset1 = Get-ADPrincipalGroupMembership (Get-ADComputer $asset).DistinguishedName | Get-ADGroup -Properties name,description | Sort-Object 
        $asset1 | Sort-Object | ForEach-Object {$item = New-Object System.Windows.Forms.ListViewItem($_.name)
        $ListView.Items.Add($item)
        $item.SubItems.Add($_.description)

        }
    $script:gridviewtitle = "$asset is member of:"   
    })  
    $form_HelloWorld.Controls.add($button_loada)

#Row 4
#-------------------------------------------------------------------------

#Label equivalent Asset Tag 
$label_easset = New-Object System.Windows.Forms.Label  
    $label_easset.Location = New-Object System.Drawing.Size(8,128)  
    $label_easset.Size = New-Object System.Drawing.Size(116,32)  
    $label_easset.TextAlign = "MiddleCenter"  
    $label_easset.Text = "$easset"  
    $form_HelloWorld.Controls.Add($label_easset)

#Button to set equivalent asset tag 
$button_easset = New-Object System.Windows.Forms.Button  
    $button_easset.Location = New-Object System.Drawing.Size(124,128)  
    $button_easset.Size = New-Object System.Drawing.Size(58,32)  
    $button_easset.TextAlign = "MiddleCenter"  
    $button_easset.Text = "Set"
    $button_easset.BackColor = "white"  
    $button_easset.Add_Click({ $script:easset = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter asset tag", "Equivalent Asset Tag")
    $script:asset2 = Get-ADPrincipalGroupMembership (Get-ADComputer $easset).DistinguishedName | Get-ADGroup -Properties name,description
    If($?)
        {
        $label_easset.Text = "$easset"
        $label_easset.foreColor = "black"
        }
    else
        {
        $label_easset.Text = "$easset not found"
        $label_easset.foreColor = "red"

        }
    
    })  
    $form_HelloWorld.Controls.add($button_easset)

#button to load equivalent asset tag groups
$button_loadea = New-Object System.Windows.Forms.Button  
    $button_loadea.Location = New-Object System.Drawing.Size(188,128)  
    $button_loadea.Size = New-Object System.Drawing.Size(58,32)  
    $button_loadea.TextAlign = "MiddleCenter"  
    $button_loadea.Text = "Load" 
    $button_loadea.BackColor = "white"
    $button_loadea.Add_Click({ 
    
        $ListView.Items.Clear()
        $script:asset2 = Get-ADPrincipalGroupMembership (Get-ADComputer $easset).DistinguishedName | Get-ADGroup -Properties name,description | Sort-Object
        $asset2 | Sort-Object | ForEach-Object {$item = New-Object System.Windows.Forms.ListViewItem($_.name)
        $ListView.Items.Add($item)
        $item.SubItems.Add($_.description)

        }
    $script:gridviewtitle = "$easset is member of:"  
    })
    $form_HelloWorld.Controls.add($button_loadea)

#Row 5
#-------------------------------------------------------------------------

#Label Compare  
$label_compare = New-Object System.Windows.Forms.Label  
    $label_compare.Location = New-Object System.Drawing.Size(8,168)  
    $label_compare.Size = New-Object System.Drawing.Size(116,32)  
    $label_compare.TextAlign = "MiddleCenter"  
    $label_compare.Text = "Compare AD groups"  
    $form_HelloWorld.Controls.Add($label_compare)

#Button users Compare
$button_compare = New-Object System.Windows.Forms.Button  
    $button_compare.Location = New-Object System.Drawing.Size(124,168)  
    $button_compare.Size = New-Object System.Drawing.Size(58,32)  
    $button_compare.TextAlign = "MiddleCenter"  
    $button_compare.Text = "Users" 
    $button_compare.backcolor = "white" 
    $button_compare.Add_Click({

        $ListView.Items.Clear()
        
        $Secondary.name | ?{$Primary.name -notcontains $_} | Sort-Object |  ForEach-Object { $item = New-Object System.Windows.Forms.ListViewItem($_)
        $ListView.Items.Add($item)

        $hook = Get-ADGroup -Identity "$_" -Properties description
        $hook2 = $hook.Description
        $item.SubItems.Add("$hook2")

        }
    $script:gridviewtitle = "AD groups that $equivalent have, that $user don't"  
    })  
    $form_HelloWorld.Controls.add($button_compare)

#Button assets compare
$button_acompare = New-Object System.Windows.Forms.Button  
    $button_acompare.Location = New-Object System.Drawing.Size(188,168)  
    $button_acompare.Size = New-Object System.Drawing.Size(58,32)  
    $button_acompare.TextAlign = "MiddleCenter"  
    $button_acompare.Text = "Comp"
    $button_acompare.BackColor = "white"
    $button_acompare.Add_Click({ 

        $ListView.Items.Clear()
        
        $asset2.name | ?{$asset1.name -notcontains $_} | Sort-Object |  ForEach-Object { $item = New-Object System.Windows.Forms.ListViewItem($_)
        $ListView.Items.Add($item)

        $hook = Get-ADGroup -Identity "$_" -Properties description
        $hook2 = $hook.Description
        $item.SubItems.Add("$hook2") }
    $script:gridviewtitle = "AD groups that $easset have, that $asset don't"  
    })  
    $form_HelloWorld.Controls.add($button_acompare)

#Row 6
#-------------------------------------------------------------------------

#Button copy
$button_copy = New-Object System.Windows.Forms.Button  
    $button_copy.Location = New-Object System.Drawing.Size(8,209)  
    $button_copy.Size = New-Object System.Drawing.Size(116,32)  
    $button_copy.TextAlign = "MiddleCenter"  
    $button_copy.Text = "&Copy"
    $button_copy.BackColor = "white" 
    $button_copy.Add_Click({
    $Selected = @($ListView.SelectedItems)
    $Selected.text |
    ForEach-Object {
    If($counter -eq 1)
        {$clip = $clip + $_
        [int]$counter = $counter + 1
        }
    else {$clip = $clip + ", " + $_}
    }
    $clip | clip
    })  
    $form_HelloWorld.Controls.add($button_copy)


#Export button
$button_export = New-Object System.Windows.Forms.Button  
    $button_export.Location = New-Object System.Drawing.Size(132,209)  
    $button_export.Size = New-Object System.Drawing.Size(116,32)  
    $button_export.TextAlign = "MiddleCenter"  
    $button_export.Text = "&Export"
    $button_export.BackColor = "white" 
    $button_export.Add_Click({ 
    $script:path = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter location and name of the file. Example: C:\Users\username\Desktop\export.csv", "please enter location")
    $Selected = @($ListView.SelectedItems) 
    $Selected |
     ForEach-Object{
        $a = $_.text
        $b = $_.subitems[1].text
            [pscustomobject]@{
              Name=$a
              Description=$b
              
        }
    } |
    Export-Csv -Path "$path" -NoTypeInformation -Append
   
    })  
    $form_HelloWorld.Controls.add($button_export)

#Row 7
#-------------------------------------------------------------------------

#Out-GridView button
$button_gridview = New-Object System.Windows.Forms.Button  
    $button_gridview.Location = New-Object System.Drawing.Size(8,249)  
    $button_gridview.Size = New-Object System.Drawing.Size(240,32)  
    $button_gridview.TextAlign = "MiddleCenter"  
    $button_gridview.Text = "&Open list in separete window"
    $button_gridview.BackColor = "white" 
    $button_gridview.Add_Click({ 
    
    $Selected = @($ListView.Items) 
    $Selected |
     ForEach-Object{
        $a = $_.text
        $b = $_.subitems[1].text
            [pscustomobject]@{
              Name=$a
              Description=$b
              
        }
    } |
    Out-GridView -Title $gridviewtitle
   
    })  
    $form_HelloWorld.Controls.add($button_gridview)

#Row 8
#-------------------------------------------------------------------------

#Label Administrator Functions
$label_admin = New-Object System.Windows.Forms.Label  
    $label_admin.Location = New-Object System.Drawing.Size(8,289)  
    $label_admin.Size = New-Object System.Drawing.Size(240,32)  
    $label_admin.TextAlign = "MiddleCenter"  
    $label_admin.BorderStyle = "fixed3d"
    $label_admin.Text = "Administrator Functions"  
    $form_HelloWorld.Controls.Add($label_admin)

#Row 9
#-------------------------------------------------------------------------

#Add-groupmembership Primary User
$button_add_group_user = New-Object System.Windows.Forms.Button  
    $button_add_group_user.Location = New-Object System.Drawing.Size(8,329)  
    $button_add_group_user.Size = New-Object System.Drawing.Size(116,32)  
    $button_add_group_user.TextAlign = "MiddleCenter"  
    $button_add_group_user.Text = "&Add group to primary user"
    $button_add_group_user.BackColor = "white" 
    $button_add_group_user.Add_Click({ 
    
    $group_name = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter AD group name", "Add primary user to AD group")
    Get-ADGroup $group_name
    If($? -eq $false)
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $error_msg = $error[0].Exception.Message
    $wshell.Popup("Operation Failed: $error_msg")
    }
    else
    {
        Add-ADGroupMember $group_name -Members $user
        If($? -eq $false)
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $error_msg = $error[0].Exception.Message
        $wshell.Popup("Operation Failed: $error_msg")
        }
        else
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Operation Succeded: $user is now member of $group_name")
        }
        
    }
    })  
    $form_HelloWorld.Controls.add($button_add_group_user)

#Remove-groupmembership Primary User
$button_remove_group_user = New-Object System.Windows.Forms.Button  
    $button_remove_group_user.Location = New-Object System.Drawing.Size(132,329)  
    $button_remove_group_user.Size = New-Object System.Drawing.Size(116,32)  
    $button_remove_group_user.TextAlign = "MiddleCenter"  
    $button_remove_group_user.Text = "&Remove group from primary user"
    $button_remove_group_user.BackColor = "white" 
    $button_remove_group_user.Add_Click({ 
    
    $group_name = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter AD group name", "Remove primary user from AD group")
    Get-ADGroup $group_name
    If($? -eq $false)
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $error_msg = $error[0].Exception.Message
    $wshell.Popup("Operation Failed: $error_msg")
    }
    else
    {
        remove-ADGroupMember $group_name -Members $user -Confirm:$false
        If($? -eq $false)
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $error_msg = $error[0].Exception.Message
        $wshell.Popup("Operation Failed: $error_msg")
        }
        else
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Operation Succeded: $user is no longer member of $group_name")
        }
        
    }
    })  
    $form_HelloWorld.Controls.add($button_remove_group_user)

#Row 10
#-------------------------------------------------------------------------

#add-groupmembership Primary asset
$button_add_group_asset = New-Object System.Windows.Forms.Button  
    $button_add_group_asset.Location = New-Object System.Drawing.Size(8,369)  
    $button_add_group_asset.Size = New-Object System.Drawing.Size(116,32)  
    $button_add_group_asset.TextAlign = "MiddleCenter"  
    $button_add_group_asset.Text = "Add group to primary a&sset"
    $button_add_group_asset.BackColor = "white" 
    $button_add_group_asset.Add_Click({ 
    
    $group_name = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter AD group name", "Add primary asset tag to AD group")
    Get-ADGroup $group_name
    If($? -eq $false)
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $error_msg = $error[0].Exception.Message
    $wshell.Popup("Operation Failed: $error_msg")
    }
    else
    {
        $assetplusdolar = $asset + "$"
        add-ADGroupMember $group_name -Members $assetplusdolar -Confirm:$false
        If($? -eq $false)
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $error_msg = $error[0].Exception.Message
        $wshell.Popup("Operation Failed: $error_msg")
        }
        else
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Operation Succeded: $asset is now member of $group_name")
        }
        
    }
    
    })  
    $form_HelloWorld.Controls.add($button_add_group_asset)

#Remove-groupmembership Equivalent User
$button_remove_group_easset = New-Object System.Windows.Forms.Button  
    $button_remove_group_easset.Location = New-Object System.Drawing.Size(132,369)  
    $button_remove_group_easset.Size = New-Object System.Drawing.Size(116,32)  
    $button_remove_group_easset.TextAlign = "MiddleCenter"  
    $button_remove_group_easset.Text = "Remove group from primary asse&t"
    $button_remove_group_easset.BackColor = "white" 
    $button_remove_group_easset.Add_Click({ 

    $group_name = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter AD group name", "Remove primary asset tag from AD group")
    Get-ADGroup $group_name
    If($? -eq $false)
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $error_msg = $error[0].Exception.Message
    $wshell.Popup("Operation Failed: $error_msg")
    }
    else
    {
        $assetplusdolar = $asset + "$"
        remove-ADGroupMember $group_name -Members $assetplusdolar -Confirm:$false
        If($? -eq $false)
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $error_msg = $error[0].Exception.Message
        $wshell.Popup("Operation Failed: $error_msg")
        }
        else
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Operation Succeded: $asset is no longer member of $group_name")
        }
        
    }

    })  
    $form_HelloWorld.Controls.add($button_remove_group_easset)

#Row 11
#-------------------------------------------------------------------------

#Rsop primary user
$button_rsop_primary = New-Object System.Windows.Forms.Button  
    $button_rsop_primary.Location = New-Object System.Drawing.Size(8,409)  
    $button_rsop_primary.Size = New-Object System.Drawing.Size(116,32)  
    $button_rsop_primary.TextAlign = "MiddleCenter"  
    $button_rsop_primary.Text = "RSoP primary user"
    $button_rsop_primary.BackColor = "white" 
    $button_rsop_primary.Add_Click({ 
    
    $script:path = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter location and name of the file. Example: C:\Users\username\Desktop\RSoP.html", "please enter location")
    $button_rsop_primary.Text = "loading..."
    $button_rsop_primary.forecolor = "red"

    Get-GPResultantSetofPolicy -computer $asset -User $user -ReportType Html -Path $path

    Get-ChildItem -Path $path
    If($? -eq $false)
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Operation Failed")
    }
    else
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Operation Succeeded: RSoP for $user on $asset got saved on $path")
    }

    $button_rsop_primary.Text = "RSoP primary user"
    $button_rsop_primary.forecolor = "black"

    })  
    $form_HelloWorld.Controls.add($button_rsop_primary)

#Rsop equivalent user
$button_rsop_equivalent = New-Object System.Windows.Forms.Button  
    $button_rsop_equivalent.Location = New-Object System.Drawing.Size(132,409)  
    $button_rsop_equivalent.Size = New-Object System.Drawing.Size(116,32)  
    $button_rsop_equivalent.TextAlign = "MiddleCenter"  
    $button_rsop_equivalent.Text = "RSoP equivalent user"
    $button_rsop_equivalent.BackColor = "white" 
    $button_rsop_equivalent.Add_Click({ 

    $script:path = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter location and name of the file. Example: C:\Users\username\Desktop\RSoP.html", "please enter location")
    $button_rsop_equivalent.Text = "loading..."
    $button_rsop_equivalent.forecolor = "red"

    Get-GPResultantSetofPolicy -computer $easset -User $equivalent -ReportType Html -Path $path

    Get-ChildItem -Path $path
    If($? -eq $false)
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Operation Failed")
    }
    else
    {
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Operation Succeeded: RSoP for $equivalent on $easset got saved on $path")
    }

    $button_rsop_equivalent.Text = "RSoP equivalent user"
    $button_rsop_equivalent.forecolor = "black"

    })  
    $form_HelloWorld.Controls.add($button_rsop_equivalent)

#Row 12
#-------------------------------------------------------------------------

#Members of the group
$button_group_members = New-Object System.Windows.Forms.Button  
    $button_group_members.Location = New-Object System.Drawing.Size(8,449)  
    $button_group_members.Size = New-Object System.Drawing.Size(240,32)  
    $button_group_members.TextAlign = "MiddleCenter"  
    $button_group_members.Text = "Show me group members"
    $button_group_members.BackColor = "white" 
    $button_group_members.Add_Click({ 
    
        $group_name = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter AD group name", "List all group members")
        Get-ADGroup $group_name
        If($? -eq $false)
        {
        $wshell = New-Object -ComObject Wscript.Shell
        $error_msg = $error[0].Exception.Message
        $wshell.Popup("Operation Failed: $error_msg")
        }
        else
        {
        Get-ADGroupMember $group_name | select Name, ObjectClass | Out-GridView -Title "Members of $group_name"
        }

    })  
    $form_HelloWorld.Controls.add($button_group_members)

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
