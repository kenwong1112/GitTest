Import-Module ActiveDirectory

$DomainAdminUN = 'IPGNA\MDTSQL622.Service'
$DomainAdminPW = ConvertTo-SecureString 'L3x1ngton1' -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($DomainAdminUN, $DomainAdminPW)
$Server = "NYCGDC30.na.corp.ipgnetwork.com"
$OU = @("BGT","CAP","CTR","CWW","ENT","MCO","MCR","MEW","MGH","MHC","MHE","MRM","MUN","MWG","MWM","OPT","XBC")

Function Refresh-List
{
    $MWGworkstations = @()
    $SearchBase = "OU=Workstations,OU=MWG,OU=6223AV,OU=NYC,DC=na,DC=corp,DC=ipgnetwork,DC=com"
    $MWGworkstations += Get-ADComputer -SearchBase "OU=Workstations,OU=MWG,OU=6223AV,OU=NYC,DC=na,DC=corp,DC=ipgnetwork,DC=com" -Filter * -Properties name,lastLogon,description | Select name,@{Name="lastLogon";Expression={[datetime]::FromFileTime($_."lastLogon")}},description
    $objListBox.items.Clear()
    $MWGworkstations | Sort-Object -Property lastLogon | %{ 
        [string]$lastLogon = $_.lastLogon
        while(($lastLogon).Length -le 20){ $lastLogon += " " }
        $listItem = $_.name + " `t " + $lastLogon + " `t " + $_.description
        [void] $objListBox.Items.Add($listItem) 
    }
}

Function Unlock-Handler
{
    Foreach ($computer in $objListBox.SelectedItems)
    {
        Remove-ADComputer -Server $Server -Credential $Credentials -Identity $computer.Split().GetValue(0) -confirm:$false
    }

    Refresh-List
}

Function Make-Form {

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "MWG OU Workstations"
    $objForm.Size = New-Object System.Drawing.Size(800,400) 
    $objForm.StartPosition = "CenterScreen"
    $objForm.FormBorderStyle = 'Fixed3D'
    $objForm.MaximizeBox = $False
    $objForm.KeyPreview = $True
    

    # 'Enter' key
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Enter") { Unlock-Handler } } )
    
    # 'ESC' key
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Escape") { $objForm.Close() } } )

    # Delete button
    $UnlockButton = New-Object System.Windows.Forms.Button
    $UnlockButton.Location = New-Object System.Drawing.Size(580,325)
    $UnlockButton.Size = New-Object System.Drawing.Size(100,40)
    $UnlockButton.Text = "Delete"
    $UnlockButton.Add_Click( { Unlock-Handler } )
    $objForm.Controls.Add($UnlockButton)

    # Refresh button
    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = New-Object System.Drawing.Size(680,325)
    $RefreshButton.Size = New-Object System.Drawing.Size(100,40)
    $RefreshButton.Text = "Refresh"
    $RefreshButton.Add_Click({ Refresh-List })
    $objForm.Controls.Add($RefreshButton)

    # Label
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,10)
    $objLabel.Size = New-Object System.Drawing.Size(280,20)
    $objLabel.Text = "Select workstation(s) to delete:"
    $objForm.Controls.Add($objLabel) 

    # User List
    $objListBox = New-Object System.Windows.Forms.ListBox
    $objListBox.SelectionMode = "MultiExtended"
    $objListBox.Location = New-Object System.Drawing.Size(10,30)
    $objListBox.Size = New-Object System.Drawing.Size(770,290)
    Refresh-List
    $objForm.Controls.Add($objListBox)
    
    $objForm.Add_Shown({$objForm.Activate()})

    [void] $objForm.ShowDialog()
}

Make-Form