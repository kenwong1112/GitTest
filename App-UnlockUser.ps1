Import-Module ActiveDirectory

$DomainAdminUN = 'IPGNA\MDTSQL622.Service'
$DomainAdminPW = ConvertTo-SecureString 'L3x1ngton!' -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($DomainAdminUN, $DomainAdminPW)
$Server = "NYCGDC30.na.corp.ipgnetwork.com"
$OU = @("BGT","CAP","CTR","CWW","ENT","MCO","MCR","MEW","MGH","MHC","MHE","MRM","MUN","MWG","MWM","OPT","XBC")
 
Function Refresh-List
{
    $LockedUser = @()
    Foreach ( $obj in $OU ) 
    {
        $SearchBase = "OU=$obj,OU=6223AV,OU=NYC,DC=na,DC=corp,DC=ipgnetwork,DC=com"
        #$LockedUser += Search-ADAccount -Server $Server -Credential $Credentials -SearchBase $SearchBase -LockedOut -UsersOnly | %{Get-ADUser $_.SamAccountName -Server $Server -Credential $Credentials -Properties SamAccountName, carLicense, badPwdCount, msDS-UserPasswordExpiryTimeComputed | Select SamAccountName, carLicense, badPwdCount, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}}
        $LockedUser += Get-ADUser -SearchBase $SearchBase -Server $Server -Credential $Credentials -Filter {carLicense -eq 'DisableEmailServices' -or lockouttime -ge 1 -and Enabled -eq $true} -Properties SamAccountName, carLicense, badPwdCount, msDS-UserPasswordExpiryTimeComputed | Select SamAccountName, carLicense, badPwdCount, @{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}
    }
    $objListBox.items.Clear()
    $LockedUser | Sort-Object -Property SamAccountName -Descending | %{ 
        $SamAccountName = $_.SamAccountName
        $carLicense = $_.carLicense
        while(($SamAccountName).Length -le 16){ $SamAccountName += " " }
        while(($carLicense).Length -le 1){ $carLicense += " `t " }
        $listItem = $SamAccountName + " `t " + $_.badPwdCount + " `t " + $carLicense + " `t " + $_.ExpiryDate
        [void] $objListBox.Items.Add($listItem) 
    }
}

Function Unlock-Handler
{
    Foreach ($Locked in $objListBox.SelectedItems)
    {
        Unlock-ADAccount -Server $Server -Credential $Credentials -Identity $Locked.Split().GetValue(0)
    }

    Refresh-List
}

Function Make-Form {

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Account Lockouts"
    $objForm.Size = New-Object System.Drawing.Size(460,300) 
    $objForm.StartPosition = "CenterScreen"
    $objForm.FormBorderStyle = 'Fixed3D'
    $objForm.MaximizeBox = $False
    $objForm.KeyPreview = $True
    

    # 'Enter' key
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Enter") { Unlock-Handler } } )
    
    # 'ESC' key
    $objForm.Add_KeyDown( { if ($_.KeyCode -eq "Escape") { $objForm.Close() } } )

    # Unlock button
    $UnlockButton = New-Object System.Windows.Forms.Button
    $UnlockButton.Location = New-Object System.Drawing.Size(230,220)
    $UnlockButton.Size = New-Object System.Drawing.Size(100,40)
    $UnlockButton.Text = "Unlock"
    $UnlockButton.Add_Click( { Unlock-Handler } )
    $objForm.Controls.Add($UnlockButton)

    # Refresh button
    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = New-Object System.Drawing.Size(340,220)
    $RefreshButton.Size = New-Object System.Drawing.Size(100,40)
    $RefreshButton.Text = "Refresh"
    $RefreshButton.Add_Click({ Refresh-List })
    $objForm.Controls.Add($RefreshButton)
    
    # Label-SamAccountName
    $objLabel1 = New-Object System.Windows.Forms.Label
    $objLabel1.Location = New-Object System.Drawing.Size(10,10)
    $objLabel1.Size = New-Object System.Drawing.Size(500,20)
    $objLabel1.Text = "Username      badPwdCount    Timesheet Status                  Password Expiration"
    $objForm.Controls.Add($objLabel1)

    # User List
    $objListBox = New-Object System.Windows.Forms.ListBox
    $objListBox.SelectionMode = "MultiExtended"
    $objListBox.Location = New-Object System.Drawing.Size(10,30)
    $objListBox.Size = New-Object System.Drawing.Size(430,180)
    Refresh-List
    $objForm.Controls.Add($objListBox)

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})

    [void] $objForm.ShowDialog()
}

Make-Form