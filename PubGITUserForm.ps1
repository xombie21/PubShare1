# Import active directory module for running AD cmdlets
Import-Module ActiveDirectory

# the return variable which is updated in the OK Click event
$Selection = 

# Add form assembly types
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# MainWindow Variables
$MainWindow = New-Object System.Windows.Forms.Form
$MainWindow.StartPosition = 'CenterScreen'
$MainWindow.Text = 'User Creation Form'
$MainWindow.Size = New-Object System.Drawing.Size(300, 150)
$MainWindow.KeyPreview = $True
$MainWindow.Select($MainWindowDropDownBox)

# OK Button variables
$MainWindowOKButton = New-Object System.Windows.Forms.Button
$MainWindowOKButton.Location = New-Object System.Drawing.Size(75, 60)
$MainWindowOKButton.Size = New-Object System.Drawing.Size(75, 23)
$MainWindowOKButton.Text = "OK"
$MainWindowOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$MainWindowOKButton.Add_Click({ 
    $Script:Selection = $MainWindowDropDownBox.SelectedItem
 })
 $MainWindowOKButton.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $MainWindowOKButton.PerformClick()
    }
 })

 $MainWindow.Controls.Add($MainWindowOKButton)




# Define text and functionality for "Cancel" button
$MainWindowCancelButton = New-Object System.Windows.Forms.Button
$MainWindowCancelButton.Location = New-Object System.Drawing.Size(75, 80)
$MainWindowCancelButton.Size = New-Object System.Drawing.Size(75, 23)
$MainWindowCancelButton.Text = "Cancel"
$MainWindowCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$MainWindow.Controls.Add($MainWindowCancelButton)
$MainWindow.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $MainWindow.Close()
    }
 })

# Define accept and cancel buttons
$MainWindow.AcceptButton = $MainWindowOKButton
$MainWindow.CancelButton = $MainWindowCancelButton

# Dropdownbox variables
$MainWindowDropDownBox = New-Object System.Windows.Forms.ComboBox
$MainWindowDropDownBox.DropDownStyle = 'DropDownList'
$MainWindowDropDownBox.Location = New-Object System.Drawing.Size(20, 20)
$MainWindowDropDownBox.Size = New-Object System.Drawing.Size(180, 20) 
$MainWindowDropDownBox.Height = 200

[void] $MainWindowDropDownBox.Items.Add("New User")
[void] $MainWindowDropDownBox.Items.Add("Existing User")
[void] $MainWindowDropDownBox.Items.Add("New Consultant User")
$MainWindowDropDownBox.SelectedIndex = 0
$MainWindowDropDownBox.Add_SelectedIndexChanged({
    # for demo, just write to console
    Write-Host "You have selected '$($this.SelectedItem)'" -ForegroundColor Cyan
    # inside here, you can refer to the $MainWindowDropDownBox as $this
    switch ($this.SelectedItem) {
        "New User" { 
            # Select the OU the new user account will be placed in
           $OU = Get-ADOrganizationalUnit -Filter * | Select-Object -Property Name, DistinguishedName | Out-GridView -PassThru | Select-Object -ExpandProperty DistinguishedName
           $OUChoice = $OU
       
           # Assembly for creating form & buttons
           [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
           [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
       
           # Define the form size & placement
           $NewUserForm = New-Object System.Windows.Forms.Form 
           $NewUserForm.Text = "New User Account Form"
           $NewUserForm.Size = New-Object System.Drawing.Size(400, 250) 
           $NewUserForm.StartPosition = "CenterScreen"
       
           $NewUserForm.KeyPreview = $True
           $NewUserForm.Add_KeyDown({ if ($_.KeyCode -eq "Enter" )
                   { $NewUserForm.Close(), [System.Windows.Forms.DialogResult]::OK } })
           $NewUserForm.Add_KeyDown({ if ($_.KeyCode -eq "Escape")
                   { $NewUserForm.Close(), [System.Windows.Forms.DialogResult]::Cancel } })
       
       
           # Define text and functionality for "OK" button
           $NewUserFormOKButton = New-Object System.Windows.Forms.Button
           $NewUserFormOKButton.Location = New-Object System.Drawing.Size(75, 160)
           $NewUserFormOKButton.Size = New-Object System.Drawing.Size(75, 23)
           $NewUserFormOKButton.Text = "OK"
           $NewUserFormOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
           $NewUserForm.AcceptButton = $NewUserFormOKButton
           $NewUserFormOKButton.Add_Click({ $NewUserForm.Close() })
           $NewUserForm.Controls.Add($NewUserFormOKButton)
       
           # Define text and functionality for "Cancel" button
           $NewUserFormCancelButton = New-Object System.Windows.Forms.Button
           $NewUserFormCancelButton.Location = New-Object System.Drawing.Size(150, 160)
           $NewUserFormCancelButton.Size = New-Object System.Drawing.Size(75, 23)
           $NewUserFormCancelButton.Text = "Cancel"
           $NewUserFormCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
           $NewUserForm.CancelButton = $NewUserFormCancelButton
           $NewUserForm.Controls.Add($NewUserFormCancelButton)
       
           # Window title
           $NewUserFormTitle = New-Object System.Windows.Forms.Label
           $NewUserFormTitle.Location = New-Object System.Drawing.Size(10, 20) 
           $NewUserFormTitle.Size = New-Object System.Drawing.Size(250, 20) 
           $NewUserFormTitle.Text = "Please enter the information in the space below:"
           $NewUserForm.Controls.Add($NewUserFormTitle) 
       
           # Label box for user first name input
           $NewUserFirstNameTextBoxLabel = New-Object System.Windows.Forms.Label
           $NewUserFirstNameTextBoxLabel.Location = New-Object System.Drawing.Size(10, 43) 
           $NewUserFirstNameTextBoxLabel.Size = New-Object System.Drawing.Size(66, 20) 
           $NewUserFirstNameTextBoxLabel.Text = "First Name:"
           $NewUserForm.Controls.Add($NewUserFirstNameTextBoxLabel) 
       
           # Label box for user last name input
           $NewUserLastNameTextBoxLabel = New-Object System.Windows.Forms.Label
           $NewUserLastNameTextBoxLabel.Location = New-Object System.Drawing.Size(10, 73) 
           $NewUserLastNameTextBoxLabel.Size = New-Object System.Drawing.Size(65, 20) 
           $NewUserLastNameTextBoxLabel.Text = "Last Name:"
           $NewUserForm.Controls.Add($NewUserLastNameTextBoxLabel) 
       
           # Label box for user email input
           $NewUserEmailTextBoxLabel = New-Object System.Windows.Forms.Label
           $NewUserEmailTextBoxLabel.Location = New-Object System.Drawing.Size(10, 103) 
           $NewUserEmailTextBoxLabel.Size = New-Object System.Drawing.Size(65, 20) 
           $NewUserEmailTextBoxLabel.Text = "Email:"
           $NewUserForm.Controls.Add($NewUserEmailTextBoxLabel) 
       
           # Label box for user phone input
           $NewUserPhoneTextBoxLabel = New-Object System.Windows.Forms.Label
           $NewUserPhoneTextBoxLabel.Location = New-Object System.Drawing.Size(10, 133) 
           $NewUserPhoneTextBoxLabel.Size = New-Object System.Drawing.Size(65, 20) 
           $NewUserPhoneTextBoxLabel.Text = "Phone:"
           $NewUserForm.Controls.Add($NewUserPhoneTextBoxLabel) 
       
           # Text box for user first name input
           $NewUserFirstNameTextBox = New-Object System.Windows.Forms.TextBox 
           $NewUserFirstNameTextBox.Location = New-Object System.Drawing.Size(80, 40) 
           $NewUserFirstNameTextBox.Size = New-Object System.Drawing.Size(250, 20) 
           $NewUserForm.Controls.Add($NewUserFirstNameTextBox) 
       
           # Text box for user last name input
           $NewUserLastNameTextBox = New-Object System.Windows.Forms.TextBox 
           $NewUserLastNameTextBox.Location = New-Object System.Drawing.Size(80, 70) 
           $NewUserLastNameTextBox.Size = New-Object System.Drawing.Size(250, 20) 
           $NewUserForm.Controls.Add($NewUserLastNameTextBox) 
       
           # Text box for user email input
           $NewUserEmailTextBox = New-Object System.Windows.Forms.TextBox 
           $NewUserEmailTextBox.Location = New-Object System.Drawing.Size(80, 100) 
           $NewUserEmailTextBox.Size = New-Object System.Drawing.Size(250, 20) 
           $NewUserForm.Controls.Add($NewUserEmailTextBox) 
       
           # Text box for user phone input
           $NewUserPhoneTextBox = New-Object System.Windows.Forms.TextBox 
           $NewUserPhoneTextBox.Location = New-Object System.Drawing.Size(80, 130) 
           $NewUserPhoneTextBox.Size = New-Object System.Drawing.Size(250, 20) 
           $NewUserForm.Controls.Add($NewUserPhoneTextBox) 
       
           # Show the form
           $NewUserForm.Topmost = $True
           $NewUserForm.Add_Shown({ $NewUserForm.Activate() })
           [void]$NewUserForm.ShowDialog()
       
           $FirstName = $NewUserFirstNameTextBox.Text
           $LastName = $NewUserLastNameTextBox.Text
           $Email = $NewUserEmailTextBox.Text
           $Telephone = $NewUserPhoneTextBox.Text
       
           # Company ID to add to end of username
           $Text = $OU
           $ModText = [regex]::Matches($Text, '(?<=- )[\d]*').value
       
           # Build the new user's username variable
           $Username = $FirstName.Substring(0, 1).ToLower() + "." + $LastName.ToLower() + $ModText
       
           # Users password
           $Password = Invoke-WebRequest -Uri https://www.dinopass.com/password/strong | Select-Object -ExpandProperty content
       
           # Define UPN
           $UPN = "FABRIKAM.com"
       
           # The variables are then used to create a new user account in Active Directory
           # Check to see if the user already exists in AD
           if (Get-ADUser -F { SamAccountName -eq $Username }) {
       
               # If user does exist, give a warning
               Write-Warning "A user account with username $Username already exists in Active Directory." -ForegroundColor Red
           }
           else {
               
               # User does not exist then proceed to create the new user account
               New-ADUser -SamAccountName $Username -UserPrincipalName "$Username@$UPN" -Name "$Firstname $Lastname" -GivenName $Firstname -Surname $lastname -Enabled $True -DisplayName "$Firstname $Lastname" -Path $OUChoice -OfficePhone $Telephone -EmailAddress $Email -AccountPassword (ConvertTo-secureString $Password -AsPlainText -Force) -PasswordNeverExpires $True 
               
               Add-ADGroupMember -Identity "$ModText Users" -Member $Username
               
               # If user is created, show message.
               Write-Host "The user account $username is created." -ForegroundColor Cyan
           }
           $DesktopPath = "$($env:USERPROFILE)\Desktop\NewUsers.txt"
           $DesktopPath | ForEach-Object { 
               If (Test-Path -Path $_) { Get-Item $_ } 
               Else { New-Item -Path $_ -Force } 
               Add-Content -Path $DesktopPath -Value $Username, $Password
           } 
        }
        "Existing User" { 
            $Filter = { (Enabled -eq $true) -and (SamAccountName -notlike "svc_*") -and (SamAccountName -notlike "*test*") -and (Name -notlike "*test*") }
        $ADUser = Get-ADUser -Filter $Filter | Select-Object -Property Name, SamAccountName | Sort-Object -Property Name | Out-GridView -Title "User" -PassThru | Select-Object -ExpandProperty SamAccountName
        $OU2 = Get-ADOrganizationalUnit -Filter * | Select-Object -Property Name, DistinguishedName | Sort-Object -Property Name | Out-GridView -PassThru | Select-Object -ExpandProperty DistinguishedName
        $Text2 = $OU2
        $ModText2 = [regex]::Matches($Text2, '(?<=- )[\d]*').value + " " + "Admins"
        Add-ADGroupMember -Identity $ModText2 -Members $ADUser
        $Member = Get-ADGroupMember -Identity $ModText2 | select-object -Property name,SamAccountName | Where-Object -Property SamAccountName -EQ $ADUser
                
        if ($Member){
        Write-Host "$ADUser has been added to $ModText2" -ForegroundColor Green
        }
        
        else {
        Write-Host "$ADUser was not added to $ModText2" -ForegroundColor Red
        }
        
        Read-Host -Prompt "Press any key to continue or CTRL+C to quit" 
         }
        "New Internal User" {     
            # Define the form size & placement
            $INTUserForm = New-Object System.Windows.Forms.Form 
            $INTUserForm.Text = "New Internal User Form"
            $INTUserForm.Size = New-Object System.Drawing.Size(400, 250) 
            $INTUserForm.StartPosition = "CenterScreen"
        
            $INTUserForm.KeyPreview = $True
            $INTUserForm.Add_KeyDown({ if ($_.KeyCode -eq "Enter" -or $_.KeyCode -eq "Escape")
                    { $INTUserForm.Close() } })
        
        
            # Define text and functionality for "OK" button
            $INTOKButton = New-Object System.Windows.Forms.Button
            $INTOKButton.Location = New-Object System.Drawing.Size(75, 160)
            $INTOKButton.Size = New-Object System.Drawing.Size(75, 23)
            $INTOKButton.Text = "OK"
            $INTOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $INTOKButton.Add_Click({ $INTUserForm.Close() })
            $INTUserForm.Controls.Add($INTOKButton)
        
            # Define text and functionality for "Cancel" button
            $INTCancelButton = New-Object System.Windows.Forms.Button
            $INTCancelButton.Location = New-Object System.Drawing.Size(150, 160)
            $INTCancelButton.Size = New-Object System.Drawing.Size(75, 23)
            $INTCancelButton.Text = "Cancel"
            $INTCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $INTUserForm.CancelButton = $INTCancelButton
            $INTUserForm.Controls.Add($INTCancelButton)
        
            # Window title
            $INTUserFormTitle = New-Object System.Windows.Forms.Label
            $INTUserFormTitle.Location = New-Object System.Drawing.Size(10, 20) 
            $INTUserFormTitle.Size = New-Object System.Drawing.Size(250, 20) 
            $INTUserFormTitle.Text = "Please enter the information in the space below:"
            $INTUserForm.Controls.Add($INTUserFormTitle) 
        
            # Label box for user first name input
            $INTUserFirstNameTextBoxLabel = New-Object System.Windows.Forms.Label
            $INTUserFirstNameTextBoxLabel.Location = New-Object System.Drawing.Size(10, 43) 
            $INTUserFirstNameTextBoxLabel.Size = New-Object System.Drawing.Size(66, 20) 
            $INTUserFirstNameTextBoxLabel.Text = "First Name:"
            $INTUserForm.Controls.Add($INTUserFirstNameTextBoxLabel) 
        
            # Label box for user last name input
            $INTUserLastNameTextBoxLabel = New-Object System.Windows.Forms.Label
            $INTUserLastNameTextBoxLabel.Location = New-Object System.Drawing.Size(10, 73) 
            $INTUserLastNameTextBoxLabel.Size = New-Object System.Drawing.Size(65, 20) 
            $INTUserLastNameTextBoxLabel.Text = "Last Name:"
            $INTUserForm.Controls.Add($INTUserLastNameTextBoxLabel) 
        
            # Text box for user first name input
            $INTUserFirstNameTextBox = New-Object System.Windows.Forms.TextBox 
            $INTUserFirstNameTextBox.Location = New-Object System.Drawing.Size(80, 40) 
            $INTUserFirstNameTextBox.Size = New-Object System.Drawing.Size(250, 20) 
            $INTUserForm.Controls.Add($INTUserFirstNameTextBox) 
        
            # Text box for user last name input
            $INTUserLastNameTextBox = New-Object System.Windows.Forms.TextBox 
            $INTUserLastNameTextBox.Location = New-Object System.Drawing.Size(80, 70) 
            $INTUserLastNameTextBox.Size = New-Object System.Drawing.Size(250, 20) 
            $INTUserForm.Controls.Add($INTUserLastNameTextBox) 
        
            # Show the form
            $INTUserForm.Topmost = $True
        
            $INTUserForm.Add_Shown({ $INTUserForm.Activate() })
            [void]$INTUserForm.ShowDialog()
        
            $FirstName = $INTUserFirstNameTextBox.Text
            $LastName = $INTUserLastNameTextBox.Text
            $OU = "OU=Users,DC=FABRIKAM,DC=com"
            # Build the new user's username variable
            $Username = "gg_" + $FirstName.Substring(0, 1).ToLower() + $LastName.ToLower()
        
        
            # Users password
            $Password = Invoke-WebRequest -Uri https://www.dinopass.com/password/strong | Select-Object -ExpandProperty content
        
            # Define UPN
            $UPN = "DataMasonsCloud.com"
        
            # The variables are then used to create a new user account in Active Directory
            # Check to see if the user already exists in AD
            if (Get-ADUser -F { SamAccountName -eq $Username }) {
                
                # If user does exist, give a warning
                Write-Warning "A user account with username $Username already exists in Active Directory."
            }
            else {
        
                # User does not exist then proceed to create the new user account
                New-ADUser -SamAccountName $Username -UserPrincipalName "$Username@$UPN" -Name "$Firstname $Lastname" -GivenName $Firstname -Surname $lastname -Enabled $True -DisplayName "$Firstname $Lastname" -Path $OU -OfficePhone $Telephone -EmailAddress $Email -AccountPassword (ConvertTo-secureString $Password -AsPlainText -Force) -PasswordNeverExpires $True 
        
                # If user is created, show message.
                Write-Host "The user account $username is created." -ForegroundColor Cyan
        
        
            }
            $DesktopPath = "$($env:USERPROFILE)\Desktop\NewUsers.txt"
            $DesktopPath | ForEach-Object { 
                If (Test-Path -Path $_) { Get-Item $_ } 
                Else { New-Item -Path $_ -Force } 
                Add-Content -Path $DesktopPath -Value $Username, $Password
            } }
        #"Item 4" { <# do something when Item 4 is selected #> }
        #"Item 5" { <# do something when Item 5 is selected #> }
    
    }
})


$MainWindow.Controls.Add($MainWindowDropDownBox) 
$MainWindow.Topmost = $True

$MainWindow.Add_Shown({$MainWindow.Activate()})
[void] $MainWindow.ShowDialog()

# IMPORTANT clean up the form when done
$MainWindow.Dispose()

$Selection
