# Trevor Watts 5/26/2022

Import-Module ActiveDirectory
Add-Type -assembly System.Windows.Forms
$cred = Get-Credential
$self = Get-ADUser -Identity $cred.username -properties *



# Static variables and setting up script
$itHelpLink = '<a href = "mailto: ithelp@scprt.com">ithelp@scprt.com</a>'
$HrDirectLink = '<a href = "mailto: HRDirect@scprt.com">HRDirect@scprt.com</a>'
$parkOpLink = '<a href = "mailto: parkops@scprt.com">parkops@scprt.com</a>'
$getcsv = $null
$Users = $null
$getcsv = (Get-ChildItem "C:\Scripts\Complete* New Users\*.csv").FullName
$Users = Import-Csv -Path $getcsv
$skippedList = $null
$skippedList = [System.Collections.Generic.List[string]]::new()



foreach($User in $Users)
{   
    #Resetting all variables so no left over information gets sent to the wrong person
    $Name = $null
    $lastName = $null
    $userObj = $null
    $userIdentity = $null
    $firstName = $null
    $userEmailLink = $null
    $userPassword = $null
    $userSupervisor = $null
    $userManager = $null
    $sendToEmail = $null
    $tempEmail = $null
    $emailList = $null
    $wcOrParkString = $null
    $emailList = [System.Collections.Generic.List[string]]::new()


    #Getting needed info from csv and active directory
    if($User."Network Logon?" -eq "No"){continue}
    $firstName = $User."First Name" -replace '[^a-zA-Z]', ''
    $lastName = $user."Last Name" -replace ' jr$| JR$| sr$| SR$| ii$| ii $| II$| II $| iii$| iii $| III$| III $| iv$| iv $| IV$| IV $|,jr$|, jr$|,JR$|, JR$|,sr$|, sr$|,SR$|, SR$|, jr $|,jr $|, JR $|,JR $|,sr $|,SR $|, sr $|, SR $|,ii$|, ii$|,II$|, II$|,ii $|,II $|, ii $|, II $|,iii$|,III$|, iii$|, III$|,iii $|,III $|, iii $|, III $|,iv$|,IV$|, iv$|, IV$|, iv $|, IV $'
    $lastName = $lastName -replace '[^a-zA-Z]', ''
    $Name = $firstName + "* *" + $lastName
    $userObj = get-aduser -Filter {Name -like $Name} -properties *
    $Name = $userObj.name
    $firstName = $userObj.GivenName
    $userIdentity = $userObj.SamAccountName
    $userEmailLink = '<a href = "mailto: '+$userObj.mail+'">'+$userObj.mail+'</a>'
    $userPassword = "SCPRT!0" + (Get-Date).Millisecond
    $userSupervisor = Get-ADUser -Identity $userObj.Manager -properties *
    $emailList.Add($userSupervisor.mail)

    #Supervisor is manager
    if($userSupervisor.description -like "*Manager*" -and $userSupervisor.description -notlike "*ass*"){
        Foreach($employee in $userSupervisor.directReports){
            $thisEmployee = Get-ADUser -Identity $employee -properties *
            #Manager has assistant manager
            if($thisEmployee.description -like "*Ass* *Manager*"){
                $emailList.Add($thisEmployee.mail)
            }
        }
    #Supervisor is Assistant Manager
    }elseif($userSupervisor.description -like "*Ass* *Manager*"){
        $userManager = Get-ADUser -Identity $userSupervisor.Manager -properties *
        $emailList.Add($userManager.mail)
    #Supervisor is neither Manager nor Assistant Manager
    }else {
        $userManager = Get-ADUser -Identity $userSupervisor.Manager -properties *
        $emailList.Add($userManager.mail)
        Foreach($employee in $userManager.directReports){
            #Manager has assistant manager
            $thisEmployee = Get-ADUser -Identity $employee -properties *
            if($thisEmployee.description -like "*Ass* *Manager*"){
                $emailList.Add($thisEmployee.mail)
            }
        }
    }

    #Main Form
    $main_form = New-Object System.Windows.Forms.Form
    $main_form.Text ='New User Email Script'
    $main_form.Width = 500
    $main_form.Height = 400
    $main_form.AutoSize = $true
    $main_form.Icon = "C:\Scripts\user.ico"

    #User's Name label
    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = $Name
    $lblUser.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
    $lblUser.Location  = New-Object System.Drawing.Point(10,10)
    $lblUser.AutoSize = $true
    $main_form.Controls.Add($lblUser)

    #Squiggle Line Label
    $lblLine = New-Object System.Windows.Forms.Label
    $lblLine.Text = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    $lblLine.Location  = New-Object System.Drawing.Point(10,35)
    $lblLine.AutoSize = $true
    $main_form.Controls.Add($lblLine)

    #ListBox For Emails
    $ListBox = New-Object System.Windows.Forms.ListBox
    $ListBox.Width = 290
    $ListBox.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
    $Users = get-aduser -filter * -Properties SamAccountName
    foreach($email in $emailList){
        if($email -ne ""){
            $ListBox.Items.Add($email)    
        }
        else{
            $listBox.Items.Add("email not found")
        }
    }
    $ListBox.Location  = New-Object System.Drawing.Point(10,60)
    $main_form.Controls.Add($ListBox)


    #Button to delete email from combobox
    $BtnDelete = New-Object System.Windows.Forms.Button
    $BtnDelete.Location = New-Object System.Drawing.Size(350,60)
    $BtnDelete.Size = New-Object System.Drawing.Size(120,23)
    $BtnDelete.Text = "Delete"
    $main_form.Controls.Add($BtnDelete)
    $BtnDelete.Add_Click(

    {
        $answer = [System.Windows.MessageBox]::Show( "Do you want to remove this user?", " Removal Confirmation", "YesNoCancel", "Warning" )
        if($answer -eq "Yes"){
            $ListBox.Items.Remove($ListBox.SelectedItem)
            $ListBox.Text = ""
        }

    })

    #Textbox to add new email
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Width = 290
    $textBox.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
    $textBox.Location  = New-Object System.Drawing.Point(10,160)
    $main_form.Controls.Add($textBox)

    #button to confirm added email in textbox
    $BtnAdd = New-Object System.Windows.Forms.Button
    $BtnAdd.Location = New-Object System.Drawing.Size(350,160)
    $BtnAdd.Size = New-Object System.Drawing.Size(120,23)
    $BtnAdd.Text = "Add"
    $main_form.Controls.Add($BtnAdd)
    $BtnAdd.Add_Click(

    {
        if($textbox.Text -eq ""){
            [void] [System.Windows.MessageBox]::Show( "Make sure to put text in the textbox before clicking add", "No text in textbox", "OK", "Warning" )
        }
        else {#if($answer -eq $dialogresult.Yes){
            $ListBox.Items.Add($textbox.Text)
            $ListBox.Text = $textbox.Text
            $textbox.Text = ""
        }

    })

    #GroupBox for radio buttons
    $groupBox = New-Object System.Windows.Forms.GroupBox 
    $groupBox.Location = New-Object System.Drawing.Size(10,180) 
    $groupBox.size = New-Object System.Drawing.Size(290,50) 
    $main_form.Controls.Add($groupBox) 
    
    #Radio button for Parks
    $rbParks = New-Object System.Windows.Forms.RadioButton 
    $rbParks.Location = new-object System.Drawing.Point(35,15) 
    $rbParks.size = New-Object System.Drawing.Size(70,20) 
    $rbParks.Checked = $true 
    $rbParks.Text = "Parks" 
    $groupBox.Controls.Add($rbParks) 

    #Radio button for Welcome centers
    $rbWelcomCenter = New-Object System.Windows.Forms.RadioButton 
    $rbWelcomCenter.Location = new-object System.Drawing.Point(135,15) 
    $rbWelcomCenter.size = New-Object System.Drawing.Size(150,20) 
    $rbWelcomCenter.Checked = $false 
    $rbWelcomCenter.Text = "Welcome Center" 
    $groupBox.Controls.Add($rbWelcomCenter) 

    #Button to send email
    $BtnSend = New-Object System.Windows.Forms.Button
    $BtnSend.Location = New-Object System.Drawing.Size(10,250)
    $BtnSend.Size = New-Object System.Drawing.Size(120,23)
    $BtnSend.Text = "Send Email"
    $main_form.Controls.Add($BtnSend)
    $BtnSend.Add_Click(
    {
        $answer = [System.Windows.MessageBox]::Show( "Are you sure you would like to send?", "Send Email", "YesNoCancel", "Warning" )
        if($answer -eq "Yes"){
            if ($rbParks.Checked -eq $true) {
            try{
                #Attaching the three pdf files to the email
                $GetFiles = Get-ChildItem "C:\Scripts\Attachments\Parks\*.pdf" 
                $Outlook = New-Object -ComObject Outlook.Application
                $Mail = $Outlook.CreateItem(0)
                Foreach($GetFile in $GetFiles)
                {
                    Write-Host "Attaching Files"
                    $mail.Attachments.Add($GetFile.FullName)
                }

                $Mail.To = $sendToEmail
                $Mail.Subject = $Name + ' Windows Login and Email Information' 
                $Mail.HTMLBody = 
                "$FirstName's email address along with temporary password has been setup and is ready to use. There are 3 attachments with this email, they are explained below. They are meant as help for the new user. If you could ensure they receive them that would be great! The new user should always sign into Windows (the screen with the palm tree) with their SCPRT email. 
                <br><br>
                Windows (email) information is as follows:
                <br><br>
                
                <p style='font-size: 20px;'><b>"+$Name+"</b><br>
                <b>logon as: "+$userEmailLink+"</b><br>
                <b>Temp Password: "+$userPassword+"</b></p><br>


                Attachment 1:&nbsp;&nbsp;&nbsp;Parks - What to Expect for First Time Sign On and Email
                <ul>
                <li>Will walk the new user through signing onto a PRT computer for the first time, how to change their temporary password, as well as how to access their email.</li>
                </ul> 
                Attachment 2:&nbsp;&nbsp;&nbsp;Parks - New User - Setting up a Receipt Printer
                <ul>
                <li>Will walk the new user through choosing the Citizen's receipt printer for their profile.</li>
                </ul>
                Attachment 3:&nbsp;&nbsp;&nbsp;How to View the Parks' Email Account
                <ul>
                <li>Will walk the new user through adding the Park's email to their own email account.</li>
                <li>This attachment may or may not be needed depending upon whether or not you would like the new user to have access to the Park's email account.</li>
                </ul>
                <br>
                If you would like the new user to have access to the Park's email (to view and send emails), please send an email to the HelpDesk ($itHelpLink) and provide the following:
                <ul>
                <li>Subject Line:  Access to Park's Email</li>
                <li>Body:  Name of Park</li>
                <li>Body:  Email Address of New User</li>
                <li>Body:  Include this statement:  Please provide access to the Park's email account.</li>
                </ul>
                <p>
                Your new user should be ready to take the Securing the Human Training and go through SCEIS for the policy review.  If you find your new user has not yet received the email from HR about the SANS training (securing the human) or they are having an issue with the SCEIS policy review, please email $HrDirectLink for assistance.
                </p>

                <p>
                If, on the PTF, you asked for the new user to have access to Itinio, POS, Revenue, or Enterprise, Park Operations ($parkOpLink) will send that information to you once the new user has completed Securing the Human Training and the SCEIS policy review.
                </p>"
                $Mail.Send()
                $Outlook.Quit() 
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
                [void] [System.Windows.MessageBox]::Show( "Email Sent Successful", "Success", "OK", "Information" )
            }
            catch{
                [void] [System.Windows.MessageBox]::Show( "Unable to send email properly", "Email Error", "OK", "Warning" )
            }}
            else{
            try{
                #Attaching the three pdf files to the email
                $GetFiles = Get-ChildItem "C:\Scripts\Attachments\WelcomeCenters\*.pdf" 
                $Outlook = New-Object -ComObject Outlook.Application
                $Mail = $Outlook.CreateItem(0)
                Foreach($GetFile in $GetFiles)
                {
                    Write-Host "Attaching Files"
                    $mail.Attachments.Add($GetFile.FullName)
                }

                $Mail.To = $sendToEmail
                $Mail.Subject = $Name + ' Windows Login and Email Information' 
                $Mail.HTMLBody = 
                "$FirstName's email address along with temporary password has been setup and is ready to use. There are 3 attachments with this email, they are explained below. They are meant as help for the new user. If you could ensure they receive them that would be great! The new user should always sign into Windows (the screen with the palm tree) with their SCPRT email. 
                <br><br>
                Windows (email) information is as follows:
                <br><br>
                
                <p style='font-size: 20px;'><b>"+$Name+"</b><br>
                <b>logon as: "+$userEmailLink+"</b><br>
                <b>Temp Password: "+$userPassword+"</b></p><br>


                Attachment 1:&nbsp;&nbsp;&nbsp;Welcome Center   New User   Email What to Expect
                <ul>
                <li>Will walk the new user through signing onto a PRT computer for the first time, how to change their temporary password, as well as how to access their email.</li>
                </ul> 
                Attachment 2:&nbsp;&nbsp;&nbsp;Welcome Center   Choosing a Printer for a New User
                <ul>
                <li>Will walk the new user through choosing a printer for their profile.</li>
                </ul>
                Attachment 3:&nbsp;&nbsp;&nbsp;How to View the Welcome Center s Email Account
                <ul>
                <li>Will walk the new user through adding the Welcome Center s email to their own email account.</li>
                <li>This attachment may or may not be needed depending upon whether or not you would like the new user to have access to the Welcome Center s email account.</li>
                </ul>
                <br>
                If you would like the new user to have access to the Welcome Center email (to view and send emails), please send an email to the HelpDesk ($itHelpLink) and provide the following:
                <ul>
                <li>Subject Line:  Access to Welcome Center Email</li>
                <li>Body:  Name of Welcome Center</li>
                <li>Body:  Email Address of New User</li>
                <li>Body:  Include this statement: Please provide access to the Welcome Center s email account.</li>
                </ul>
                <p>
                Your new user should be ready to take the Securing the Human Training and go through SCEIS for the policy review.  If you find your new user has not yet received the email from HR about the SANS training (securing the human) or they are having an issue with the SCEIS policy review, please email $HrDirectLink for assistance.
                </p>

                <p>
                If, on the PTF, you asked for the new user to have access to Itinio, POS, Revenue, or Enterprise, Park Operations ($parkOpLink) will send that information to you once the new user has completed Securing the Human Training and the SCEIS policy review.
                </p>"
                $Mail.Send()
                $Outlook.Quit() 
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
                [void] [System.Windows.MessageBox]::Show( "Email Sent Successful", "Success", "OK", "Information" )
            }catch{
                [void] [System.Windows.MessageBox]::Show( "Unable to send email properly", "Email Error", "OK", "Warning" )
            }}
        }

    })

    #Button to skip user
    $BtnSkip = New-Object System.Windows.Forms.Button
    $BtnSkip.Location = New-Object System.Drawing.Size(180,250)
    $BtnSkip.Size = New-Object System.Drawing.Size(120,23)
    $BtnSkip.Text = "Skip user"
    $main_form.Controls.Add($BtnSkip)
    $BtnSkip.Add_Click(
    {
        $answer = [System.Windows.MessageBox]::Show( "Are you sure you would like to skip this user?", "Skip User", "YesNoCancel", "Warning" )
        if($answer -eq "Yes"){
            $skippedList.Add($userObj.name)
            $main_form.Close()
        }

    })

    $main_form.ShowDialog()

    try{
    #Sets password for new user
    Set-ADAccountPassword -Identity $userIdentity -Credential $cred -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $userPassword -Force)
    Set-ADuser -Identity $userIdentity -Credential $cred -ChangePasswordAtLogon $true
    }
    catch{
        [void] [System.Windows.MessageBox]::Show( "Unable to reset password", "Password Reset Error", "OK", "Warning" )
    }
    $main_form.Close()
}
    
try{
    #Creates a new path and puts used csv file into 'Sent New Users' Directory
    $path = "C:\Scripts\Complete* New Users\Sent New Users"
    If(!(test-path $path)){New-Item -ItemType Directory -Force -Path $path}
    Move-Item -Path $getcsv -Destination $path
}catch{
    [void] [System.Windows.MessageBox]::Show( "Unable to move csv file. Make sure to close it next time. Please Manually move it into the 'Sent New Users' Folder", "Moving CSV Error", "OK", "Warning" )
}

#SKIPPED USERS FORM

#Skipped users form
$skipped_form = New-Object System.Windows.Forms.Form
$skipped_form.Text ='Skipped Users'
$skipped_form.Width = 300
$skipped_form.Height = 300
$skipped_form.AutoSize = $true
$skipped_form.Icon = "C:\Scripts\user.ico"

#User's Name label
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Skipped Users"
$lblUser.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$lblUser.Location  = New-Object System.Drawing.Point(10,10)
$lblUser.AutoSize = $true
$skipped_form.Controls.Add($lblUser)

#User's Name label
$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "Please Make sure to send the email to these users"
$lblUser.Font = New-Object System.Drawing.Font("Tahoma",8,[System.Drawing.FontStyle]::Regular)
$lblUser.Location  = New-Object System.Drawing.Point(10,35)
$lblUser.AutoSize = $true
$skipped_form.Controls.Add($lblUser)

#Squiggle Line Label
$lblLine = New-Object System.Windows.Forms.Label
$lblLine.Text = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
$lblLine.Location  = New-Object System.Drawing.Point(10,55)
$lblLine.AutoSize = $true
$skipped_form.Controls.Add($lblLine)

#ListBox For Emails
$ListBox = New-Object System.Windows.Forms.ListBox
$ListBox.Width = 290
$listBox.Height = 400
$ListBox.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Users = get-aduser -filter * -Properties SamAccountName
foreach($person in $skippedList){
    $ListBox.Items.Add($person)    
}
$ListBox.Location  = New-Object System.Drawing.Point(10,80)
$skipped_form.Controls.Add($ListBox)

$skipped_form.ShowDialog()