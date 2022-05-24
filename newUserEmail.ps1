# Trevor Watts 5/12/2022

Import-Module ActiveDirectory
$cred = Get-Credential
$self = Get-ADUser -Identity $cred.username -properties *



# Static variables and setting up script
$itHelpLink = '<a href = "mailto: ithelp@scprt.com">ithelp@scprt.com</a>'
$HrDirectLink = '<a href = "mailto: HRDirect@scprt.com">HRDirect@scprt.com</a>'
$parkOpLink = '<a href = "mailto: parkops@scprt.com">parkops@scprt.com</a>'
$getcsv = $null
$Users = $null
$getcsv = (Get-ChildItem "C:\Scripts\Complete New Users\*.csv").FullName
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

    #This asks the user if they are happy with email recipients
    $end = $false
    while($end -eq $false){
        echo " "
        echo $userObj.name
        for($i = 0; $i -le ($emailList.count - 1); $i++){
            $tempEmail = '' + ($i + 1)+ '. ' + $emailList[$i]
            echo $tempEmail
        }
        echo "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        echo "Are these emails correct?"
        echo "Type 1 to send email to these recipients"
        echo "Type 2 to add an email"
        echo "Type 3 to delete an email"
        echo "Type 4 to send email to yourself to manually modify email"
        echo "Type 5 to skip this user"

        $answer = Read-Host "Answer"
        Switch ($answer){
            
            1 {
                "Sending Email"
                $end = $true
                break
            }
        
            2 {
                $addEmail = read-host "New Email"
                $emailList.Add($addEmail)
                break
            }
            
            3 {
                for($i = 0; $i -le ($emailList.count - 1); $i++){
                    $string = '' + ($i + 1)+ '. ' + $emailList[$i]
                    echo $string
                }
                $removeEmail = read-host "Delete Email"
                $emailList.remove($emailList[$removeEmail - 1])
                break
            }
            
            4 {
                $sendToEmail = $self.mail
                $end = $true
                break
            }

            5 {
                'Skipped user'
                $skippedList.Add($userObj.name)
                $end = $true
                break
            }
            
            default {'Not a valid answer, Please try again'}
        }
    }
    if($answer -eq 5){continue}
    if($answer -ne 4){
        foreach($email in $emailList){
            $sendToEmail += $email + ";"
        }
    }

    $wcOrParkString = "" + $userObj.name + " - " + $userObj.office
    echo " "
    echo $wcOrParkString
    echo "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    echo "Type 1 for Park"
    echo "Type 2 for Welcome Center"
    $wcOrParkAnswer = Read-Host "Answer"
        Switch ($wcOrParkAnswer){
            
            1 {
                "Park"

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
                break
            }
           
            2 {
                "Welcome Center"

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


                Attachment 1:&nbsp;&nbsp;&nbsp;Welcome Center – New User – Email What to Expect
                <ul>
                <li>Will walk the new user through signing onto a PRT computer for the first time, how to change their temporary password, as well as how to access their email.</li>
                </ul> 
                Attachment 2:&nbsp;&nbsp;&nbsp;Welcome Center – Choosing a Printer for a New User
                <ul>
                <li>Will walk the new user through choosing a printer for their profile.</li>
                </ul>
                Attachment 3:&nbsp;&nbsp;&nbsp;How to View the Welcome Center’s Email Account
                <ul>
                <li>Will walk the new user through adding the Welcome Center’s email to their own email account.</li>
                <li>This attachment may or may not be needed depending upon whether or not you would like the new user to have access to the Welcome Center’s email account.</li>
                </ul>
                <br>
                If you would like the new user to have access to the Welcome Center email (to view and send emails), please send an email to the HelpDesk ($itHelpLink) and provide the following:
                <ul>
                <li>Subject Line:  Access to Welcome Center Email</li>
                <li>Body:  Name of Welcome Center</li>
                <li>Body:  Email Address of New User</li>
                <li>Body:  Include this statement: Please provide access to the Welcome Center’s email account.</li>
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
                break
            }
            
            default {'Not a valid answer, Please try again'}
        }

    #Sets password for new user
    Set-ADAccountPassword -Identity $userIdentity -Credential $cred -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $userPassword -Force)
    Set-ADuser -Identity $userIdentity -Credential $cred -ChangePasswordAtLogon $true
}
    
#Creates a new path and puts used csv file into 'Sent New Users' Directory
$path = "C:\Scripts\Complete New Users\Sent New Users"
If(!(test-path $path)){New-Item -ItemType Directory -Force -Path $path}
Move-Item -Path $getcsv -Destination $path

echo " "
echo " "
echo "SKIPPED: "
echo "~~~~~~~~~"
for($i = 0; $i -le ($skippedList.count - 1); $i++){
    $string = '' + ($i + 1)+ '. ' + $skippedList[$i]
    echo $string
}