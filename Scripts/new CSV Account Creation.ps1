Import-Module ActiveDirectory
$cred = Get-Credential

#These are the static variables that are independent of EOD report data.
$OU = "OU=New User,OU=Middle Earth,DC=scprt,DC=com"
$OUforFTE = "OU=O365 E3 OneDrive Users,OU=Users2017,OU=Rollout 2017,OU=Middle Earth,DC=scprt,DC=com"
$OUforA2Parks = "OU=Parks,OU=O365 EOP2 Users,OU=Users2017,OU=Rollout 2017,OU=Middle Earth,DC=scprt,DC=com"
$OUforA2WelcomeCenters = "OU=Welcome Centers,OU=O365 EOP2 Users,OU=Users2017,OU=Rollout 2017,OU=Middle Earth,DC=scprt,DC=com"
$OUforA2CentralOffice = "OU=Central Office,OU=O365 EOP2 Users,OU=Users2017,OU=Rollout 2017,OU=Middle Earth,DC=scprt,DC=com"
$Company = "SCPRT"
$getcsv = (Get-ChildItem "C:\Scripts\*.csv").FullName
$Users = Import-Csv -Path $getcsv
foreach ($User in $Users)
{      
    #This is pulling all of the data from the EOD report and putting it in variables for the script.
    $UserFirstnamestart = $User."First Name"
    $UserLastnamestart = $User."Last Name"
    $UserMiddleInitial = $User."Middle Initial"
    $TrueUserMiddleInitial = $UserMiddleInitial[0]
    $ManagerName = $User.Supervisor
    $Password = "SCPRT!0" + (Get-Date).Millisecond #This variable must be run for every user individually so it is unique.
    $Department = $User."Document DistributionCode (P Code)"
    $Title = $User."Position Title"
    $Office = $User.Department
    $NetworkLogon = $User."Network Logon?"
    $PositionCategory = $User."Position Category"
    $Division = $User.Division
    $TransactionReason = $User."Transaction Reason"

    #This is to remove jr, sr, II, III, IV, etc from the last name variable

    $UserLastnamestart = $UserLastnamestart -replace ' jr$| JR$| sr$| SR$| ii$| ii $| II$| II $| iii$| iii $| III$| III $| iv$| iv $| IV$| IV $|,jr$|, jr$|,JR$|, JR$|,sr$|, sr$|,SR$|, SR$|, jr $|,jr $|, JR $|,JR $|,sr $|,SR $|, sr $|, SR $|,ii$|, ii$|,II$|, II$|,ii $|,II $|, ii $|, II $|,iii$|,III$|, iii$|, III$|,iii $|,III $|, iii $|, III $|,iv$|,IV$|, iv$|, IV$|, iv $|, IV $'

    #This is to remove non english alphabet characters from name variables

    $UserFirstnamenospecialcharacters = $UserFirstNamestart -replace '[^a-zA-Z]', ''
    $UserLastnamenospecialcharacters = $UserLastnamestart -replace '[^a-zA-Z]', ''
    
    #This is to remove possible spaces from the first and last names
    
    $UserFirstNamenospace = $UserFirstnamenospecialcharacters.replace(' ', '')
    $UserLastnamenospace = $UserLastnamenospecialcharacters.replace(' ', '')
    
    #These variables are to change the strings of the name and title variables from possibly all uppercase to Title Case
    
    $TextInfo = (Get-Culture).TextInfo
    $Title = $TextInfo.ToTitleCase($Title.ToLower())
    $UserFirstname = $TextInfo.ToTitleCase($UserFirstNamenospace.ToLower())
    $UserLastname = $TextInfo.ToTitleCase($UserLastnamenospace.ToLower())

    #This is to change update Position Title to be accurate for Campground Hosts so they don't show as title "Non-Employee" in GAL and AD
    If ($Title -eq "Non-Employee" -And $PositionCategory -eq "Volunteer/Campground Host"){
        $Title = "Campground Host"
    }

    #These variables need to come after the text manipulation to Title Case since it is using those variables in the variables listed below
    $Displayname = $UserFirstName + " " + $UserLastname

    #These variables are to change the strings of the name variables to all lowercase for future variable use for creation of the SAM variable
    $UserFirstnamelowercase = $UserFirstnamenospecialcharacters.ToLower()
    $UserLastnamelowercase = $UserLastnamenospecialcharacters.ToLower()

    #These variables are using the lowercase variables from the previous step. This is just for uniformity in AD, not functionality.
    $SAM = $UserFirstnamelowercase[0] + $UserLastnamelowercase
    $UPN = $SAM + "@scprt.com"
    $Proxyaddress = "SMTP:" + $UPN
    
    #This is to check for a specific Transaction Reason. This transaction reason may mean the user already has an account so we must check AD manually if prompted.
    If ($TransactionReason -eq "New Hire- Hourly to FTE" -And $NetworkLogon -eq "Yes") {
        $NetworkLogon = Read-Host "$SAM at $Office may already have an account. Please check AD for an account. If they need an account type 'Yes' if they do not need an account type 'No'"
    }

    #This is to see if the UPN is already in use. If it is in use it will change the SAM and UPN. If after looking at Active Directory and it is determined that the user already has an account it will change the SAM to a value that is then used to change the NetworkLogon variable value all so that an account does not get created for that user. It's possible this part especially could be scripted better.
    If ($NetworkLogon -eq "Yes"){
        $aduserexists = $null
        $aduserexists = Get-ADUser $SAM
        If ($aduserexists -ne $null)
            {$SAM = Read-Host "Username $SAM is taken. They are a $Title at $Office. Check ActiveDirectory and try a different username if they don't already have an account. If they already have an account type 'No'"
                $UPN = $SAM + "@scprt.com"
                $Proxyaddress = "SMTP:" + $UPN
            }
        Else {Write-Host "The username $SAM was available." -ForegroundColor DarkGreen -BackgroundColor White}
        }
    If ($SAM -eq "No"){
        $NetworkLogon = "No"
    }
#This switch paramater checks if the network logon variable contains yes. If so, it creates the account, if not then it does nothing and moves to the next user in EOD report.
    Switch ($NetworkLogon) {
        'Yes' {
            #These variables need to be below the switch statement for NetworkLogon so they aren't done unnecessarily. If a user doesn't need a network account we don't want to create these variables since they involve human interaction.
            $ManagerUsername = Read-Host "Carefully enter the username NO EMAIL for $ManagerName at $Office. If no manager name is shown here then enter the Park Manager username. For example jdoe"

            New-ADUser -Credential $cred -DisplayName "$DisplayName" -SamAccountName "$SAM" -Company "$Company" -UserPrincipalName "$UPN" -Department "$Department" -Manager "$ManagerUsername" -GivenName "$UserFirstname" -Surname "$UserLastname" -EmailAddress "$UPN" -Office "$Office" -Description "$Title" -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $true -Path "$OU" -ChangePasswordAtLogon $true -PasswordNeverExpires $false -Title "$Title" -Initials "$TrueUserMiddleInitial" -ScriptPath "logon.bat" -Name $DisplayName
            Get-ADUser -Credential $cred $SAM -Properties * | Set-ADUser -Add @{ProxyAddresses=$Proxyaddress}
            #This switch paramater uses the Office variable to add the user to the appropriate site group and add the site phone number if appropriate.
            Switch ($Office) {
                'Aiken State Park' {Add-ADGroupMember -Credential $cred -Identity "Aiken Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-649-2857"}
                'Andrew Jackson State Park' {Add-ADGroupMember -Credential $cred -Identity "Andrew Jackson Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-285-3344"}
                'Barnwell State Park' {Add-ADGroupMember -Credential $cred -Identity "Barnwell Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-284-2212"}
                'Calhoun Falls State Park' {Add-ADGroupMember -Credential $cred -Identity "Calhoun Falls Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-447-8267"}
                'Charles Towne Landing State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Charles Towne Landing Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-852-4200"}
                'Cheraw State Park' {Add-ADGroupMember -Credential $cred -Identity "Cheraw Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-537-9656"}
                'Chester State Park' {Add-ADGroupMember -Credential $cred -Identity "Chester Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-385-2680"}
                'Colleton State Park' {Add-ADGroupMember -Credential $cred -Identity "Colleton Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-538-8206"}
                'Colonial Dorchester State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Colonial Dorchester Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-873-1740"}
                'Croft State Park' {Add-ADGroupMember -Credential $cred -Identity "Croft Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-585-1283"}
                'Devils Fork State Park' {Add-ADGroupMember -Credential $cred -Identity "Devils Fork Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-944-2639"}
                'Dreher Island State Park' {Add-ADGroupMember -Credential $cred -Identity "Dreher Island Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-364-0756"}
                'Edisto Beach State Park' {Add-ADGroupMember -Credential $cred -Identity "Edisto Beach Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-869-2756"}
                'Givhans Ferry State Park' {Add-ADGroupMember -Credential $cred -Identity "Givhans Ferry Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-873-0692"}
                'H. Cooper Black Field Trial Area' {Add-ADGroupMember -Credential $cred -Identity "H. Cooper Black Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-378-1555"}
                'Hamilton Branch State Park' {Add-ADGroupMember -Credential $cred -Identity "Hamilton Branch Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-333-2223"}
                'Hampton Plantation State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Hampton Plantation Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-546-9361"}
                'Hickory Knob State Resort Park' {Add-ADGroupMember -Credential $cred -Identity "Hickory Knob Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-391-2450"}
                'Hunting Island State Park' {Add-ADGroupMember -Credential $cred -Identity "Hunting Island Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-838-2011"}
                'Huntington Beach State Park' {Add-ADGroupMember -Credential $cred -Identity "Huntington Beach Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-237-4440"}
                'Keowee-Toxaway State Park' {Add-ADGroupMember -Credential $cred -Identity "Keowee Toxaway Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-868-2605"}
                'Kings Mountain State Park' {Add-ADGroupMember -Credential $cred -Identity "Kings Mountain Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-222-3209"}
                'Lake Greenwood State Park' {Add-ADGroupMember -Credential $cred -Identity "Lake Greenwood Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-543-3535"}
                'Lake Hartwell State Park' {Add-ADGroupMember -Credential $cred -Identity "Lake Hartwell Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-972-3352"}
                'Lake Warren State Park' {Add-ADGroupMember -Credential $cred -Identity "Lake Warren Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-943-4736"}
                'Lake Wateree State Park' {Add-ADGroupMember -Credential $cred -Identity "Lake Wateree Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-482-6401"}
                'Landsford Canal State Park' {Add-ADGroupMember -Credential $cred -Identity "Landsford Canal Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-789-5800"}
                'Lee State Park' {Add-ADGroupMember -Credential $cred -Identity "Lee Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-428-5307"}
                'Little Pee Dee State Park' {Add-ADGroupMember -Credential $cred -Identity "Little Pee Dee Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-774-8872"}
                'Mountain Bridge Wilderness Area' {Add-ADGroupMember -Credential $cred -Identity "Mountain Bridge Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-836-6115"}
                'Musgrove Mill State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Musgrove Mill Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-938-0167"}
                'Myrtle Beach State Park' {Add-ADGroupMember -Credential $cred -Identity "Myrtle Beach Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-238-5325"}
                'Oconee State Park' {Add-ADGroupMember -Credential $cred -Identity "Oconee Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-638-5353"}
                'Oconee Station State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Oconee Station Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-638-0079"}
                'Paris Mountain State Park' {Add-ADGroupMember -Credential $cred -Identity "Paris Mountain Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-244-5565"}
                'Poinsett State Park' {Add-ADGroupMember -Credential $cred -Identity "Poinsett Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-494-8177"}
                'Redcliffe Plantation State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Redcliffe Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-827-1473"}
                'Rivers Bridge State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Rivers Bridge Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-267-3675"}
                'Rose Hill Plantation State Historic Site' {Add-ADGroupMember -Credential $cred -Identity "Rose Hill Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-427-5966"}
                'Sadlers Creek State Park' {Add-ADGroupMember -Credential $cred -Identity "Sadlers Creek Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-226-8950"}
                'Santee State Park' {Add-ADGroupMember -Credential $cred -Identity "Santee Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-854-2408"}
                'Sesquicentennial State Park' {Add-ADGroupMember -Credential $cred -Identity "Sesquicentenial Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-788-2706"}
                'Table Rock State Park' {Add-ADGroupMember -Credential $cred -Identity "Table Rock Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-878-9813"}
                'Woods Bay State Park' {Add-ADGroupMember -Credential $cred -Identity "Woods Bay Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-659-4445"}
                'State House Gift Shop & Tour Office' {Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-734-2430"}
                'Dillon WC' {Add-ADGroupMember -Credential $cred -Identity "Dillon Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-774-4711"}
                'Blackburg WC' {Add-ADGroupMember -Credential $cred -Identity "Blacksburg Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-839-6742"}
                'Fairplay WC' {Add-ADGroupMember -Credential $cred -Identity "Fair Play Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-972-3731"}
                'Fort Mill WC' {Add-ADGroupMember -Credential $cred -Identity "Fort Mill Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-548-2880"}
                'Hardeeville WC' {Add-ADGroupMember -Credential $cred -Identity "Hardeeville Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-784-3275"}
                'Landrum WC' {Add-ADGroupMember -Credential $cred -Identity "Landrum Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "864-457-2228"}
                'Little River WC' {Add-ADGroupMember -Credential $cred -Identity "Little River Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "843-249-1111"}
                'North Augusta WC' {Add-ADGroupMember -Credential $cred -Identity "North Augusta Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-279-6756"}
                'Santee WC' {Add-ADGroupMember -Credential $cred -Identity "Santee Welcome Center Security Group" -Members $SAM
                                    Set-ADUser -Credential $cred -Identity $SAM -OfficePhone "803-854-2442"}                     
            }
            #This switch adds park users to Parks-FIELD
            switch ($Division) {
                'State Parks - Coastal Region' {Add-ADGroupMember -Credential $cred -Identity "Parks-FIELD" -Members $SAM}
                'State Parks - Lakes Region' {Add-ADGroupMember -Credential $cred -Identity "Parks-FIELD" -Members $SAM}
                'State Parks - Mountain Region' {Add-ADGroupMember -Credential $cred -Identity "Parks-FIELD" -Members $SAM}
                'State Parks - Sandhills Region' {Add-ADGroupMember -Credential $cred -Identity "Parks-FIELD" -Members $SAM}
            }
            #This switch moves the AD User to the appropriate OU based on the Position Category and Division variables. Both are needed because an A2 could go in Parks, WC, or other OU
            switch ($PositionCategory) {
                'A2/Temporary' {
                    switch ($Division) {
                        'State Parks - Coastal Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Lakes Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Mountain Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Sandhills Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'Welcome Centers - Field' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2WelcomeCenters
                        }
                        'State Parks - Central Office' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Directorate' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Administration' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Tourism Sales & Marketing' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                    }
                 }
                'Volunteer/Campground Host' {
                    Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                }
                'A2/Temporary - Seasonal' {
                    switch ($Division) {
                        'State Parks - Coastal Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Lakes Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Mountain Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Sandhills Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'Welcome Centers - Field' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2WelcomeCenters
                        }
                        'State Parks - Central Office' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Directorate' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Administration' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Tourism Sales & Marketing' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                    }
                }
                'FTE' {
                    Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforFTE
                }
                'Intern' {
                    switch ($Division) {
                        'State Parks - Coastal Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Lakes Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Mountain Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'State Parks - Sandhills Region' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2Parks
                        }
                        'Welcome Centers - Field' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2WelcomeCenters
                        }
                        'State Parks - Central Office' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Directorate' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Administration' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                        'Tourism Sales & Marketing' {
                            Get-ADUser -Credential $cred $SAM | Move-ADObject -Credential $cred -TargetPath $OUforA2CentralOffice
                        }
                    }
                }
                
            }
    }

}
If ($NetworkLogon -eq 'Yes'){
    [PsCustomObject]@{
        SAM = $SAM
        FullName = $Displayname
        Password = $Password
        Supervisor = $ManagerName
        Location = $Office
        SentEmail = "No"
    } | Export-Csv -append -Path 'C:\Scripts\Logininfo\newLoginInfo.csv' -NoTypeInformation}
}

Move-Item -Path $getcsv -Destination 'C:\Scripts\Complete New Users'