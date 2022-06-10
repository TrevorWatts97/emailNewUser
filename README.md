# New Active Directory User
This script has been created by a former colleague and myself. 
There are two scripts.
Please read through this before attempting run. Thank you.

Set up:
--------
- Make sure you have the latest outlook client installed and you are logged in. No other outlook client can be on your machine.
- Copy the 'Scripts' Folder into your C drive
- Copy over the Onboarding CSV file 'PTF Onboarding EOD Report - MM/DD/YYYY' into the root of the scripts folder
- Check the csv file for any missing fields and then close it once finished
- Any CSV in the scripts folder will be ran at this point on, so if you don't want it ran, please drag it into the 'complete new users' folder


CSV ACCOUNT CREATION Instructions:
----------------------------------
- Open powershell ise as an admin
- Use file -> open to open the 'CSV Account Creation' script
- click the green arrow at the top
- It will ask for your credentials, please login to an account that can reset user passwords
- The script creates a first-attempt user ID based on the user's first and last name. It respnds that it cannot find it, then creates it. THIS IS NOT AN ERROR
- The script then prompts to enter the username for each supervisor of each user account in the CSV list.
- Enter the user ID for the supervisor, looking it up in AD to confirm if necessary.
- Script completes and ends in command prompt.
- Script moves the .CSV file to the 'Complete New Users' folder

- Check AD to see if the user accounts were created. Verify their settings match with the intended settings from the .CSV file.
  - Member of tab: "Domain Users," as well as "Parks-FIELD" (if in a park) and the actual park group
  - General tab: Name, Description, Office, Phone Number, Email 
  - Profile tab: Logon script "logon.bat"
  - Organization tab: Job Title, Department code, Company SCPRT, Manager Name

PLEASE NOTE: 
-------------
  - In one instance the manager's user ID was entered and it caused an error, resulting in no new user creation
  - I checked and the manager's user ID was listed as one thing under "User logon name" and another under "E-mail" and "User logon name (pre-windows 2000)"
  - the correct one was either the ID listed in the email or the one in the pre-Win2k logon name, NOT the User logon name
  - not sure where this discrepancy happened but this is a worthwhile note for which version the script is checking against




USER EMAIL Instructions:
------------------------
- Before using this script, Please wait about an hour after users have been liscensed for their email.
- Check the LoginInfo csv file inside the Login Info folder
- Make sure only the users you want to send emails about have their 'SentEmail' set to 'No'
- Open powershell ise as an admin
- Use file -> open to open the 'CSV Account Creation' script
- Click the green arrow at the top
- It will ask for your credentials, please login to an account that can reset user passwords
- Minimize your powershell ise display and you should see the application window
- The users name aswell as where they are going should appear at the top aswell as their supervisors emails below
- Check AD for discrepancies here. If there is anything wrong, you can delete or add emails to the listbox
- To delete: Click on the desired email and click on delete
- To add: Click in the text field underneath the listbox and type your new email. Click add and that email should appear in the listbox
- Click on parks or welcome center radio button depending on where they are going
- If they are going to neither a park or a welcome center, Please skip this user
- You may also delete all emails from the email list and add yours so you can modify the email to yourself before you send it out.
- Click on yes or no if you would like to reset this users password when you send the email or skip the user
- Finally click send email and if successful check your sent mail and you should have a sent email.


PLEASE NOTE:
------------
- CSV file being used cannot be opened
- The close button (Top right X) Will not work as intended. If you would like to close out of the program, open the ise and click the stop button (red square) at the top. Check the do not reset password radio button and then press the x button at the top right. Note: This will still reset password if that option is checked.
- You may get multiple errors if you do not have access to reset users passwords or if you have any outdated outlook client installed on your machine
- This script relies on the previous scripts LoginInfo.csv file. Before you run the script, make sure only the users you want to send an email about have the 'SentEmail' status set to 'No'
