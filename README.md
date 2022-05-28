# emailNewUser
There is a script before this one
Run the new user script that puts the user into the active directory
Make sure the new users have been licensed with an email before using the script

Set up:
--------
- Make sure you have the latest outlook client installed and you are logged in. No other outlook client can be on your machine.
- Since you ran the first script, you should have a folder in your C drive called 'Scripts'
- In this directory there should be a few more directories, 'Complete New Users', and 'LoginInfo'
- Please copy over the 'Attachments' directory and the 'userEmailGUIv2.ps1' in the 'Scripts' directory
- Any user in the CSV files in the 'Complete New Users' directory will be ran.
- If you have files in there that you would like to not be ran, Please create a new directory in 'Complete new users' called 'Sent New Users'
- Make sure you have your desired csv file in the 'Complete New Users' and you are ready to run the script



Instructions:
-------------
- Open powershell ise as an admin
- Use file -> open to open our script
- click the green arrow at the top
- It will ask for your credentials, please login to an account that can reset user passwords
- Minimize your powershell ise display and you should see the application window
- The users name aswell as where they are going should appear at the top aswell as their supervisors emails below
- If there is anything wrong, you can delete or add emails to the listbox
- To delete: Click on the desired email and click on delete
- To add: Click in the text field underneath the listbox and type your new email. Click add and that email should appear in the listbox
- Click on parks or welcome center radio button depending on where they are going
- If they are going to neither a park or a welcome center, Please skip this user
- You may also delete all emails from the email list and add yours so you can modify the email to yourself before you send it out.
- Click on yes or no if you would like to reset this users password when you send the email or skip the user
- Finally click send email and if successful check your sent mail and you should have a sent email.




MAKE SURE:
------------
- CSV file being used cannot be opened
- The close button (Top right X) Will not work as intended. Please do not press unless you absolutely want to close the program. Note: This will still reset password if that option is checked.
- You may get multiple errors if you do not have access to reset users passwords or if you have any outdated outlook client installed on your machine
