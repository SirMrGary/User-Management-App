Step 1. Place the UserRoles.csv file where you want it (if its in sharepoint  you need to sync it via onedrive. If in SharePoint then multiple people can use the same source file).

Step 2. Update the Role_builder.ps1 file with the above location and clear the example content.

Step 3. Run the Role Builder.ps1 file to populate your current org roles (You will the need to filter them down and choose one of each role to use as a template)

Step 4. Update the Userroles.csv file location in the UserMangagement.ps1 file (Line 524)

Step 5. Update your Global Config information in the UserManagement.ps1 file (Line 13)

Step 6. Update your Office / Location data within the UserManagement.ps1 file ( from line 476)

Step 7. Update your company name on  line 1338 (this will show it at the top of the app "

All the hard work is now done and you can speed up your User Account Creation and Transfers.  

Everyone's exit processes are different, so feel free to finish this part yourself or leave it blank.
