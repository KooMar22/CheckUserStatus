# CheckUserStatus
This app performs the user status check (active or not) in command prompt by using the "net user" command.

It allows users to select an Excel file containing users and, when they run the "Pokreni" command, it performs the check and returns the result in that Excel file.

It gets the values from "Username" column and returns the status to "Account Status" column, regardless of their column positions.

This app accepts both "user" and "domain/user" input within the Excel file, ignoring the "domain/" part and taking only the username to perform the check.

If you are checking the users within a domain, be sure to include the "/domain" or applicable domain name after the "{username}" and as explained within the comment, else you'll get wrong results.
