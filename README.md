<h3>AutoReporter - automate your Outlook and Access Database</h3>

AutoReporter is a tool, which automates daily tasks, that are being repeated over and over again.
If a certain pattern occurs in e-mails, it can be easily accessed with Python and be scanned with regex.

Python script looks for a certain keywords in messages subjects in order to find repeatable pattern.
It then uses regex patterns to find what is important in such messages and makes a list of with the results. 
Then, based on what he has collected, he looks for the data in Access Database.
After that he takes the owners of the data from Active Directory and in the end displays all the data to the end user.
It saves around 30mins-1hr a day and allows you to do things, that cannot be actually automated.

Requirements (additional packages):
- win32com.client (to connect to outlook)
- pyad (to connect to Active Directory)
- tabulate (to nicely display the data)

