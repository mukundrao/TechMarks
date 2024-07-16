Steps to setup the project:

1. Clone this repository to your local system
2. Create a virtual environment and activate it
3. Open a terminal and run pip install requirements.txt to install the dependencies
4. Setup the system environment variables on your system by navigating to 'Edit the system environment variables'. Navigate to Environment Variables -> New (User environment variables):
        4.1. Name :
            • Variable Name : ‘name’
            • Variable value : [Sender’s Google Account name]
        4.2. Email ID:
            • Variable Name : ‘mailid’
            • Variable value : [Sender’s Gmail ID]
        4.3. Password:
            • Variable Name : ‘sagpwd’
            • Variable value : [App password created for the sender’s Google Account]
5. Close the IDE/terminal and restart it
6. Navigate to the project folder and activate the virtual environment
7. Configure the marks1.xlsx file by adding your marksheet content to it and the students' email IDs and save your changes
8. Run the main.py file using the IDE GUI or if using the terminal, run the command python main.py
9. Give the necessary inputs
10. Wait for the mails to be sent...the processed marksheet can be viewed in the project folder 
11. Verify that the mails have been sent and the marksheet has been processed correctly...it takes around 4 seconds to mail 1 student
12. Hurray! You've set up the project on your local system!

Steps to obtain 'sagpwd' - app password for Sender's Google Account:

1. Go to your Google Account Profile
2. Navigate to Security -> 2-Step Verification
3. Turn on 2-Step Verification
4. Navigate to the search bar and search for 'App passwords'
5. Create an app password by giving a name for the app
6. Copy the app password generated to be saved as a system environment variable under the name - 'sagpwd'

Python Libraries to be Installed :

1. numpy
2. openpyxl
3. matplotlib
4. smtplib


