from processing import process
from mailing import send_mail
from datetime import date
import statistics as st
from barchart import chart, get_val
import os
import openpyxl.utils.exceptions
import sys
import pathlib

input_file = input('enter the name (along with .xlsx extension) of the excel file to process : ')
name_of_sheet = input(f"enter the name of the sheet to be processed in the excel file '{input_file}' : ")
save_file = input(f"enter the file name (along with .xlsx extension) under which you want to save the processed excel sheet '{input_file}' : ")

extension = pathlib.Path(save_file).suffix        #finding the extension of 'save_file'
if extension != '.xlsx':
    # sys.exit('[statement]') terminates the program by printing the [statement] to the output
    sys.exit("Invalid file name to save the processed Excel data. Should contain a '.xlsx' extension")

try:
    student_data = process(input_file, save_file, name_of_sheet)              #function call to process 'input_file' which returns a list of dictionaries, each for a student
except openpyxl.utils.exceptions.InvalidFileException:                        #exception for 'input_file' file name error - not containing '.xlsx' extension
    sys.exit("Invalid file name for the excel file to be processed. Should contain '.xlsx' extension")
except FileNotFoundError:                                                     #exception for 'input_file' not present in project folder
    sys.exit(f"Excel file to be processed i.e '{input_file}' is not found in project folder")
except KeyError:                                                              #exception for 'name_of_sheet' not found in 'input_file'
    sys.exit(f"No sheet '{name_of_sheet}' found in excel file '{input_file}'")
else:
    print(f"Excel file '{input_file}' processed successfully!")

acceptance = input("Do you want to send mails to the students? (yes/no) : ")
if acceptance == 'yes' or acceptance == 'YES' or acceptance == 'Yes':

    c1 = []
    c2 = []
    c3 = []
    a1 = []
    a2 = []
    c = []
    a = []
    inter = []
    for data in student_data:               #storing in lists, particular test data of all students
        c1.append(data.get('cie1'))
        c2.append(data.get('cie2'))
        c3.append(data.get('cie3'))
        a1.append(data.get('aat1'))
        a2.append(data.get('aat2'))
        c.append(data.get('cie'))
        a.append(data.get('aat'))
        inter.append(data.get('internals'))

    fr = os.environ.get('name')            #obtaining the name of sender's Google Account from the system environment
    if fr == None:
        sys.exit(f"No environment variable called 'name' found ")

    course = input("enter course name  : ")
    today = date.today()
    date = today.strftime("%d/%m/%Y")      #obtaining the date of sending students' marks through mail for date stamping

    att = ['CIE1', 'CIE2', 'CIE3', 'AAT1', 'AAT2', 'CIE', 'AAT', 'Internals'] #list to store different test headings
    val2 = [st.mean(c1), st.mean(c2), st.mean(c3), st.mean(a1), st.mean(a2), st.mean(c), st.mean(a), st.mean(inter)] #list to store mean marks of class for different tests
    val3 = [max(c1), max(c2), max(c3), max(a1), max(a2), max(c), max(a), max(inter)] #list to store highest marks in class for different tests
    val4 = [20, 20, 20, 5, 5, 40, 10, 50] #list to store maximum marks for different tests
    print("Logging into Google Account and sending mails.....")

    for data in student_data:                                 #accessing each student's dictionary in the list 'student_data' to send personalized mails to him/her
        val1 = get_val(data)                               #function call which returns a list containg the student's marks in all tests
        chart(att, val1, val2, val3, val4, data.get('name'))  #function call which draws the student's performance analysis chart(PAC)
        to = data.get('email')

        #below is the body of the email sent to students which is written in html to display marks in a tabular form
        html = f"""                                            
              <!DOCTYPE html>
              <html>
              <head></head>
              <body>
              <p>Please find your Internals marks and Performance graph for the course '{course}' below.</p>
              <br>
              <br>
              <table border = '5'>
              <tr>
              <th>Date</th>
              <th>Name</th>
              <th>Course</th>
              <th>CIE1 Marks(20)</th>
              <th>CIE2 Marks(20)</th>
              <th>CIE3 Marks(20)</th>
              <th>AAT1 Marks(5)</th>
              <th>AAT2 Marks(5)</th>
              <th>Total CIE Marks(40)</th>
              <th>Total AAT Marks(10)</th>
              <th>Total Internals Marks(50)</th>
              <th>Attendance(%)</th>
              <th>Eligibility for SEE</th>
              </tr>
              <tr>
              <th>{date}</th>
              <th>{data.get('name')}</th>
              <th>{course}</th>
              <th>{data.get('cie1')}</th>
              <th>{data.get('cie2')}</th>
              <th>{data.get('cie3')}</th>
              <th>{data.get('aat1')}</th>
              <th>{data.get('aat2')}</th>
              <th>{data.get('cie')}</th>
              <th>{data.get('aat')}</th>
              <th>{data.get('internals')}</th>
              <th>{data.get('attendance')}</th>
              <th>{data.get('eligibility')}</th>
              </tr>
              </table>
              <br>
              <br>
              <p>Regards,</p>
              <p>{fr}</p>
              </body>
              </html>
    """
        send_mail(fr, to, course, html, f"{data.get('name')}.png")        #function call that sends marks to students through mail along with their PAC as an attachment
        os.remove(f"{data.get('name')}.png")              #the PAC for each student, which is saved in the project folder, is deleted to eliminate garbage data for the user


    print(f"All Mails sent successfully!...Check for the updated excel file '{save_file}' in your Project Folder!\nThank You!")

else:
    print(f"Check for the updated excel file '{save_file}' in your Project Folder!\nThank You!")






