# Kevin Wu (kw875@drexel.edu)
# https://gitlab.cci.drexel.edu/wfr23/CCI-Commons
# HyperionCSV.py
# Converted a formatted spreadsheet listing classes and schedules to .csv's for Echo360 not in UTF-8 format
# Made for Drexel University use

import sys
import os
import xlrd

# load file from cmd line
try:
    file_name = sys.argv[1]
    workbook = xlrd.open_workbook(file_name, on_demand=False)
except IndexError:
    print("No file name provided")
    print("Usage: python HyperionCSV.py <file_name>")
    sys.exit(-1)
except FileNotFoundError:
    print("File not found")
    print("Usage: python HyperionCSV.py <file_name>")
    sys.exit(-1)

worksheet = workbook.sheet_by_index(0)
header_row = worksheet.row(0)

# Get all indices of needed columns
email_i = -1
course_i = -1
seq_num_i = -1
subj_code_i = -1
crse_num_i = -1
other_email_i = -1
room_code_i = -1
day_i = -1
start_date_i = -1
rec_start_i = -1
rec_end_i = -1
term_i = -1
end_date_i = -1

i = 0
while i < len(header_row):
    cell = header_row[i]
    if cell.value == "Primary Instr Email Address":
        email_i = i
    elif cell.value == "Course":
        course_i = i
    elif cell.value == "Seq Numb":
        seq_num_i = i
    elif cell.value == "Subj Code":
        subj_code_i = i
    elif cell.value == "Crse Numb":
        crse_num_i = i
    elif cell.value == "Other Instr Email":
        other_email_i = i
    elif cell.value == "Room Code":
        room_code_i = i
    elif cell.value == "Day":
        day_i = i
    elif cell.value == "Ptrm Start Date":
        start_date_i = i
    elif cell.value == "Begin Time":
        rec_start_i = i
    elif cell.value == "End Time":
        rec_end_i = i
    elif cell.value == "Term Code":
        term_i = i
    elif cell.value == "Ptrm End Date":
        end_date_i = i
    i += 1

# TODO: Make users.csv
with open(file_name + " users.csv", "w") as output:
    output.write("Role,Last Name,First Name,Email Address,Course Code,Section Code,Term\n")
    role = "Instructor"
    for i in range(1, worksheet.nrows):
        output.write("Instructor" + ",")  # Role
        email = worksheet.cell(i, email_i).value.replace("@drexel.edu", "").split(".")  # Email Address
        first_name = email[0][0].upper() + email[0][1:]
        last_name = email[-1][0].upper() + email[-1][1:]
        output.write(last_name + ",")  # Last Name
        output.write(first_name + ",")  # First Name
        output.write(worksheet.cell(i, email_i).value + ",")  # Email Address
        output.write(
            worksheet.cell(i, subj_code_i).value + " " + worksheet.cell(i, crse_num_i).value + ",")  # Course Code
        output.write(worksheet.cell(i, term_i).value + " " + worksheet.cell(i, seq_num_i).value + ",")  # Section Code
        output.write(str(worksheet.cell(i, term_i).value) + ",")  # Term
        output.write("\n")

# TODO: Make schedules.csv
with open(file_name + " schedules.csv", "w") as output:
    output.write("Campus,Building,Room,Start Date,Recording Start Time,Recording End Time,Inputs,Title,Course Code,"
                 "Section Code,Term,Instructor Email,Guest Instructor Email,Repeating,Repeat Patterns,End Date,Quality,"
                 "Live Stream,Closed Captioning\n")
    for i in range(1, worksheet.nrows):
        output.write("Drexel University,")  # Campus
        output.write("3675 Market,")  # Building
        if worksheet.cell(i, room_code_i).value == "1054-1055":  # Room
            output.write("1055 Primary,")
        else:
            output.write(worksheet.cell(i, room_code_i).value + ",")
        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(worksheet.cell(i, start_date_i).value,
                                                                      workbook.datemode)  # Start Date
        if len(str(month)) == 1:
            month = "0" + str(month)
        if len(str(day)) == 1:
            day = "0" + str(day)
        start_date = str(year) + "-" + str(month) + "-" + str(day)
        output.write(start_date + ",")

        # Recording Start Time
        start_time = worksheet.cell(i, rec_start_i).value
        output.write(start_time[0:2] + ":" + start_time[2:4] + ":00,")
        # Recording End Time
        end_time = worksheet.cell(i, rec_end_i).value
        output.write(start_time[0:2] + ":" + start_time[2:4] + ":00,")

        output.write("D1|V1,")  # Inputs
        output.write(worksheet.cell(i, subj_code_i).value + " " + worksheet.cell(i, crse_num_i).value + ",")  # Title
        output.write(
            worksheet.cell(i, subj_code_i).value + " " + worksheet.cell(i, crse_num_i).value + ",")  # Course Code
        output.write(worksheet.cell(i, term_i).value + " " + worksheet.cell(i, seq_num_i).value + ",")  # Section Code
        output.write(worksheet.cell(i, term_i).value + ",")  # Term
        output.write(worksheet.cell(i, email_i).value + ",")  # Primary Instructor Email
        output.write(worksheet.cell(i, other_email_i).value + ",")  # Guest Instructor Email
        output.write("Yes,")  # Repeating
        output.write('|'.join(list(worksheet.cell(i, day_i).value)) + ",")  # Repeat Patterns

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(worksheet.cell(i, end_date_i).value,
                                                                      workbook.datemode)  # End Date
        if len(str(month)) == 1:
            month = "0" + str(month)
        if len(str(day)) == 1:
            day = "0" + str(day)
        end_date = str(year) + "-" + str(month) + "-" + str(day)
        output.write(end_date + ",")

        output.write("HD,")  # Quality
        output.write(",")  # Live Stream
        output.write("")  # Closed Captioning
        output.write("\n")

# TODO: Make courses.csv
with open(file_name + " courses.csv", "w") as output:
    output.write(
        "Organization,Department,Course Code,Course Name,Term,Section Code,Primary Instructor Email,"
        "Secondary Instructor Email\n")
    for i in range(1, worksheet.nrows):
        output.write("College of Computing & Informatics" + ",")  # Organization
        output.write(worksheet.cell(i, subj_code_i).value + ",")  # Department
        output.write(
            worksheet.cell(i, subj_code_i).value + " " + worksheet.cell(i, crse_num_i).value + ",")  # Course Code
        output.write(
            worksheet.cell(i, subj_code_i).value + " " + worksheet.cell(i, crse_num_i).value + ",")  # Course Name
        output.write(worksheet.cell(i, term_i).value + ",")  # Term
        output.write(worksheet.cell(i, term_i).value + " " + worksheet.cell(i, seq_num_i).value + ",")  # Section Code
        output.write(worksheet.cell(i, email_i).value + ",")  # Primary Instructor Email
        output.write(worksheet.cell(i, other_email_i).value + ",")  # Secondary Instructor Email
        output.write("\n")

    # Any CS or INFO course that is 500 or above for the course number,
    # & where the following is true:
    # There is an 001 and a 900 section being taught by the same instructor.
    # Or there is an 002 and a 901 section being taught by the same instructor.
