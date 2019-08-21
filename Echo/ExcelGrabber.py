import xlrd
import sys
import time


def findIndexes(srchTerm, headerRow):
    location = []
    while not len(srchTerm) is 0:
        index = 0
        srchFound = False
        for cell in headerRow:
            if cell.value == srchTerm[0]:
                location.append(index)
                srchFound = True
                srchTerm.pop(0)
                break
            index += 1
        if not srchFound:
            location.append(srchTerm[0])
            srchTerm.pop(0)
    return location


def csvUser(length, headerRow):
    newHeaders = ["Role", "Last Name", "First Name", "Email Address", "Course Code", "Section Code", "Term"]
    srchHeader = ["Primary Instructor", "Primary Instr Email Address", "Course", "Term Code", "Seq Numb"]
    location = findIndexes(srchHeader, headerRow)

    fullCourseList = []
    for row in length:
        # TODO: "Instructor" is hardcoded here because it is not in the Echo spreadsheet.
        courseLine = ["Instructor"]
        for columnIndex in location:
            currentCell = worksheet.cell(row, columnIndex).value
            if "," in currentCell:
                currentCell = currentCell.split(", ")
                courseLine.append(currentCell[0])
                courseLine.append(currentCell[1])
            else:
                courseLine.append(currentCell)
        fullCourseList.append(courseLine)

    with open("../output/Users.csv", "w") as f:
        f.write(",".join(newHeaders) + "\n")
        for course in fullCourseList:
            f.write(",".join(course) + "\n")


def csvSchedule(length, headerRow):
    # TODO: Fix the new headers so that they are all separate elements
    newHeaders = [
        "Campus", "Building", "Room", "Start Date", "Recording Start Time", "Recording End Time", "Inputs", "Title",
        "Course Code", "Section Code",
        "Term", "Instructor Email", "Guest Instructor Email", "Repeating,Repeat Patterns", "End Date,Quality",
        "Live Stream",
        "Closed Captioning"]
    srchHeader1 = ["Room Code"]
    srchHeader2 = ["Ptrm Start Date"]
    srchHeader3 = ["Begin Time", "End Time"]
    srchHeader4 = ["Course"]
    srchHeader5 = ["Term Code", "Primary Instr Email Address", "Other Instr Email"]
    srchHeader6 = ["Day"]
    srchHeader7 = ["Ptrm End Date"]
    location1 = findIndexes(srchHeader1, headerRow)
    location2 = findIndexes(srchHeader2, headerRow)
    location3 = findIndexes(srchHeader3, headerRow)
    location4 = findIndexes(srchHeader4, headerRow)
    location5 = findIndexes(srchHeader5, headerRow)
    location6 = findIndexes(srchHeader6, headerRow)
    location7 = findIndexes(srchHeader7, headerRow)

    fullCourseList = []
    for row in length:
        courseLine = ["Drexel University", "3675 Market"]
        for columnIndex in location1:
            currentCell = worksheet.cell(row, columnIndex).value
            courseLine.append(currentCell)

        for columnIndex in location2:
            currentCell = worksheet.cell(row, columnIndex).value
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(currentCell, workbook.datemode)
            if len(str(month)) == 1:
                month = "0" + str(month)
            if len(str(day)) == 1:
                day = "0" + str(day)
            date = str(year) + "-" + str(month) + "-" + str(day)
            courseLine.append(date)

        for columnIndex in location3:
            currentCell = worksheet.cell(row, columnIndex).value
            hour = currentCell[:2]
            minute = currentCell[2:]
            second = "00"
            time = str(hour) + ":" + str(minute) + ":" + str(second)
            courseLine.append(time)

        courseLine.append("D1|V1")

        for columnIndex in location4:
            currentCell = worksheet.cell(row, columnIndex).value
            course = currentCell.split()
            course = course[0] + " " + course[1]
            courseLine.append(course)
            courseLine.append(course)

        for columnIndex in location5:
            currentCell = worksheet.cell(row, columnIndex).value
            courseLine.append(currentCell)

        courseLine.append("Yes")

        for columnIndex in location6:
            currentCell = worksheet.cell(row, columnIndex).value
            days = currentCell
            days = days.replace("M", "onday")
            days = days.replace("T", "uesday")
            days = days.replace("W", "ednesday")
            days = days.replace("R", "hursday")
            days = days.replace("F", "riday")
            days = days.replace("onday", "Monday")
            days = days.replace("uesday", "Tuesday")
            days = days.replace("ednesday", "Wednesday")
            days = days.replace("hursday", "Thursday")
            days = days.replace("riday", "Friday")
            days = 'y|'.join(days.split('y'))
            days = days[0:len(days) - 1]

            courseLine.append(days)

        for columnIndex in location7:
            currentCell = worksheet.cell(row, columnIndex).value
            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(currentCell, workbook.datemode)
            date = str(year) + "-" + str(month) + "-" + str(day)
            courseLine.append(date)

        courseLine.append("HD")

        fullCourseList.append(courseLine)

    with open("../output/Schedule.csv", "w") as f:
        f.write(",".join(newHeaders) + "\n")
        for course in fullCourseList:
            f.write(",".join(course) + ",,\n")


def csvCourses(length):
    srchHeaders = ["Course", "Term Code", "Primary Instr Email Address",
                   "Other Instr Email"]  # removed  in list because i manipulated the data with brute force/manually
    newHeaders = ["Organization", "Department", "Course Code", "Course Name", "Section Code", "Term", "Section Code",
                  "Primary Instructor Email", "Secondary Instructor Email"]
    location = findIndexes(srchHeaders, headerRow)  # TODO: Do we need to pass through the headerRow from main?
    fullCourseList = []

    for row in length:
        courseLine = ["College of Computing & Informatics"]
        fullCourseName = worksheet.cell(row, location[0]).value
        fullCourseName = fullCourseName.split()
        dept = fullCourseName[0]
        courseNum = fullCourseName[1]
        courseSect = fullCourseName[2]
        courseName = dept + " " + courseNum
        term = worksheet.cell(row, location[1]).value
        sectionCode = term + " " + courseSect

        courseLine.append(dept)
        courseLine.append(courseName)
        courseLine.append(
            courseName)  # again, because one is for "Course Code" and the other is for "Course Name" which are identical
        courseLine.append(term)
        courseLine.append(sectionCode)
        courseLine.append(worksheet.cell(row, location[2]).value)
        courseLine.append(worksheet.cell(row, location[3]).value)

        fullCourseList.append(courseLine)

        with open("../output/Courses.csv", "w") as f:
            f.write(",".join(newHeaders) + "\n")
            for course in fullCourseList:
                f.write(",".join(course) + "\n")


if __name__ == "__main__":
    type = input("What type of CSV file will this be? Select 1, 2, or 3."
                 "\n1. Users"
                 "\n2. Schedule"
                 "\n3. Courses"
                 "\n: ")
    start = time.time()

    filename = sys.argv[1]
    workbook = xlrd.open_workbook(filename, on_demand=True)
    worksheet = workbook.sheet_by_index(0)
    headerRow = worksheet.row(0)
    length = range(1, worksheet.nrows)

    if type == "1":
        csvUser(length, headerRow)
    elif type == "2":
        csvSchedule(length, headerRow)
    elif type == "3":
        csvCourses(length, headerRow)

    # Finds the index of the desired columns
    # for i in headerRow:
    #    if i.value in headers:
    #        location.append(index)
    #    index += 1

    print("Finished in %s seconds" % (time.time() - start))
