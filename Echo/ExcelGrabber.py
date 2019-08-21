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
    newHeaders = ['Role', 'Last Name', 'First Name', 'Email Address', 'Course Code', 'Section Code', 'Term']
    srchHeader = ['Primary Instructor', 'Primary Instr Email Address', 'Course', 'Term Code', 'Seq Numb']
    location = findIndexes(srchHeader, headerRow)

    fullCourseList = []
    for row in length:
        # TODO: "Instructor" is hardcoded here because it is not in the Echo spreadsheet.
        courseLine = ["Instructor"]
        for columnIndex in location:
            currentCell = worksheet.cell(row, columnIndex).value
            if ',' in currentCell:
                currentCell = currentCell.split(', ')
                courseLine.append(currentCell[0])
                courseLine.append(currentCell[1])
            else:
                courseLine.append(currentCell)
        fullCourseList.append(courseLine)

    with open('../output/Users.csv', 'w') as f:
        f.write(','.join(newHeaders) + '\n')
        for course in fullCourseList:
            f.write(','.join(course) + '\n')


def csvSchedule(length):
    # TODO: Fix the new headers so that they are all separate elements
    newHeaders = [
        "Campus,Building,Room,Start Date,Recording Start Time,Recording End Time,Inputs,Title,Course Code,Section Code,",
        "Term,Instructor Email,Guest Instructor Email,Repeating,Repeat Patterns,End Date,Quality,Live Stream,",
        "Closed Captioning"]
    srchHeader1 = ['Room Code', 'Ptrm Start Date', 'Begin Time', 'End Time']
    srchHeader2 = ["Course", 'Term Code', 'Primary Instr Email Address', "Other Instr Email"]
    srchHeader3 = ['Day']
    srchHeader4 = ["Ptrm End Date"]
    location1 = findIndexes(srchHeader1, headerRow)  # TODO: Do we need to pass through the headerRow from main?
    location2 = findIndexes(srchHeader2, headerRow)
    location3 = findIndexes(srchHeader3, headerRow)
    location4 = findIndexes(srchHeader4, headerRow)

    fullCourseList = []
    for row in length:
        courseLine = ["Drexel University", "3675 Market"]

        for columnIndex in location1:
            currentCell = worksheet.cell(row, columnIndex).value
            print(worksheet.cell(row, columnIndex))
            print(currentCell)
            # TODO: Properly format the time and date. We might have to splt the search header/queries so that time is
            #  on its own. We may have to change the "workbook mode".
            if '/' in currentCell:
                currentCell = currentCell.split('/')
                currentCell.insert(0, currentCell[-1])
                currentCell.pop()
                currentCell = '-'.join(currentCell)
            elif currentCell[-1] == "0" and not " " in currentCell and len(currentCell) == 4:
                currentCell = currentCell[0:2] + ":" + currentCell[2:] + ":00"
            courseLine.append(currentCell)

        courseLine.append("D1|V1")

        for columnIndex in location2:
            currentCell = worksheet.cell(row, columnIndex).value
            if " " in currentCell:
                courseNameBroken = currentCell.split(" ")
                currentCell = courseNameBroken[0] + " " + courseNameBroken[1]
                courseLine.append(currentCell)
            courseLine.append(currentCell)

        courseLine.append("Yes")

        currentCell = worksheet.cell(row, location3[0]).value
        currentCell = '|'.join(list(currentCell))
        courseLine.append(currentCell)

        currentCell = worksheet.cell(row, location4[0]).value
        currentCell = currentCell.split('/')
        currentCell.insert(0, currentCell[-1])
        currentCell.pop()
        currentCell = '-'.join(currentCell)
        courseLine.append(currentCell)
        # TODO: Hardcode append HD. If time persists, we will code the "Live Stream" column in as well (which requires
        #  crosschecking with Hyperion.xlsx. However, in the mean time, we shall leave it to the human.
        fullCourseList.append(courseLine)

    with open('../output/Schedule.csv', 'w') as f:
        f.write(','.join(newHeaders) + '\n')
        for course in fullCourseList:
            f.write(','.join(course) + '\n')


def csvCourses(length):
    srchHeaders = ['Course', 'Term Code', 'Primary Instr Email Address',
                   'Other Instr Email']  # removed  in list because i manipulated the data with brute force/manually
    newHeaders = ['Organization', 'Department', 'Course Code', 'Course Name', 'Section Code', 'Term', 'Section Code',
                  'Primary Instructor Email', 'Secondary Instructor Email']
    location = findIndexes(srchHeaders, headerRow)  # TODO: Do we need to pass through the headerRow from main?
    fullCourseList = []

    for row in length:
        courseLine = ['College of Computing & Informatics']
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

        with open('../output/Courses.csv', 'w') as f:
            f.write(','.join(newHeaders) + '\n')
            for course in fullCourseList:
                f.write(','.join(course) + '\n')


if __name__ == "__main__":
    type = input('What type of CSV file will this be? Select 1, 2, or 3.'
                 '\n1. Users'
                 '\n2. Schedule'
                 '\n3. Courses'
                 '\n: ')
    start = time.time()

    filename = sys.argv[1]
    workbook = xlrd.open_workbook(filename, on_demand=True)
    worksheet = workbook.sheet_by_index(0)
    headerRow = worksheet.row(0)
    length = range(1, worksheet.nrows)

    if type == '1':
        csvUser(length, headerRow)
    elif type == '2':
        csvSchedule(length)
    elif type == '3':
        csvCourses(length)

    # Finds the index of the desired columns
    # for i in headerRow:
    #    if i.value in headers:
    #        location.append(index)
    #    index += 1

    print("Finished in %s seconds" % (time.time() - start))
