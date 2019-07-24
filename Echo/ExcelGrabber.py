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

    with open('Users.csv', 'w') as f:
        f.write(','.join(newHeaders) + '\n')
        for course in fullCourseList:
            f.write(','.join(course) + '\n')


def csvSchedule(length):
    srchHeader = ['Drexel University', '3675 Market', 'Room Code', 'Ptrm Start Date', 'Begin Time', 'End Time', "D1|V1",
                  "Course", 'Term Code', 'Primary Instr Email Address', "Other Instr Email", "Yes", 'Day',
                  "Ptrm End Date", "HD", "[live stream place holder]", "[closed captioning placeholder]"]
    # Is it okay to merge COURSE and SECTION together? CS 260 001 instead of CS 260 | 001?
    # Are all courses repeating? I would assume so.
    # How to check if live stream?
    # Is closed captioning column used at all?
    newHeaders = ['Campus', 'Buildings', 'Room', 'Start Date', 'Recording Start Time', 'Recording End Time', 'Inputs',
                  'Title', 'Course Code', 'Section Code', 'Term', 'Instructor Email', 'Guest Instructor Email',
                  'Repeating', 'Repeat Patterns', 'End Date', 'Quality', 'Live Stream', 'Closed Captioning']
    location = findIndexes(srchHeader, headerRow)

    fullCourseList = []
    for row in length:
        courseLine = []
        for columnIndex in location:
            courseLine.append(worksheet.cell(row, columnIndex).value)
        fullCourseList.append(courseLine)

    f = open('Schedule.csv', 'w')
    for m in newHeaders:
        f.write(m + ',')
    f.write('\n')
    for l in fullCourseList:
        for columnIndex in l:
            if ',' in columnIndex:
                a, b = columnIndex.split(', ', 1)
                f.write(str(a) + ',')
                f.write(str(b) + ',')
            elif ' ' in columnIndex:
                columnIndex = columnIndex.split(' ')
                columnIndex = str(columnIndex[0] + ' ' + columnIndex[1])
                f.write(str(columnIndex) + ',')
            else:
                f.write(str(columnIndex) + ',')
        f.write('\n')
    f.close()


def csvCourses(length):
    headers = ('Primary Instr Email Address' 'Course', 'Term Code')
    newHeaders = ('Organization', 'Department', 'Course Code', 'Course Name',
                  'Section Code', 'Primary Instructor Email', 'Secondary Instructor Email')
    location = [6, 0, 62]
    data = []

    for j in length:
        temp = []
        for k in location:
            temp.append(worksheet.cell(j, k).value)
        data.append(temp)

    f = open('Courses.csv', 'w')
    for m in newHeaders:
        f.write(m + ',')
    f.write('\n')
    for l in data:
        for k in l:
            if ' ' in k:
                a, b, c = k.split(' ')
                f.write(str(a) + ',')
                f.write((a + ' ' + b) + ',')
                f.write((a + ' ' + b) + ',')
            else:
                f.write(str(k) + ',')
        f.write('\n')
    f.close()


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
