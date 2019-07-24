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
    newHeaders = [
        "Campus,Building,Room,Start Date,Recording Start Time,Recording End Time,Inputs,Title,Course Code,Section Code,",
        "Term,Instructor Email,Guest Instructor Email,Repeating,Repeat Patterns,End Date,Quality,Live Stream,",
        "Closed Captioning"]
    srchHeader1 = ['Room Code', 'Ptrm Start Date', 'Begin Time', 'End Time']
    srchHeader2 = ["Course", 'Term Code', 'Primary Instr Email Address', "Other Instr Email"]
    srchHeader3 = ['Day']
    srchHeader4 = ["Ptrm End Date"]
    location1 = findIndexes(srchHeader1, headerRow)
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
        coursLine.append(currentCell)

        fullCourseList.append(courseLine)

    with open('Schedule.csv', 'w') as f:
        f.write(','.join(newHeaders) + '\n')
        for course in fullCourseList:
            f.write(','.join(course) + '\n')

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
