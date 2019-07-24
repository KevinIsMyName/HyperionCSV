import xlrd
import sys
import time


def csvUser(length, headers):

    newHeaders = ('Role', 'Last Name', 'First Name', 'Email Address', 'Course Code', 'Section Code', 'Term')
    hardlocation = [59, 62, 6, 0, 5]
    location = []
    header = ['Primary Instructor', 'Primary Instr Email Address', 'Course', 'Term Code', 'Seq Numb']
    data = []
    index = 0
    # Finds the index of the desired columns

    for i in headers:
        if i.value in header:
            print(i)
            location.append(index)
        index += 1

    print(location)
    for j in length:
        temp = []
        for k in location[:-1]:
            temp.append(worksheet.cell(j, k).value)
        sectionCode = worksheet.cell(j, 5).value
        temp.append(worksheet.cell(j, 0).value + ' ' + sectionCode)
        temp[3], temp[4] = temp[4], temp[3]
        data.append(temp)

    f = open('Users.csv', 'w')
    for m in newHeaders[:-1]:
        f.write(m + ',')
    else:
        f.write(newHeaders[-1])
    f.write('\n')
    for l in data:
        for k in l[:-1]:
            if ',' in k:
                f.write('Instructor' + ',')
                a, b = k.split(', ')
                f.write(a + ',')
                f.write(b + ',')
            else:
                f.write(k + ',')
        else:
            f.write(l[-1])
        f.write('\n')
    f.close()


def csvSchedule(length):

    headers = ('Primary Instructor', 'Primary Instr Email Address' 'Course', 'Term Code')
    newHeaders = ('Term Code', 'Last Name', 'First Name', 'Email Address', 'Course Code')
    location = [0, 59, 62, 6]
    data = []

    for j in length:
        temp = []
        for k in location:
            temp.append(worksheet.cell(j, k).value)
        data.append(temp)

    f = open('Schedule.csv', 'w')
    for m in newHeaders:
        f.write(m + ',')
    f.write('\n')
    for l in data:
        for k in l:
            if ',' in k:
                a,b = k.split(', ', 1)
                f.write(str(a) + ',')
                f.write(str(b) + ',')
            elif ' ' in k:
                k = k.split(' ')
                k = str(k[0] + ' ' + k[1])
                f.write(str(k) + ',')
            else:
                f.write(str(k) + ',')
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
    length = worksheet.nrows
    length = range(1, length)

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
