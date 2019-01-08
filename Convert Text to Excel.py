import xlsxwriter
# file = open('testing.txt', 'r')
# Activity = file.readlines()
# file.close()

def main():
    inname = input("Input txt file name: ")
    outname = input("Name output Excel file: ")
    file = open(inname, 'r')    # default open for reading
    Activity = file.readlines()
    file.close()

    Activitylog = []
    for line in Activity:
        # Strip whitespace, should leave nothing if empty line was just "\n"
        if not line.strip():
            continue
        # We got something, save it
        else:
            Activitylog.append(line)

    #print (Activitylog)

    print("Dividing into blocks... \n")
    curr = 0
    ind = 0
    block = []
    for line in Activitylog:
        if Activitylog[curr] == 'Assistant\n':
            if Activitylog[ind] != 'Unknown voice command\n' and Activitylog[ind+1] != 'Something went wrong. Please try again later.\n':
                block.append(Activitylog[ind:curr])
            ind = curr + 1
            curr += 1
        else:
            curr += 1

    print("Categorizing... \n")
    question = []
    answer = []
    timestamp = []
    info = []
    length = len(block)
    block.reverse()
    for i in range(length):
        question.append(block[i][0])
        for j in range(len(block[i])):
            if block[i][j] == 'Products:\n':
                timestamp.append(block[i][j - 1])
        for j in range(len(block[i])):
            if block[i][j] == 'Products:\n':
                temp = []
                for num in range(1, j - 1):
                    temp.append(block[i][num])
                str = ' '.join(temp)
                answer.append(str)
        if '\u2003From unrecognized speaker\n' in block[i]:
            info.append("unrecognized speaker")
        else:
            info.append(" ")


    # Excel
    workbook = xlsxwriter.Workbook(outname)
    worksheet1 = workbook.add_worksheet()
    worksheet1.write('A1', "Questions")
    worksheet1.write('B1', "Answers")
    worksheet1.write('C1', "Timestamp")
    worksheet1.write('D1', "Speaker info")

    worksheet1.write_column('A2', question)
    worksheet1.write_column('B2', answer)
    worksheet1.write_column('C2', timestamp)
    worksheet1.write_column('D2', info)

    workbook.close()
    print("Conversion complete!")
# run it
main()

