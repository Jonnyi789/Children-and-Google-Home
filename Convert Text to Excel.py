import xlsxwriter
#file = open('Test 101.txt', 'r')
#Activity = file.readlines()
#file.close()

def main():
    inname = input("Input txt file name: ")
    outname = input("Name output Excel file: ")
    file = open(inname, 'r')    # default open for reading
    Activity = file.readlines()
    file.close()

    Activitylog = []
    for line in Activity:
        # Strip whitespace; remove leading and trailing whitespaces
        if not line.strip():    # ignore if it's an empty line "\n"
            continue
        # Add to the list
        else:
            Activitylog.append(line)
    #print(Activitylog)

    print("Dividing into blocks... \n")
    curr = 0
    fst = 0
    block = []
    while curr < len(Activitylog):  # First line "Assistant\n" needs to be moved to the end of .txt file.
    #for line in Activitylog:
        if Activitylog[curr] == 'Assistant\n':
            if Activitylog[fst] != 'Unknown voice command\n' and Activitylog[fst+1] != 'Something went wrong. Please try again later.\n':
                block.append(Activitylog[fst:curr]) # Discard useless blocks
            fst = curr + 1
            curr += 1
        else:
            curr += 1

    print("Categorizing... \n")
    question = []
    answer = []
    timestamp = []
    speaker = []
    length = len(block)
    block.reverse() # chronological
    for i in range(length):
        question.append(block[i][0])
        for j in range(len(block[i])):
            if block[i][j] == 'Products:\n':
                timestamp.append(block[i][j - 1])
        for j in range(len(block[i])):
            if block[i][j] == 'Products:\n':
                temp = []
                for line in range(1, j - 1):
                    temp.append(block[i][line])
                str = ' '.join(temp)    # ["A", "B", "C"] into "A B C"
                answer.append(str)
        if ' From unrecognized speaker\n' in block[i]:
            speaker.append("unrecognized speaker")
        else:
            speaker.append("Participant")


    # Excel generation
    workbook = xlsxwriter.Workbook(outname)
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', "Questions")
    worksheet.write('B1', "Answers")
    worksheet.write('C1', "Timestamp")
    worksheet.write('D1', "Speaker info")

    worksheet.write_column('A2', question)
    worksheet.write_column('B2', answer)
    worksheet.write_column('C2', timestamp)
    worksheet.write_column('D2', speaker)

    workbook.close()
    print("Conversion completed!")
# run it
main()
