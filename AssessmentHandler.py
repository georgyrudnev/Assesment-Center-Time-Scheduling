
from os import system
import pandas as pd
import xlsxwriter as xs


df_inputTable = pd.read_excel('JunITer_Intergationsaufgabe_Georgy.xlsx', sheet_name=0)


print(df_inputTable)

# initialize some variables
bool_selInter1 = False
str_condGender = 'n'
str_condExp = 'n'



# 1.1 Save all Interviewer and Bewerber in seperate DF
df_interviewer = df_inputTable[(df_inputTable.Rolle == "Interviewer")]
df_bewerber = df_inputTable[(df_inputTable.Rolle == "Bewerber")]

# 1.2 Save the time capacity sheets in seperate DF 
df_firstDay = pd.read_excel('JunITer_Intergationsaufgabe_Georgy.xlsx', sheet_name=1)
df_secDay = pd.read_excel('JunITer_Intergationsaufgabe_Georgy.xlsx', sheet_name=2)
df_thirdDay = pd.read_excel('JunITer_Intergationsaufgabe_Georgy.xlsx', sheet_name=3)


xls = pd.read_excel('JunITer_Intergationsaufgabe_Georgy.xlsx', sheet_name = None)
str_firstDay = ""
str_secDay = ""
str_thirdDay = ""

j = 0
for y in xls.keys():
    if (j == 1):
        str_firstDay = y
    if (j == 2):
        str_secDay = y
    if (j == 3):
        str_thirdDay = y
        break
    j = j + 1

# 1.3 Prepare output Excel File

workbook = xs.Workbook('Output.xlsx')
worksheet = workbook.add_worksheet(str_firstDay)
worksheet2 = workbook.add_worksheet(str_secDay)
worksheet3 = workbook.add_worksheet(str_thirdDay)
# Start from the first cell.

row = 0
column = 0
 

# Get amount of applicants and interviewers for iteration
int_bewerberSize = df_bewerber[df_bewerber.columns[0]].count()
int_interviewerSize = df_interviewer[df_interviewer.columns[0]].count()

# 2. Iterate over the remaining Bewerber and find matching Inverviewer

int_interviewerIterator = 0 # Sorgt für gleichmäßige Zuteilung

# 3. In the loop, create the output excel and write the cells with matching solutions

print(df_interviewer.iloc[1, 0])

usez = -999
bool_firstDay = True
bool_secDay = True
bool_thirdDay = True
reservedInterviewer = int_interviewerSize+1

for i in range(0,int_bewerberSize):
    worksheet.write(row+i, column, df_bewerber.iloc[i, 0])
    # Wenn aktueller Interviewer (im iterator) Zeit hat
    z = 0
    infinloop = int_interviewerSize*int_interviewerSize
    print("Interview Iterator anfang äußere for-loop:")
    print(int_interviewerIterator)

    int_sumTriedInterviewer = 0
    
    while(z<5):
        int_sumTriedInterviewer = int_sumTriedInterviewer+1
        
        if (usez > -1):
            z = usez

        #print(" Dienstag:")
        #print(df_firstDay)

        #print("Mittwoch: ")
        #print(df_secDay)

        #print("Donnerstag: ")
        #print(df_thirdDay)


        if (df_firstDay.iloc[int_interviewerIterator, z+2].__contains__('+') & df_firstDay.iloc[int_interviewerSize+i, z+2].__contains__('+') & bool_firstDay):
            print("in first while")
            if (bool_selInter1 & (int_interviewerIterator != reservedInterviewer)):
                # in case first interview partner is already selected
                worksheet.write(row+i, column+1+z, df_interviewer.iloc[reservedInterviewer, 0] + ", " + df_interviewer.iloc[int_interviewerIterator, 0])
                df_firstDay.iloc[int_interviewerIterator, z+2] = "-" # set interviewer 2 time slot to -
                df_firstDay.iloc[reservedInterviewer, reservedTimeslot] = "-" # set interviewer 1 time slot to -
                reserverdInterviewer = int_interviewerSize+1
                usez = -999
                bool_selInter1 = False

                bool_secDay = True
                bool_thirdDay = True
                break

            # reserve time slot of interviewer 1
            reservedInterviewer = int_interviewerIterator
            reservedTimeslot = z+2

            usez = z
            z = 4
            bool_selInter1 = True
            bool_secDay = False
            bool_thirdDay = False
            
            
        


        elif (df_secDay.iloc[int_interviewerIterator, 2].__contains__('+') & df_secDay.iloc[int_interviewerSize+i, z+2].__contains__('+') & bool_secDay):
            if (bool_selInter1 & (int_interviewerIterator != reservedInterviewer)):
                # in case first interview partner is already selected
                worksheet2.write(row+i, column+1+z, df_interviewer.iloc[reservedInterviewer, 0] + ", " + df_interviewer.iloc[int_interviewerIterator, 0])
                df_secDay.iloc[int_interviewerIterator, z+2] = "-" # set interviewer 2 time slot to -
                df_secDay.iloc[reservedInterviewer, reservedTimeslot] = "-" # set interviewer 1 time slot to -
                reserverdInterviewer = int_interviewerSize+1

                usez = -999
                bool_selInter1 = False

                bool_firstDay = True
                bool_thirdDay = True
                break

            str_interviewer1 = df_interviewer.iloc[int_interviewerIterator, 0]

            # reserve time slot of interviewer 1
            reservedInterviewer = int_interviewerIterator
            reservedTimeslot = z+2
            

            usez = z
            z = 4
            bool_selInter1 = True
            bool_firstDay = False
            bool_thirdDay = False

        elif (df_thirdDay.iloc[int_interviewerIterator, 2].__contains__('+') & df_thirdDay.iloc[int_interviewerSize+i, z+2].__contains__('+') & bool_thirdDay):
            if (bool_selInter1 & (int_interviewerIterator != reservedInterviewer)):
                # in case first interview partner is already selected
                worksheet3.write(row+i, column+1+z, df_interviewer.iloc[reservedInterviewer, 0] + ", " + df_interviewer.iloc[int_interviewerIterator, 0])
                df_thirdDay.iloc[int_interviewerIterator, z+2] = "-" # set interviewer 2 time slot to -
                df_thirdDay.iloc[reservedInterviewer, reservedTimeslot] = "-" # set interviewer 1 time slot to -
                reserverdInterviewer = int_interviewerSize+1
                usez = -999
                bool_selInter1 = False

                bool_secDay = True
                bool_firstDay = True
                break

            # reserve time slot of interviewer 1
            reservedInterviewer = int_interviewerIterator
            reservedTimeslot = z+2

            usez = z
            z = 4
            bool_selInter1 = True
            bool_secDay = False
            bool_firstDay = False
     
        
        
        if (usez > -1):
            int_interviewerIterator = int_interviewerIterator+1
            if (int_interviewerIterator == int_interviewerSize):
                int_interviewerIterator = 0
        else:
            z = z+1
            if (z == 5):
             # reset time index and try next interviewer
                z = 0
                int_interviewerIterator = int_interviewerIterator+1
                if (int_interviewerIterator == int_interviewerSize):
                    int_interviewerIterator = 0
        # in case no fitting time slot is found
        if (int_sumTriedInterviewer > int_interviewerSize):
            reserverdInterviewer = int_interviewerSize+1
            usez = -999
            bool_selInter1 = False
            bool_thirdDay = True
            bool_secDay = True
            bool_firstDay = True
            z = z+1

        infinloop = infinloop - 1
        if (infinloop == 0):
            print(df_thirdDay)
            print(bool_firstDay)
            print(bool_secDay)
            print(bool_thirdDay)
            print(z)
            print(int_interviewerIterator)

            raise Exception("Infinite Loop catched. Check if there really are enough timeslots available")
            


    # Bei dem zweiten Interview Partner müssen Bedingungen beachtet werden (versch. Geschlechter und unerfahren+erfahren)


    int_interviewerIterator= int_interviewerIterator+1
    # Setze interviewerIterator wieder zurück, falls alle interviewer einmal dran waren
    if (int_interviewerIterator == int_interviewerSize):
        int_interviewerIterator = 0
    str_condGender = 'n'
    str_condExp = 'n'
    print("Interview Iterator Ende äußere for-loop:")
    print(int_interviewerIterator)


workbook.close()