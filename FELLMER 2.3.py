import openpyxl
from decimal import *


def main():
    Excel = openpyxl.load_workbook('FELLMER.xlsx')

    Week = input("Δώσε μου εβδομάδα (Π.Χ 'ΕΒΔΟΜΑΔΑ 1', 'ΕΒΔΟΜΑΔΑ 2', ... 'ΕΒΔΟΜΑΔΑ 5')\n")

    sheet = Excel[Week]
    #sheet = Excel['ΕΒΔΟΜΑΔΑ 3']
    tuple(sheet['B9' : 'H27'])

    Counter = 9

    for rowOfCellObjects in sheet['B9' : 'H29']:
        
        hours = 0
        int_hours = 0
        count_single = 0
        count_normal = 0
        sunday_hours = 0

        Col_J = 'J'
        Col_K = 'K'
        Col_L = 'L'
        Col_M = 'M'

        int_beg1 = 0
        int_end1 = 0
        int_beg2 = 0
        int_end2 = 0
        
        a = 0
        b = 0
        
        try:

            for cellObj in rowOfCellObjects:

                time = cellObj.value
                
                if(cellObj.fill.start_color.index == 'FFFF0000'):
                    if(time is not None):
                        if(' ' in time):
                            a,b = time.split(' ')

                            beg1,end1 = a.split('--')
                            beg2,end2 = b.split('--')

                            f_beg1 = float(beg1)
                            f_end1 = float(end1)
                            f_beg2 = float(beg2)
                            f_end2 = float(end2)

                            sunday_hours = sunday_hours + ((f_end1 - f_beg1) + (f_end2 - f_beg2))

                        else:
                            beg,end = time.split('--')

                            f_beg = float(beg)
                            f_end = float(end)

                        if((f_end - f_beg) <= 1):
                            sunday_hours = sunday_hours + 1

                        else:
                            sunday_hours = sunday_hours + (f_end - f_beg)
                    
                time = cellObj.value

                if(time is not None):

                    if(' ' in time):
                        
                        a,b = time.split(' ')

                        beg1,end1 = a.split('--')
                        beg2,end2 = b.split('--')

                        f_beg1 = float(beg1)
                        f_end1 = float(end1)
                        f_beg2 = float(beg2)
                        f_end2 = float(end2)

                        count_normal = count_normal + 1

                        hours = hours + ((f_end1 - f_beg1) + (f_end2 - f_beg2))

                        r_hours_dif = (round(hours,1)) - int(hours)

                        r_hours = round(r_hours_dif,1)

                        if(r_hours == 0.3):
                            hours = hours + 0.7

                        if(r_hours == 0.7):
                            hours = hours + 1.3

                        int_hours = int(hours)

                    else:

                        beg,end = time.split('--')

                        f_beg = float(beg)
                        f_end = float(end)


                        if((f_end - f_beg) <= 1):

                            count_single = count_single + 1

                            hours = hours + 1
                            int_hours = int(hours)

                        else:

                            count_normal = count_normal + 1

                            hours = hours + (f_end - f_beg)

                            r_hours_dif = (round(hours,1)) - int(hours)

                            r_hours = round(r_hours_dif,1)

                            if(r_hours == 0.3):
                                hours = hours + 0.7

                            if(r_hours == 0.7):
                                hours = hours + 1.3
                            
                            int_hours = int(hours)


            str_Counter = str(Counter)

            Cell_J = (Col_J + str_Counter)
            Cell_K = (Col_K + str_Counter)
            Cell_L = (Col_L + str_Counter)
            Cell_M = (Col_M + str_Counter)

            sheet[Cell_J] = count_single
            sheet[Cell_K] = count_normal
            sheet[Cell_L] = int_hours
            sheet[Cell_M] = sunday_hours

            Counter = Counter + 1
        
        except ValueError:
            print(f'Problem at cell: {cellObj.coordinate}')
            return

    Excel.save('FELLMER.xlsx')

    print("ΤΕΛΟΣ")


if __name__ == "__main__":
    main()