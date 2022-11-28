import xlsxwriter
from tkinter import *
from tkinter import filedialog
def openFile():
    input_loc = filedialog.askopenfilename(filetypes=(('text files','txt'), ('All Files', '.*')))
    return input_loc

window = Tk()
button = Button(text="Please wait...", command=openFile)
button.pack()
#window.mainloop()

#input_loc = 'C:/Project work/Macro Tool/Mostafa vi/CBEA9_MRRFIL00-00000003191.txt'
input_loc = openFile()
output_path = '/'.join(input_loc.split('/')[0:-1])
#input_loc = path
f = open(input_loc, 'r')
lines = f.readlines()
count = 0
for line in lines:
    count += 1
print('Total Lines', count)

def write_rxlev(workbook):
    worksheet = workbook.add_worksheet('RXLEV')

    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Cell')
    worksheet.write(0,2,'CHGR')
    worksheet.write(0,3,'RXLEV')
    worksheet.write(0,4,'DL/UL')
    worksheet.write(0,5,'Vector')
    worksheet.write(0,6,'Samples')


    for key, value_list in rxlev_data.items():
        #print(f'Values for the Item {key} are:')
        start = 1
        if key == 'Date':
            for item in value_list:
                #print(item)
                worksheet.write(start,0,item)
                start+=1

        if key == 'Cell':
            for item in value_list:
                #print(item)
                worksheet.write(start,1,item)
                start+=1

        if key == 'CHGR':
            for item in value_list:
                #print(item)
                worksheet.write(start,2,int(item))
                start+=1
                
        if key == 'RXLEV':
            for item in value_list:
                #print(item)
                worksheet.write(start,3,item)
                start+=1

        if key == 'DL/UL':
            for item in value_list:
                #print(item)
                worksheet.write(start,4,item)
                start+=1

        if key == 'Vector':
            for item in value_list:
                #print(item)
                worksheet.write(start,5,int(item))
                start+=1

        if key == 'Samples':
            for item in value_list:
                #print(item)
                worksheet.write(start,6,int(item))
                start+=1

def write_rxqual(workbook):
    worksheet = workbook.add_worksheet('RXQUAL')

    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Cell')
    worksheet.write(0,2,'CHGR')
    worksheet.write(0,3,'RXQUAL')
    worksheet.write(0,4,'DL/UL')
    worksheet.write(0,5,'Vector')
    worksheet.write(0,6,'Samples')


    for key, value_list in rxqual_data.items():
        #print(f'Values for the Item {key} are:')
        start = 1
        if key == 'Date':
            for item in value_list:
                #print(item)
                worksheet.write(start,0,item)
                start+=1

        if key == 'Cell':
            for item in value_list:
                #print(item)
                worksheet.write(start,1,item)
                start+=1

        if key == 'CHGR':
            for item in value_list:
                #print(item)
                worksheet.write(start,2,int(item))
                start+=1
                
        if key == 'RXQUAL':
            for item in value_list:
                #print(item)
                worksheet.write(start,3,item)
                start+=1

        if key == 'DL/UL':
            for item in value_list:
                #print(item)
                worksheet.write(start,4,item)
                start+=1

        if key == 'Vector':
            for item in value_list:
                #print(item)
                worksheet.write(start,5,int(item))
                start+=1

        if key == 'Samples':
            for item in value_list:
                #print(item)
                worksheet.write(start,6,int(item))
                start+=1

def write_ta(workbook):
    worksheet = workbook.add_worksheet('TAVAL')

    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Cell')
    worksheet.write(0,2,'CHGR')
    worksheet.write(0,3,'TAVAL')
    #worksheet.write(0,4,'DL/UL')
    worksheet.write(0,4,'Vector')
    worksheet.write(0,5,'Samples')


    for key, value_list in ta_data.items():
        #print(f'Values for the Item {key} are:')
        start = 1
        if key == 'Date':
            for item in value_list:
                #print(item)
                worksheet.write(start,0,item)
                start+=1

        if key == 'Cell':
            for item in value_list:
                #print(item)
                worksheet.write(start,1,item)
                start+=1

        if key == 'CHGR':
            for item in value_list:
                #print(item)
                worksheet.write(start,2,int(item))
                start+=1
                
        if key == 'TAVAL':
            for item in value_list:
                #print(item)
                worksheet.write(start,3,item)
                start+=1

        if key == 'Vector':
            for item in value_list:
                #print(item)
                worksheet.write(start,4,int(item))
                start+=1

        if key == 'Samples':
            for item in value_list:
                #print(item)
                worksheet.write(start,5,int(item))
                start+=1

def write_msbspwr(workbook):
    worksheet = workbook.add_worksheet('MSBSPWR')

    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Cell')
    worksheet.write(0,2,'CHGR')
    worksheet.write(0,3,'MSBSPWR')
    worksheet.write(0,4,'MS/BS')
    worksheet.write(0,5,'Vector')
    worksheet.write(0,6,'Samples')


    for key, value_list in msbspwr_data.items():
        #print(f'Values for the Item {key} are:')
        start = 1
        if key == 'Date':
            for item in value_list:
                #print(item)
                worksheet.write(start,0,item)
                start+=1

        if key == 'Cell':
            for item in value_list:
                #print(item)
                worksheet.write(start,1,item)
                start+=1

        if key == 'CHGR':
            for item in value_list:
                #print(item)
                worksheet.write(start,2,int(item))
                start+=1
                
        if key == 'MSBSPWR':
            for item in value_list:
                #print(item)
                worksheet.write(start,3,item)
                start+=1

        if key == 'MS/BS':
            for item in value_list:
                #print(item)
                worksheet.write(start,4,item)
                start+=1

        if key == 'Vector':
            for item in value_list:
                #print(item)
                worksheet.write(start,5,int(item))
                start+=1

        if key == 'Samples':
            for item in value_list:
                #print(item)
                worksheet.write(start,6,int(item))
                start+=1

def write_ploss(workbook):
    worksheet = workbook.add_worksheet('PLOSS')

    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Cell')
    worksheet.write(0,2,'CHGR')
    worksheet.write(0,3,'PLOSS')
    worksheet.write(0,4,'DL/UL')
    worksheet.write(0,5,'Vector')
    worksheet.write(0,6,'Samples')


    for key, value_list in ploss_data.items():
        #print(f'Values for the Item {key} are:')
        start = 1
        if key == 'Date':
            for item in value_list:
                #print(item)
                worksheet.write(start,0,item)
                start+=1

        if key == 'Cell':
            for item in value_list:
                #print(item)
                worksheet.write(start,1,item)
                start+=1

        if key == 'CHGR':
            for item in value_list:
                #print(item)
                worksheet.write(start,2,int(item))
                start+=1
                
        if key == 'PLOSS':
            for item in value_list:
                #print(item)
                worksheet.write(start,3,item)
                start+=1

        if key == 'DL/UL':
            for item in value_list:
                #print(item)
                worksheet.write(start,4,item)
                start+=1

        if key == 'Vector':
            for item in value_list:
                #print(item)
                worksheet.write(start,5,int(item))
                start+=1

        if key == 'Samples':
            for item in value_list:
                #print(item)
                worksheet.write(start,6,int(item))
                start+=1

def write_pldiff(workbook):
    worksheet = workbook.add_worksheet('PLDIFF')

    worksheet.write(0,0,'Date')
    worksheet.write(0,1,'Cell')
    worksheet.write(0,2,'CHGR')
    worksheet.write(0,3,'PLDIFF')
    #worksheet.write(0,4,'DL/UL')
    worksheet.write(0,4,'Vector')
    worksheet.write(0,5,'Samples')


    for key, value_list in pldiff_data.items():
        #print(f'Values for the Item {key} are:')
        start = 1
        if key == 'Date':
            for item in value_list:
                #print(item)
                worksheet.write(start,0,item)
                start+=1

        if key == 'Cell':
            for item in value_list:
                #print(item)
                worksheet.write(start,1,item)
                start+=1

        if key == 'CHGR':
            for item in value_list:
                #print(item)
                worksheet.write(start,2,int(item))
                start+=1
                
        if key == 'PLDIFF':
            for item in value_list:
                #print(item)
                worksheet.write(start,3,item)
                start+=1

        if key == 'Vector':
            for item in value_list:
                #print(item)
                worksheet.write(start,4,int(item))
                start+=1

        if key == 'Samples':
            for item in value_list:
                #print(item)
                worksheet.write(start,5,int(item))
                start+=1

def wrightInToExcel():
    workbook = xlsxwriter.Workbook(output_path+'/'+'MRR.xlsx')
    write_rxlev(workbook)
    write_rxqual(workbook)
    write_ta(workbook)
    write_msbspwr(workbook)
    write_ploss(workbook)
    write_pldiff(workbook)
    workbook.close() 
    print("Processing completed .......")

def process_rxlev(i):
    while i <= count-1:
            line = lines[i]
            if 'Cell name:' in line:
                string = line.replace(" ","") # replace the space with no space
                Cell = string.split() # split the sentense for space List
                CellName = Cell[1] #CA9RJ2C
                #print(CellName)

            if 'Channel group:' in line:
                string = line.replace(" ","")
                CG = string.split() # split the sentense for space List
                ChannelGroup = CG[1]
                #print(ChannelGroup)  

            if 'RXLEV' in line:
                string = line.replace(" ","")
                RxLev = string.split() # List
                x = RxLev[0] #RXLEVUL0:
                y = RxLev[1] #Samples
                DL_UL = x[5:7] # take 2 latters from the middle
                if len(x)== 9:
                    Vector = x[7:8]
                elif len(x) ==10:
                    Vector = x[7:9]

                #load data in dictionary
                rxlev_data['Date'].append(Date)
                rxlev_data['Cell'].append(CellName)
                rxlev_data['CHGR'].append(ChannelGroup)
                rxlev_data['RXLEV'].append(x)
                rxlev_data['DL/UL'].append(DL_UL)
                rxlev_data['Vector'].append(Vector)
                rxlev_data['Samples'].append(y)

            if '\n' == line or i>= count:
                break
            i=i+1 

def process_rxqual(i):
    while i <= count-1:
            line = lines[i]
            if 'Cell name:' in line:
                string = line.replace(" ","") # replace the space with no space
                Cell = string.split() # split the sentense for space List
                CellName = Cell[1] #CA9RJ2C
                #print(CellName)

            if 'Channel group:' in line:
                string = line.replace(" ","")
                CG = string.split() # split the sentense for space List
                ChannelGroup = CG[1]
                #print(ChannelGroup)  

            if 'RXQUAL' in line:
                string = line.replace(" ","")
                RxQual = string.split() # List
                x = RxQual[0] #RXLEVUL0:
                y = RxQual[1] #Samples
                DL_UL = x[6:8] # take 2 latters from the middle
                if len(x)== 10:
                    Vector = x[8:9]
                elif len(x) ==11:
                    Vector = x[8:10]

                #load data in dictionary
                rxqual_data['Date'].append(Date)
                rxqual_data['Cell'].append(CellName)
                rxqual_data['CHGR'].append(ChannelGroup)
                rxqual_data['RXQUAL'].append(x)
                rxqual_data['DL/UL'].append(DL_UL)
                rxqual_data['Vector'].append(Vector)
                rxqual_data['Samples'].append(y)

            if '\n' == line or i>= count:
                break
            i=i+1 

def process_ta(i):
    while i <= count-1:
            line = lines[i]
            if 'Cell name:' in line:
                string = line.replace(" ","") # replace the space with no space
                Cell = string.split() # split the sentense for space List
                CellName = Cell[1] #CA9RJ2C
                #print(CellName)

            if 'Channel group:' in line:
                string = line.replace(" ","")
                CG = string.split() # split the sentense for space List
                ChannelGroup = CG[1]
                #print(ChannelGroup)  

            if 'TAVAL' in line:
                string = line.replace(" ","")
                TaVal = string.split() # List
                x = TaVal[0] #TAVAL0:
                y = TaVal[1] #Samples
                #DL_UL = x[5:7] # take 2 latters from the middle
                if len(x)== 7:
                    Vector = x[5:6] #it will take one digit in the position 6
                elif len(x) ==8:
                    Vector = x[5:7] #it will take two digit in the position 6 & 7

                #load data in dictionary
                ta_data['Date'].append(Date)
                ta_data['Cell'].append(CellName)
                ta_data['CHGR'].append(ChannelGroup)
                ta_data['TAVAL'].append(x)
                #rxqual_data['DL/UL'].append(DL_UL)
                ta_data['Vector'].append(Vector)
                ta_data['Samples'].append(y)

            if '\n' == line or i>= count:
                break
            i=i+1 

def process_msbspwr(i):
    while i <= count-1:
            line = lines[i]
            if 'Cell name:' in line:
                string = line.replace(" ","") # replace the space with no space
                Cell = string.split() # split the sentense for space List
                CellName = Cell[1] #CA9RJ2C
                #print(CellName)

            if 'Channel group:' in line:
                string = line.replace(" ","")
                CG = string.split() # split the sentense for space List
                ChannelGroup = CG[1]
                #print(ChannelGroup)  

            if 'MSPOWER' in line or 'BSPOWER' in line:
                string = line.replace(" ","")
                msbspwr = string.split() # List
                x = msbspwr[0] #MSPOWER0:
                y = msbspwr[1] #Samples
                MS_BS = x[0:2] # take 2 latters from the begining
                if len(x)== 9:
                    Vector = x[7:8] #it will take one digit in the position 8
                elif len(x) ==10:
                    Vector = x[7:9] #it will take two digit in the position 8 & 9

                #load data in dictionary
                msbspwr_data['Date'].append(Date)
                msbspwr_data['Cell'].append(CellName)
                msbspwr_data['CHGR'].append(ChannelGroup)
                msbspwr_data['MSBSPWR'].append(x)
                msbspwr_data['MS/BS'].append(MS_BS)
                msbspwr_data['Vector'].append(Vector)
                msbspwr_data['Samples'].append(y)

            if '\n' == line or i>= count:
                break
            i=i+1 

def process_ploss(i):
    while i <= count-1:
            line = lines[i]
            if 'Cell name:' in line:
                string = line.replace(" ","") # replace the space with no space
                Cell = string.split() # split the sentense for space List
                CellName = Cell[1] #CA9RJ2C
                #print(CellName)

            if 'Channel group:' in line:
                string = line.replace(" ","")
                CG = string.split() # split the sentense for space List
                ChannelGroup = CG[1]
                #print(ChannelGroup)  

            if 'PLOSS' in line:
                string = line.replace(" ","")
                ploss = string.split() # List
                x = ploss[0] #PLOSSUL0:
                y = ploss[1] #Samples
                DL_UL = x[5:7] # take 2 latters from the position 6
                if len(x)== 9:
                    Vector = x[7:8] #it will take one digit in the position 8
                elif len(x) ==10:
                    Vector = x[7:9] #it will take two digit in the position 8 & 9

                #load data in dictionary
                ploss_data['Date'].append(Date)
                ploss_data['Cell'].append(CellName)
                ploss_data['CHGR'].append(ChannelGroup)
                ploss_data['PLOSS'].append(x)
                ploss_data['DL/UL'].append(DL_UL)
                ploss_data['Vector'].append(Vector)
                ploss_data['Samples'].append(y)

            if '\n' == line or i>= count:
                break
            i=i+1 

def process_pldiff(i):
    while i <= count-1:
            line = lines[i]
            if 'Cell name:' in line:
                string = line.replace(" ","") # replace the space with no space
                Cell = string.split() # split the sentense for space List
                CellName = Cell[1] #CA9RJ2C
                #print(CellName)

            if 'Channel group:' in line:
                string = line.replace(" ","")
                CG = string.split() # split the sentense for space List
                ChannelGroup = CG[1]
                #print(ChannelGroup)  

            if 'PLDIFF' in line:
                string = line.replace(" ","")
                pldiff = string.split() # List
                x = pldiff[0] #PLDIFF0:
                y = pldiff[1] #Samples
                #DL_UL = x[5:7] # take 2 latters from the position 6
                if len(x)== 8:
                    Vector = x[6:7] #it will take one digit in the position 7
                elif len(x) ==9:
                    Vector = x[6:8] #it will take two digit in the position 7 & 8

                #load data in dictionary
                pldiff_data['Date'].append(Date)
                pldiff_data['Cell'].append(CellName)
                pldiff_data['CHGR'].append(ChannelGroup)
                pldiff_data['PLDIFF'].append(x)
                #msbspwr_data['MS/BS'].append(DL_UL)
                pldiff_data['Vector'].append(Vector)
                pldiff_data['Samples'].append(y)

            if '\n' == line or i>= count:
                break
            i=i+1 


        

rxlev_data = {'Date':[], 'Cell':[], 'CHGR':[], 'RXLEV':[], 'DL/UL':[], 'Vector':[], 'Samples':[]} # empty dictionary
rxqual_data = {'Date':[], 'Cell':[], 'CHGR':[], 'RXQUAL':[], 'DL/UL':[], 'Vector':[], 'Samples':[]} # empty dictionary
ta_data = {'Date':[], 'Cell':[], 'CHGR':[], 'TAVAL':[], 'Vector':[], 'Samples':[]} # empty dictionary
msbspwr_data = {'Date':[], 'Cell':[], 'CHGR':[], 'MSBSPWR':[], 'MS/BS':[], 'Vector':[], 'Samples':[]} # empty dictionary
ploss_data = {'Date':[], 'Cell':[], 'CHGR':[], 'PLOSS':[], 'DL/UL':[], 'Vector':[], 'Samples':[]} # empty dictionary
pldiff_data = {'Date':[], 'Cell':[], 'CHGR':[], 'PLDIFF':[], 'Vector':[], 'Samples':[]} # empty dictionary

i=0
while i <= count-1:
    line = lines[i]
    if 'Start date:' in line:
        string = line.replace(" ","") # replace the space with no space
        dd = string.split() # split the sentense for space List
        Date = dd[1] 
        print(Date)
    if 'UPLINK AND DOWNLINK SIGNAL STRENGTH CELL DATA RECORD' in line:
        process_rxlev(i) 

    if 'UPLINK AND DOWNLINK SIGNAL QUALITY CELL DATA RECORD' in line:
        process_rxqual(i)

    if 'ACTUAL TIMING ADVANCE CELL DATA RECORD' in line:
        process_ta(i)

    if 'BTS AND MS TRANSMIT POWER LEVEL CELL DATA RECORD' in line:
        process_msbspwr(i)

    if 'UPLINK AND DOWNLINK PATH LOSS CELL DATA RECORD' in line:
        process_ploss(i)

    if 'PATH LOSS DIFFERENCE CELL DATA RECORD' in line:
        process_pldiff(i)

    i=i+1

wrightInToExcel()