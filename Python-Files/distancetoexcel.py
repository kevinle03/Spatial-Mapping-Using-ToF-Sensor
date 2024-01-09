import xlwt
import serial

import serial


style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
    num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

wb = xlwt.Workbook()
ws = wb.add_sheet('Distance Data')

ws.write(0,0,"Step")
ws.write(0,1,"Angle")
for row in range(1,513):
    ws.write(row,0,row)
for angle in range(1,513):
    ws.write(angle,1,(360/512)*angle)


ser = serial.Serial(
    port='COM5',\
    baudrate=115200,\
    parity=serial.PARITY_NONE,\
    stopbits=serial.STOPBITS_ONE,\
    bytesize=serial.EIGHTBITS,\
        timeout=0)

print("connected to: " + ser.portstr)
count=1

print("Press button 2 to begin data acquisition.")
is_break2 = False
while True:
    if is_break2:
        is_break2 = False
        break
    for line in ser.read():
        if chr(line) == '@':
            print("Data acquisition will begin.")
            print("Press butotn 1 to start the scanner")
            is_break2 = True
            break

col = 2
ws.write(0,col,0)
while True:            
    meas = ""
    is_break = False
    while True:
        if is_break:
            is_break = False
            break
        for line in ser.read():
            if line >= 48 and line <= 57:
                print(chr(line) )
                meas += chr(line)
            if chr(line) == '!':
                is_break = True
                break
    ws.write(count,col,meas)
    wb.save('dataset.xls')
    print(str(count) + str(': ') + meas)
    if count == 512:
        count = 0
        val = input("Enter X coordinate (in mm) to continue or 's' to stop\n")
        if val == 's':
            break
        else:
            print("Press button 2 to begin data acquisition.")
            is_break2 = False
            while True:
                if is_break2:
                    is_break2 = False
                    break
                for line in ser.read():
                    if chr(line) == '@':
                        print("Data acquisition will begin.")
                        print("Press butotn 1 to start the scanner")
                        is_break2 = True
                        break
            col += 1
            ws.write(0,col,val)
    count += 1

wb.save('dataset.xls')
