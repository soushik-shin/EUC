#--------------------------------------------------------------------------
# 샤프론 엑셀자료 data cleansing
#--------------------------------------------------------------------------
import os
import openpyxl
import sys

print('Start of job')


wb = openpyxl.load_workbook(r"C:\Users\soush\Downloads\무선통신고등학교.xlsx")
sh = wb.active

area = sh["K1"].value
section = sh["K2"].value
school = sh["K3"].value
fileno = sh["K4"].value

if school == None:
    print('기본 설정값을 확인해 주세요')
    exit()

filename = fileno + '_' + school + '_' + 'A.csv'
outfile = open(filename, 'w')

header = ', '.join(['지역', '지구', '학교명', '직책', '이름', '휴대전화', '신규', '기존', '학생명', '년반', '파일No', '그룹명', '샤프론수', '프론티어수', '샤프론단비', '프론티어단비', '중복']) + '\n'
outfile.write(header)

title = ""
namelist = []
for row in range(6, sh.max_row + 1):
    if row in (9, 10, 11, 12, 13):
        continue

    if row in range(6,9):
        phone = sh["E" + str(row)].value
        newMember = None
        student = ' '
        grade = ' '
    else:
        phone = sh["D" + str(row)].value
        newMember = sh["F" + str(row)].value 
        student = sh["I" + str(row)].value
        grade = sh["J" + str(row)].value

    title_temp = sh["B" + str(row)].value
    if title != title_temp:
        if title_temp == '"':
            pass
        else:
            title = title_temp

    name = sh["C" + str(row)].value
    try:
        if len(name) > 5:
            continue
    except:
        continue
    
    # dup check
    dupCheck = ""
    if name in namelist:
        dupCheck = "중복"
    else:
        namelist.append(name)

    reMember = sh["G" + str(row)].value 
    memberType = ""
    saffron_No = ''
    saffron_fee = ''
    if newMember != None:
        saffron_No = '1'
        saffron_fee = '20000'
    frontier_No = ''
    frontier_fee = ''
    if student != ' ':
        frontier_No = '1'
        frontier_fee = '12500'

    outfile.write(area + ', ' +  section + ', ' + school + ', ' + title + ', ' + name + ', ' + phone + ', ' )
    if newMember == None:
        outfile.write('' + ', ')
    else:
        outfile.write(str(newMember) + ', ')
    if reMember == None:
        outfile.write('' + ', ')
    else:
        outfile.write(str(reMember) + ', ')
    if title in ('교장', '교감', '지도교사'):
        memberType = title
    else:
        memberType = '샤프론'
    outfile.write(grade + ', ' + fileno + ', ' + student + ', ' + memberType + ', ')
    outfile.write(saffron_No + ', ' + frontier_No + ', ' + saffron_fee + ', ' + frontier_fee + ', ' + dupCheck )
    outfile.write('\n')
outfile.close()

print('Done of Sorter')    