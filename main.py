# tool lọc điểm thi vào 10 chuyên dùng định dạng như file score.xlsx
# Made by: @anonymous - from 12IT

import xlwings as xw
from xlwings.constants import LineStyle
import os

# file điểm chung, sheet, khoảng
FILE = "score.xlsx"
SHEET = "Sheet1"
FR_TO = "A4:N668"

# file điểm đã lọc
RESULT = "result.xlsx"

# định dạng cột môn
txt_tin_ = "Tin Học (Chuyên)"
txt_toan_ = "Toán (Chuyên)"
txt_anh_ = "Tiếng Anh (Chuyên)"
txt_van_ = "Ngữ Văn (Chuyên)"
txt_su_ = "Lịch Sử (Chuyên)"
txt_dia_ = "Địa Lý (Chuyên)"
txt_sinh_ = "Sinh Học (Chuyên)"
txt_phap_ = "Tiếng Pháp (Chuyên)"
txt_li_ = "Vật Lý (Chuyên)"
txt_hoa_ = "Hóa (Chuyên)"

# số học sinh 1 lớp
# năm 2023-2024
hs_ = 35

# điểm tối thiểu thay đổi theo năm
# năm 2023-2024
toan_ = 2
anh_ = 2
van_ = 2
chuyen_ = 4

# điểm tối thiểu xét chuyên anh vào chuyên pháp
# năm 2023-2024
sum_p = 34
toan_p = 4
anh_p = 8
van_p = 4
anh_cp = 10

# cột của điểm các môn tỉnh từ 0
toan_s = 5
van_s = 6
anh_s = 7
chuyen_s1 = 9
chuyen_s2 = 11
mon_s1 = 8
mon_s2 = 10

#cột của điểm tổng nv1 và nv2
sum1_s = 12
sum2_s = 13

headers = ["SBD", "Họ và tên", "giới tính", "Học sinh trường", "toán", "văn", "anh", "nguyện vọng 1", "điểm", "nguyện vọng 2", "điểm", "tổng nv1", "tổng nv2"]

def num(var):
    return isinstance(var, (int, float))

def toi_thieu_p(toan, anh, van, chuyen, sum):
    if num(toan) == False or num(anh) == False or num(van) == False or num(chuyen) == False or num(sum) == False:
        return False
    if(toan >= toan_p and anh >= anh_p and van >= van_p and chuyen >= chuyen_ and sum >= sum_p):
        return True
    return False

def toi_thieu(toan, anh, van, chuyen):
    if num(toan) == False or num(anh) == False or num(van) == False or num(chuyen) == False:
        return False
    if(toan >= toan_ and anh >= anh_ and van >= van_ and chuyen >= chuyen_):
        return True
    return False

def get(type, data, mon):
    if type == 1:
        c = chuyen_s1 
        m = mon_s1
    if type == 2:
        c = chuyen_s2
        m = mon_s2
        
    ok = []
    for i in range(len(data)):
      if data[i][m] == mon and toi_thieu(data[i][toan_s], data[i][anh_s], data[i][van_s], data[i][c]):
        ok.append(data[i])
    return ok

def xep(nv1, nv2, nv):
    result = []
    for i in range(len(nv1)):
        if len(result) < hs_:
          result.append(nv1[i])
        else:
            break
    for i in range(len(nv2)):
        if len(result) < hs_ and nv2[i] not in nv:
          result.append(nv2[i])
        else:
            break
    return result
         
def solve(data):
    tin_nv1 = sorted(get(1, data, txt_tin_), key=lambda x: x[sum1_s], reverse=True)
    toan_nv1 = sorted(get(1, data, txt_toan_), key=lambda x: x[sum1_s], reverse=True)
    anh_nv1 = sorted(get(1, data, txt_anh_), key=lambda x: x[sum1_s], reverse=True)
    van_nv1 = sorted(get(1, data, txt_van_), key=lambda x: x[sum1_s], reverse=True)
    su_nv1 = sorted(get(1, data, txt_su_), key=lambda x: x[sum1_s], reverse=True)
    dia_nv1 = sorted(get(1, data, txt_dia_), key=lambda x: x[sum1_s], reverse=True)
    sinh_nv1 = sorted(get(1, data, txt_sinh_), key=lambda x: x[sum1_s], reverse=True)
    phap_nv1 = sorted(get(1, data, txt_phap_), key=lambda x: x[sum1_s], reverse=True)
    li_nv1 = sorted(get(1, data, txt_li_), key=lambda x: x[sum1_s], reverse=True)
    hoa_nv1 = sorted(get(1, data, txt_hoa_), key=lambda x: x[sum1_s], reverse=True)
    
    nv1 = tin_nv1 + toan_nv1 + anh_nv1 + van_nv1 + su_nv1 + dia_nv1 + sinh_nv1 + phap_nv1 + li_nv1 + hoa_nv1
    
    tin_nv2 = sorted(get(2, data, txt_tin_), key=lambda x: x[sum2_s], reverse=True)
    toan_nv2 = sorted(get(2, data, txt_toan_), key=lambda x: x[sum2_s], reverse=True)
    anh_nv2 = sorted(get(2, data, txt_anh_), key=lambda x: x[sum2_s], reverse=True)
    van_nv2 = sorted(get(2, data, txt_van_), key=lambda x: x[sum2_s], reverse=True)
    su_nv2 = sorted(get(2, data, txt_su_), key=lambda x: x[sum2_s], reverse=True)
    dia_nv2 = sorted(get(2, data, txt_dia_), key=lambda x: x[sum2_s], reverse=True)
    sinh_nv2 = sorted(get(2, data, txt_sinh_), key=lambda x: x[sum2_s], reverse=True)
    phap_nv2 = sorted(get(2, data, txt_phap_), key=lambda x: x[sum2_s], reverse=True)
    li_nv2 = sorted(get(2, data, txt_li_), key=lambda x: x[sum2_s], reverse=True)
    hoa_nv2 = sorted(get(2, data, txt_hoa_), key=lambda x: x[sum2_s], reverse=True)
    
    tin = xep(tin_nv1, tin_nv2, nv1)
    toan = xep(toan_nv1, toan_nv2, nv1)
    anh = xep(anh_nv1, anh_nv2, nv1)
    van = xep(van_nv1, van_nv2, nv1)
    su = xep(su_nv1, su_nv2, nv1)
    dia = xep(dia_nv1, dia_nv2, nv1)
    sinh = xep(sinh_nv1, sinh_nv2, nv1)	
    phap = xep(phap_nv1, phap_nv2, nv1)
    li = xep(li_nv1, li_nv2, nv1)
    hoa = xep(hoa_nv1, hoa_nv2, nv1)
    return tin, toan, anh, van, su, dia, sinh, phap, li, hoa

def solve_p(phap, xet_anh):
    res = len(phap)
    anh = sorted(get(1, data, txt_anh_), key=lambda x: x[sum1_s], reverse=True)
    predict = []
    for i in range(len(anh)):
        if res < hs_ and toi_thieu_p(anh[i][toan_s], anh[i][anh_s], anh[i][van_s], anh[i][chuyen_s1], anh[i][sum1_s]) and anh[i] not in xet_anh:
            predict.append(anh[i])
            res =+ 1
    return predict

def write(file, name_s, data):
    if os.path.exists(file) == False:
      wb = xw.Book()
      wb.save(file)
      wb.close()
      
    wb = xw.Book(file)
    if name_s not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(name=name_s)
        
    sheet = wb.sheets[name_s]
    for row_index, row in enumerate(data):
      for col_index, value in enumerate(row):
        cell = sheet.cells(row_index + 1, col_index + 1)
        cell.value = value
    wb.save()

def head(file, headers):
    wb = xw.Book(file)
    for sheet in wb.sheets:
        sheet.api.Rows("1:1").Insert()
        for i, header in enumerate(headers):
            sheet.cells(1, i + 1).value = header
    wb.save()
    
def fit(file):
    wb = xw.Book(file)
    for sheet in wb.sheets:
        sheet.autofit('c')
        sheet.autofit('r')
        used_range = sheet.used_range
        used_range.api.Borders.LineStyle = LineStyle.xlContinuous
    wb.save()
    
def rm(arr):
    return [r[1:] for r in arr]
        
wb = xw.Book(FILE)
sheet = wb.sheets[SHEET]
data = sheet.range(FR_TO).value
wb.close()

tin, toan, anh, van, su, dia, sinh, phap, li, hoa = solve(data)
phap_predict = solve_p(phap, anh)

tin = rm(tin)
toan = rm(toan)
anh = rm(anh)
van = rm(van)
su = rm(su)
dia = rm(dia)
sinh = rm(sinh)
phap = rm(phap)
li = rm(li)
hoa = rm(hoa)
phap_predict = rm(phap_predict)

write(RESULT, 'Toan', toan)
write(RESULT, 'Tin', tin)
write(RESULT, 'Van', van)
write(RESULT, 'Anh', anh)
write(RESULT, 'Su', su)
write(RESULT, 'Dia', dia)
write(RESULT, 'sinh', sinh)
write(RESULT, 'phap', phap)
write(RESULT, 'phap_anh', phap_predict)
write(RESULT, 'hoa', hoa)
write(RESULT, 'ly', li)

head(RESULT, headers)
fit(RESULT)