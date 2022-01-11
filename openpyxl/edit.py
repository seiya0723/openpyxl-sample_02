import openpyxl as px
import datetime

#編集対象のファイルを読み込み
wb  = px.load_workbook('salary.xlsx')

#アクティブシートを選択(新規作成時に最初からあるシート)
#ws  = wb.active
ws  = wb.worksheets[0]

#ws["C2"].value  = "テスト"


start       = ws["A7"].value
rest_start  = ws["B7"].value
rest_end    = ws["C7"].value
end         = ws["D7"].value

print(type(start      ))
print(type(rest_start ))
print(type(rest_end   ))
print(type(end        ))

work_time   = (rest_start - start) + (end - rest_end)
print(work_time)


print(type(work_time))

normal_time = datetime.timedelta(hours=8)

if normal_time < work_time:
    print("残業しています")
    print(work_time - normal_time)
    #TODO:ここで残業分の賃金計算




#TODO:特定の時間を含むかの条件式。画像参照







#別名で保存
today   = str(datetime.date.today())

wb.save("salary" + today + ".xlsx")

