import xlsxwriter
works=xlsxwriter.Workbook('qq.xlsx')
worksname=works.add_worksheet("qq信息表")
chard=works.add_chart({'type':'column'})

q=[1,2,3,4,5,6,7]
q1="asdfghj"
for i,j in enumerate(q):
    num="A%d"%(i+1)
    worksname.write(num,j)
for i,j in enumerate(q1):
    num='B%d'%(i+1)
    worksname.write_string(num,j)
chard.add_series(
        {
            "categories":"=qq信息表!$B$1:$B$7",
            "values":"=qq信息表!$A$1:$A$7",
            "line":{"color":"red"}
        }
    )
worksname.insert_chart("A8",chart=chard)
works.close()
