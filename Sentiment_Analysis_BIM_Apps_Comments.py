# This code is from ZHOU Shenghua, Department of Civil Engineering, The University of Hong Kong
from textblob import TextBlob
from  textblob.sentiments  import  NaiveBayesAnalyzer
import xlrd
import xlwt
# import nltk
# nltk.download('movie_reviews')
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('My Worksheet')
#style = xlwt.XFStyle()
# font = xlwt.Font()
# font.name = 'Times New Roman'
# font.bold = True
# font.underline = True
# font.italic = True
# style.font = font
myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
data = xlrd.open_workbook("C:\\Users\\zhou\\Desktop\\new.xlsx")
table = data.sheet_by_name(u'Sheet1')
n = 0
for i in table.col_values(2):
    blob = TextBlob(i)
    # if len(blob)>2:
    #     try:
    #         blob.translate(to='es')
    #     # print(blob.sentiment.polarity)
    #     except:
    #         print(n + 1)
    if len(blob) > 2:
        blob.translate(to='es')
        p=blob.sentiment.polarity
        s=blob.sentiment.subjectivity
    # print(blob.sentiment.subjectivity, blob.sentiment.subjectivity)

    else:
        s=100
        p=100
    # worksheet.write(n, 1, blob.sentiment.subjectivity)p=100
    worksheet.write(n, 0, p)
    worksheet.write(n, 1,s)
    n = n + 1

worksheet.write(n, 0, p)
worksheet.write(n, 1,s)


workbook.save('formatting1.xls')
