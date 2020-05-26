import requests 
from bs4 import BeautifulSoup
import re
from datetime import timedelta,date
import xlsxwriter


amirxlsx = xlsxwriter.Workbook('amir2.xlsx',{'strings_to_numbers':True})

amirxlsxbook = amirxlsx.add_worksheet()


def datarange(start_date,end_date):
    for n in range((end_date - start_date).days):
        yield start_date + timedelta(n)

start_date = date(2019,10,10)
end_date = date(2019,11,5)
row = 1

for single_date in datarange(start_date,end_date):
    amir= single_date.strftime("%Y%m%d")
    url = "http://cdn.tsetmc.com/Loader.aspx?ParTree=15131P&i=47302318535715632&d={}".format(amir)
    response = requests.get(url)
    data = response.text
    soup = BeautifulSoup(data,'html.parser')

    hiderow = re.compile("var\s+ClosingPriceData=(.*)")
    hidden = hiderow.findall(soup.text)
    mamad = 2
    for char in hidden :
        for char2 in char : 
            mamad +=1 

    if(mamad > 500):   
        row +=1
        col = 1
        column = 1

        #خرید حقیقی به میلیارد تومن
        formol = "(H{}-I{})*((D{}+E{})/2)/10000000000".format(row+1,row+1,row+1,row+1)
        amirxlsxbook.write_formula(row,9,formol)

        #خرید تجمعی حقیقی    
        amirxlsxbook.write_formula(2,10,'(J3)')
        formol2 = "(J{}+K{})".format(row+1,row)
        amirxlsxbook.write_formula(row,10,formol2)

        #قدرت خریدار به فروشنده
        formol3 = "((H{}*G{})/(I{}*F{}))".format(row+1,row+1,row+1,row+1)
        amirxlsxbook.write_formula(row,11,formol3)


        #Get End Price
        end = re.compile("var\s+ClosingPriceData=(.*)")
        endprice = end.findall(soup.text)
        i=1
        for item in endprice :
            item5 = item.replace('[[','[')
            item6 = item5.replace(']]',']')
            item7 = item6.split('],[')
            item8 = item7[-1]
            item9 = item8.replace("'","")
            item10 = item9.split(',')
            for item11 in item10:
                if col in [1]:
                    column +=1
                    item12 = item11.split(' ')
                    item13 = item12[0]
                    amirxlsxbook.write(row,column,item13)
                      
                if col in [3,4]:
                    column +=1
                    amirxlsxbook.write(row,column,item11)                  
                col +=1

        #Get Haghighi va Hoghoghi
        p = re.compile("var\s+ClientTypeData=(.*)")
        b = p.findall(soup.text) 
        for item in b :
            item2 = item.split(',')
            for item3 in item2 :
                item4 = item3.replace("[","")
                item5 = item4.replace("]","")
                if col in [14,16,18,20] :
                    column +=1
                    amirxlsxbook.write(row,column,item5)
                col +=1
                
    #f.write(d)
    #f.close()
amirxlsxbook.write(1,2,'تاریخ')   
amirxlsxbook.write(1,3,'آخرین قیمت')
amirxlsxbook.write(1,4,'قیمت پایانی')   
amirxlsxbook.write(1,5,'تعداد خرید حقیقی')   
amirxlsxbook.write(1,6,'تعداد فروش حقیقی')   
amirxlsxbook.write(1,7,'حجم خرید حقیقی')   
amirxlsxbook.write(1,8,'حجم فروش حقیقی') 
amirxlsxbook.write(1,9,'خرید حقیقی به میلیارد تومن') 
amirxlsxbook.write(1,10,'خرید تجمعی حقیقی')   
amirxlsxbook.write(1,11,'قدرت خریدار به فروشنده')   

categoties_chart = '=Sheet1!$C3:$C{}'.format(row+1)
# اضافه کردن نمودار ملیه ای
column_chart = amirxlsx.add_chart({'type':'column'})
column_chart_row = '=Sheet1!$J3:$J{}'.format(row+1)
column_chart.add_series({
    'name':'خرید روزانه حقیقی به میلیارد تومان',
    'categories':categoties_chart,
    'values':column_chart_row
})

#اضافه کردن نمودار خطی
line_chart = amirxlsx.add_chart({'type':'line'})
line_chart_row = '=Sheet1!$K3:$K{}'.format(row+1)
line_chart.add_series({
    'name':'خرید تجمعی حقیقی به میلیارد تومان',
    'categories':categoties_chart,
    'values':line_chart_row,
     'y2_axis':True
})

column_chart.combine(line_chart)
amirxlsxbook.insert_chart('L10',column_chart)

amirxlsx.close()
