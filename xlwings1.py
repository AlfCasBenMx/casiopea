import xlwings as xw
wb = xw.Book(r'C:\Users\Alfredo Castaneda\Documents\smartP\TMDB.xlsm')  # on Windows: use raw strings to escape backslashes
sht = wb.sheets['1']
#sht.range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
#list=[sht.range((x,y)) for x in range(20) for y in range(20) if sht.range((x,y)) !=""]
#for i in range(20):
#    print(sht.range((1,i)).value)
rng[:, 3:]
