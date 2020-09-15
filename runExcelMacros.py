import win32com.client 
import glob

xl = win32com.client.Dispatch('Excel.Application')
xl.Workbooks.Open(Filename = "d:\\excel\\Book1.xlsm", ReadOnly=1)

files = glob.glob('d:\\excel\\*.bas')

try:
    for filename in files:
        filename = filename.replace('d:\\excel\\','')
        filename = filename.replace('.bas','')
        res1 = xl.Application.Run(filename+".FizzBuzz")
        res2 = xl.Application.Run(filename+".Bond",0.03, 2000000, 0.04, 10)
        res3 = xl.Application.Run(filename+".vecMatMult")
        print(filename, res1, " ", res2, " ", res3)
finally:
    print("Completed")
    
xl.Application.Quit()
del xl
 
