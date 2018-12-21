# python_xlsx
# read xlsx excel file and collect data

# xlsx.py

# import openpyxl module
import openpyxl

# import os module
import os

# change the working directory
while True:
		cwd = os.getcwd()
		wd_ch = input("Working on "+cwd+", do you wish to change? Y/N:")
    if wd_ch == 'Y' or wd_ch == 'y' or wd_ch == 'Yes' or wd_ch == 'yes' :
      while True:
        cwd_n = input("Please input the folder directory:")
        try:
          os.chdir(cwd_n) 
          break
        except FileNotFoundError or OSError:
          print("Error! Please input a valid directory, e.g.'f:/myfolder'")
          continue
      print(f"Successfully changed to {cwd_n}")
      break
    elif wd_ch == 'N' or wd_ch == 'n' or wd_ch == 'No' or wd_ch == 'no' :
      cwd_n = cwd
      break
    else:
        print("Error! Please answer 'Yes' or 'No'")
        continue
# show all the files in working directory
files = []
print('\n'+"This folder contains following files:"+'\n')
for filename in os.listdir(cwd_n):
  print(filename)
  files.append(filename)
# ask the user's purpose
while True:
  tar = input('\n'+"What do you want to do? "+'\n'"1=read and modify"+'\n'"2=read and copy to a new"+'\n'"\
  3=read and append to an exist"+'\n'"4=read and compare"+'\n'"5=exit"+'\n'":")
  if tar == "1":
    import xlsx_mo
    break
  elif tar == "2":
    import xlsx_ne
    break
  elif tar == "3":
    import xlsx_ap
    break
  elif tar == "4":
    import xlsx_co
    break
  elif tar == "5":
    print("Exit now"+'\n')
    break
  else:
    print("Error! Please input the No. of operations")
    continue
