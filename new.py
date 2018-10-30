elif(rule == 4):

          my_data = inp.split(',')


          for i in range(3,1000):
              for j in my_data:

                  if(input_sheet.cell(row = i, column = col2num(j)).value != None):
                      fin = ""

                      for a,a1 in zip(my_data,5):
                          fin = fin +" "+headers[a1]+": "+str(input_sheet.cell(row = i, column = col2num(a)).value)+";"
                      output_sheet.cell(row = i, column = col2num(out)).value = fin
