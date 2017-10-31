# ExcelTableFormatter
This is small app that format table from specific way.
I use this app in my work.
-Just copy some data from Sap application and my app clear it and add other data.
-First you must change path to excel file Table.xlsx
If you set the valiable clsExcel.Visible = false to true
the app will open excel and you can see the all proccess.
-When you copy information from Exsample.xlsx(Sap data) to Table.xlsx and run exe file, the app
will delete unused columns, add some data to table and make borders.
-DataMatrix folder is for storage the barcode images so don't delete it.
-In method CreateDataMatrixCode() you can comment row N203 "File.Delete($"{path}Table\\DataMatrix\\{orderNumber}.bmp");"
that will storage images of data matrix code in foder DataMatrix
