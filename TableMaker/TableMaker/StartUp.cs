namespace TableMaker
{
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.IO;
    using OnBarcode.Barcode;
    using System.Runtime.InteropServices;

    public class StartUp
    {
        static void Main()
        {
            Console.WriteLine("Open Exel...");
            Application clsExcel = new Application();
            clsExcel.Visible = false;
            Workbooks workbooks = clsExcel.Workbooks;
            string path = "C:\\Users\\bg4u0059\\Desktop\\";

            Console.WriteLine("Open Template...");
            Workbook clsWorkbook = workbooks.Open(
                     path + "Table\\Table.xlsx", 2, false, 5, "", "", true,
                     XlPlatform.xlWindows, "",
                     false, true, 0, false, true,
                     XlCorruptLoad.xlNormalLoad);
            Console.WriteLine("Start...");
            Worksheet clsWorksheet = clsWorkbook.Sheets[1];
            try
            {
                DeleteColumns(clsWorksheet);
                AddHeader(clsWorksheet);
                Range defaultRange = clsWorksheet.get_Range("A1", "E200");
                int rowsRange = GetRowsRange(clsWorksheet, defaultRange);
                Range range = clsWorksheet.get_Range("A1", $"E{rowsRange}");
                FormatTable(clsWorksheet, range);
                AddDataInCells(clsWorksheet, range);
                CreateDataMatrixCode(clsWorksheet, range, path);
                range = clsWorksheet.get_Range("A1", $"E{rowsRange-1}");
                AddBorders(clsWorksheet, range);

                Console.WriteLine("Saving...");
                clsWorkbook.Save();
                Marshal.FinalReleaseComObject(clsWorksheet);
                clsWorkbook.Close();
                Marshal.ReleaseComObject(clsWorkbook);
                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
                clsExcel.Application.Quit();
                Marshal.ReleaseComObject(clsExcel);
                clsExcel = null;
            }
            catch (Exception e)
            {
                Marshal.FinalReleaseComObject(clsWorksheet);
                clsWorkbook.Close();
                Marshal.ReleaseComObject(clsWorkbook);
                workbooks.Close();
                Marshal.ReleaseComObject(workbooks);
                clsExcel.Application.Quit();
                Marshal.ReleaseComObject(clsExcel);
                clsExcel = null;
                Console.WriteLine("------------ERROR------------");
                Console.WriteLine(e.Message);
                Console.WriteLine("Close Command Promp manually");
                Console.ReadKey();
            }
            finally
            {
                Console.WriteLine("YOUR TABLE WAS CREATED SUCCESSFULLY");
            }

        }

        private static void DeleteColumns(Worksheet clsWorksheet)
        {
            Console.WriteLine("Delete columns...");
            for (int i = 0; i < 4; i++)
            {
                ((Range)clsWorksheet.Cells[1, 1]).EntireColumn.Delete(null);
            }

            for (int i = 0; i < 4; i++)
            {
                ((Range)clsWorksheet.Cells[2, 2]).EntireColumn.Delete(null);
            }
           ((Range)clsWorksheet.Cells[3, 3]).EntireColumn.Delete(null);
            for (int i = 0; i < 3; i++)
            {
                ((Range)clsWorksheet.Cells[4, 4]).EntireColumn.Delete(null);
            }
        }

        private static void AddHeader(Worksheet clsWorksheet)
        {
            Console.WriteLine("Add Header...");
            Range rangeForSorting = clsWorksheet.UsedRange;
            rangeForSorting.Sort(rangeForSorting.Columns[3], XlSortOrder.xlAscending);
            Range line = (Range)clsWorksheet.Rows[1];
            line.Insert();
            //order number
            clsWorksheet.Cells[1, 1] = "№ Поръчка";
            //serial number of pcb
            clsWorksheet.Cells[1, 2] = "№ Изделие";
            //pcb quantity
            clsWorksheet.Cells[1, 3] = "Бр. Платки";
            //panels quantity
            clsWorksheet.Cells[1, 4] = "Бр. Панели";
            //serial number of the panels
            clsWorksheet.Cells[1, 5] = "№ Панел";
            ((Range)clsWorksheet.Rows[1]).EntireRow.Font.Bold = true;
        }

        private static int GetRowsRange(Worksheet clsWorksheet, Range defaultRange)
        {
            int lastRow = 0;
            for (int row = 1; row < defaultRange.Rows.Count; row++)
            {
                var columnA = ((Range)clsWorksheet.Cells[row, 1]).Value;
                Convert.ToString(columnA);
                var columnB = ((Range)clsWorksheet.Cells[row, 2]).Value;
                Convert.ToString(columnB);
                var columnC = ((Range)clsWorksheet.Cells[row, 3]).Value;
                Convert.ToString(columnC);
                if (columnA == null && columnB == null && columnC == null)
                {
                    lastRow = row;
                    break;
                }
            }
            return lastRow;
        }

        private static void FormatTable(Worksheet clsWorksheet, Range range)
        {
            Console.WriteLine("Format Table....");
            clsWorksheet.Columns.EntireColumn.ColumnWidth = 15;
            clsWorksheet.Rows.EntireRow.RowHeight = 28;
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;

            //delete '.' and ',' from cells with pcb quantity
            for (int row = 2; row < range.Rows.Count; row++)
            {
                dynamic cellValue = ((Range)clsWorksheet.Cells[row, 3]).Value;
                string cellData = Convert.ToString(cellValue);
                if (cellData != null)
                {
                    if (cellData.Contains("."))
                    {
                        string newValue = cellData.Replace(".", "");
                        newValue = newValue.Remove(newValue.Length - 4);
                        clsWorksheet.Cells[row, 3] = newValue;
                    }
                }
            }
            //make first column vertical alignment bottom to have room for data matrix code image
            for (int row = 2; row < range.Rows.Count; row++)
            {
                range[row, 1].VerticalAlignment = XlVAlign.xlVAlignBottom;

            }
        }

        private static void AddDataInCells(Worksheet clsWorksheet, Range range)
        {
            Console.WriteLine("Add Data To Table...");

            //add panels quantity in order
            for (int row = 2; row < range.Rows.Count; row++)
            {
                //column C
                clsWorksheet.Cells[row, 4] = $"=C{row}/VLOOKUP(B{row},Sheet2!A$1:L$700,5,0)";
            }

            // add bord serial number
            for (int row = 2; row < range.Rows.Count; row++)
            {
                //column D
                clsWorksheet.Cells[row, 5] = $"=VLOOKUP(B{row},Sheet2!A$1:L$700,3,0)";
            }
        }

        private static void CreateDataMatrixCode(Worksheet clsWorksheet, Range range, string path)
        {
            Console.WriteLine("Create Data Matrix Code...");
            for (int row = 2; row < range.Rows.Count; row++)
            {
                dynamic cellValue = ((Range)clsWorksheet.Cells[row, 1]).Value;
                string orderNumber = Convert.ToString(cellValue);

                DataMatrix datamatrix = new DataMatrix();
                datamatrix.Data = orderNumber;
                datamatrix.DataMode = DataMatrixDataMode.ASCII;
                datamatrix.ImageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
                datamatrix.drawBarcode($"{path}Table\\DataMatrix\\{orderNumber}.bmp");

                Range oRange = range.Cells[row, 1];
                float Left = (float)((double)oRange.Left + 35);
                float Top = (float)((double)oRange.Top + 3);
                const float ImageSize = 13;
                clsWorksheet.Shapes.AddPicture($"{path}Table\\DataMatrix\\{orderNumber}.bmp",
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, ImageSize, ImageSize);
                File.Delete($"{path}Table\\DataMatrix\\{orderNumber}.bmp");
            }

        }

        private static void AddBorders(Worksheet clsWorksheet, Range range)
        {
            Borders borders = range.Borders;
            borders.LineStyle = XlLineStyle.xlContinuous;
            borders.Weight = 2d;
        }
    }
}
