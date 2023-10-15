using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelManipulator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter File to be manipulated along with path");
            var path = Console.ReadLine();
            Console.WriteLine("Enter Sheet Name to be read");
            var sheetname = Console.ReadLine();
            int column = 1, row=1;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();
            var workSheet = excel.Workbook.Worksheets.Add(sheetname);
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var firstsheet = package.Workbook.Worksheets[sheetname];
                if (firstsheet != null)
                {
                    Console.WriteLine("Enter the address to start manipulation:");
                    var firstcell = Convert.ToInt32(Console.ReadLine());
                    Console.WriteLine("Enter the address to end manipulation:");
                    var lastcell = Convert.ToInt32(Console.ReadLine());

                    for (int i = firstcell; i <= lastcell; i++)
                    {
                        string name = firstsheet.Cells["C" + i].Text;
                        string address = firstsheet.Cells["E" + i].Text;
                        string phonenumber = firstsheet.Cells["G" + i].Text;
                        if(column<=3)
                        {
                            workSheet.Cells[row, column].Value = name + " \n " + address+" \n Phone:"+ phonenumber;
                            workSheet.Row(row).CustomHeight=true;
                            workSheet.Column(column).Width = 60;
                            workSheet.Column(column).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            workSheet.Column(column).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            workSheet.Column(column).Style.WrapText = true;
                            workSheet.Column(column).Style.Font.Bold = true;
                            workSheet.Column(column).Style.Font.Size =Convert.ToSingle(14);
                            workSheet.Column(row).AutoFit();
                            column++;
                        }
                        else
                        {
                            row++;
                            column = 1;

                            workSheet.Cells[row, column].Value = name + " \n " + address+ " \n Phone:" + phonenumber;
                            workSheet.Row(row).CustomHeight = true;
                            workSheet.Column(column).Width = 60;
                            workSheet.Column(column).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            workSheet.Column(column).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                            workSheet.Column(column).Style.WrapText = true;
                            workSheet.Column(column).Style.Font.Bold = true;
                            workSheet.Column(column).Style.Font.Size = Convert.ToSingle(14);
                            workSheet.Column(row).AutoFit();
                            column++;
                        }
             
                    }

               
                }
            }
            
            var modelTable = workSheet.Cells["A:C"];
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thick;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thick;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thick;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            // file name with .xlsx extension  
            Console.WriteLine("Enter path to save the processed file to:");
            string str_path = Console.ReadLine();
            string p_strPath = str_path+"\\" + sheetname + ".xlsx";

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package 
            excel.Dispose();
            Console.WriteLine("Processing Completed");
            Console.ReadKey();
        }
    }
}
