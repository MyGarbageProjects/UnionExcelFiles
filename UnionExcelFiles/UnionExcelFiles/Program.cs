using System;
using System.Collections.Generic;
using System.Drawing.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using  OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.Style;

namespace UnionExcelFiles
{
    class Program
    {
        static List<Excel> ExcelList = new List<Excel>();
        private static string[,] unionTables;
        private static readonly string pathFolder = $"{AppDomain.CurrentDomain.BaseDirectory}Контакты с Caspian Agency";
        static void Main(string[] args)
        {
            DirectoryInfo di = new DirectoryInfo(pathFolder);
            List<string[,]> _tmp = new List<string[,]>();
            //foreach (FileInfo file in di.GetFiles())
            //{
            //    Console.WriteLine(file.Name);
            //    ExcelList.Add(new Excel(file));
            //    _tmp.Add(ExcelList.Last().ExcelTable);
            //}
            for (int i = 0; i < di.GetFiles().Length; i++)
            {
                FileInfo file = new FileInfo($"{pathFolder}\\amocrm__contacts ({i}).xlsx");
                Console.WriteLine(file.Name);
                ExcelList.Add(new Excel(file));
                _tmp.Add(ExcelList.Last().ExcelTable);
            }

            Console.WriteLine("Порядо объединения соблюден?Если нет, просто закройте программу, если да, нажмите на любую клавишу");
            Console.ReadKey();
            int cursorRowIndex = 0;
            int offsetRow = 0;


            using (OfficeOpenXml.ExcelPackage eP = new OfficeOpenXml.ExcelPackage())
            {
                eP.Workbook.Worksheets.Add("Worksheet1");
                using (OfficeOpenXml.ExcelWorksheet eWs = eP.Workbook.Worksheets[1])
                {

                    for (int i = 0; i < ExcelList.Count; i++)
                    {
                        for (int row = 1; row < ExcelList[i].ExcelTable.GetLength(0); row++)
                        {
                            for (int column = 1; column < ExcelList[i].ExcelTable.GetLength(1); column++)
                            {
                                eWs.Cells[row + offsetRow, column].Value = ExcelList[i].ExcelTable[row - 1, column - 1];
                            }
                        }

                        offsetRow += ExcelList[i].ExcelTable.GetLength(0);
                    }

                    string nameUnionExcel = "UnionExcel.xlsx";
                    //
                    string filePath = AppDomain.CurrentDomain.BaseDirectory + nameUnionExcel;
                    //
                    eP.SaveAs(new FileInfo(filePath));                           //eP.SaveAs(new FileInfo($"{reportDir.FullName}\\{nameReport}"));
                    //WwasLogMsg(null, $"Лог: Отчет сохранен в папке {filePath}"); //WwasLogMsg(null, $"Лог: Отчет сохранен в папке с программой \"{nameReport}\"");
                    Console.WriteLine($"Фаил сохранент: {nameUnionExcel}");
                }
                //OfficeOpenXml.ExcelPackage eP = new OfficeOpenXml.ExcelPackage(new FileInfo(oFD.FileName))
            }

            

            Console.WriteLine();
        }

        public static void SaveExcelFile(string Path)
        {
            //var reportDir = Directory.CreateDirectory($"{AppDomain.CurrentDomain.BaseDirectory}");
            //using (OfficeOpenXml.ExcelPackage eP = new OfficeOpenXml.ExcelPackage())
            //{
            //    eP.Workbook.Worksheets.Add("Worksheet1");
            //    using (OfficeOpenXml.ExcelWorksheet eWs = eP.Workbook.Worksheets[1])
            //    {
            //        eWs.Cells[1, 1].Value = "Валидные номера";
            //        eWs.Cells[1, 2].Value = "Не валидные номера";
            //        for (int i = 0; i < SuccessfulPhone.Count; i++)
            //        {
            //            eWs.Cells[i + 2, 1].Value = SuccessfulPhone[i];
            //        }
            //        for (int i = 0; i < UnsuccessfulPhone.Count; i++)
            //        {
            //            eWs.Cells[i + 2, 2].Value = UnsuccessfulPhone[i];
            //        }

            //        string nameReport = $"{nameDBfile}.report({DateTime.Now.ToString("HH.mm - dd.MM.yyyy")}).xlsx";
            //        //
            //        string filePath = toSaveReport(nameReport);
            //        //
            //        eP.SaveAs(new FileInfo(filePath));                           //eP.SaveAs(new FileInfo($"{reportDir.FullName}\\{nameReport}"));
            //        WwasLogMsg(null, $"Лог: Отчет сохранен в папке {filePath}"); //WwasLogMsg(null, $"Лог: Отчет сохранен в папке с программой \"{nameReport}\"");
            //    }
            //    //OfficeOpenXml.ExcelPackage eP = new OfficeOpenXml.ExcelPackage(new FileInfo(oFD.FileName))
            //}
        }


        public static void UnionExcelFiles()
        {

        }
        
    }

    static class ExtentionLINQ
    {
        //public 
    }

    class Excel
    {
        private string[,] excelTable;
        public string[,] ExcelTable {get => excelTable;}

        public Excel(FileInfo ExcelFile, string password = null )
        {
            if (password is null)
            {
                using (ExcelPackage ep = new ExcelPackage(ExcelFile))
                {
                    ExcelWorksheet ew = ep.Workbook.Worksheets.First();
                    int endRow = ew.Dimension.End.Row;
                    int endColumn = ew.Dimension.End.Column;
                    excelTable = new string[endRow, endColumn];
                    for (int i = 1; i <= endRow; i++)
                    {
                        for (int j = 1; j <= endColumn; j++)
                        {
                            excelTable[i - 1, j - 1] = ew.Cells[i, j].Text;
                        }
                    }
                }
            }
        }
    }
}
