using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelComparator
{
    class comparator
    {
        static void Main(string[] args)
        {
            String targetPath = "C:\\Users\\rashmi\\Desktop\\ExcelComparatorFiles\\ResultFiles\\";
            try
            {
                Excel.Application xlApp;
                Excel.Workbook pathFileWbook, resultFilewbook, expFilewbook;
                Excel.Worksheet pathFileSheet, resultFileWsheet, expFilewsheet;
                Excel.Range pathFilerange, resultFileRange, expFileRange;

                //open path file
                xlApp = new Excel.Application();
                pathFileWbook = xlApp.Workbooks.Open("C:\\Users\\rashmi\\Desktop\\ExcelComparatorFiles\\PathFile.xlsx");
                pathFileSheet = (Excel.Worksheet)pathFileWbook.Worksheets.get_Item(1);

                pathFilerange = pathFileSheet.UsedRange;

                string actualFile, expectedfile, resultFile;


                for (int i = 2; i <= pathFilerange.Rows.Count; i++)          //Starting from row 2 as 1st row is header
                {
                    string execute = pathFileSheet.Cells[i, 1].value.ToString();

                    if (execute.ToUpper() == "Y")
                    {
                        actualFile = pathFileSheet.Cells[i, 2].value.ToString();
                        expectedfile = pathFileSheet.Cells[i, 3].value.ToString();

                        //copy actual file and add that path to result file
                        if (!System.IO.Directory.Exists(targetPath))
                        {
                            System.IO.Directory.CreateDirectory(targetPath);
                        }

                        String[] actFileNameArr = actualFile.Split('\\');

                        System.IO.File.Copy(actualFile, targetPath + actFileNameArr[actFileNameArr.Length - 1], true);
                        resultFile = targetPath + actFileNameArr[actFileNameArr.Length - 1];
                        pathFileSheet.Cells[i, 4].value = resultFile;
                        pathFileSheet.Cells[i, 5].value = "";

                        //opening result file, expected file
                        expFilewbook = xlApp.Workbooks.Open(expectedfile);
                        resultFilewbook = xlApp.Workbooks.Open(resultFile);

                        resultFileWsheet = (Excel.Worksheet)resultFilewbook.Worksheets.get_Item(1);
                        expFilewsheet = (Excel.Worksheet)expFilewbook.Worksheets.get_Item(1);

                        //getting range for both file
                        resultFileRange = resultFileWsheet.UsedRange;
                        expFileRange = expFilewsheet.UsedRange;

                        if (resultFileRange.Rows.Count == expFileRange.Rows.Count && resultFileRange.Columns.Count == expFileRange.Columns.Count)
                        {
                            for (int j = 1; j <= resultFileRange.Rows.Count; j++)
                            {
                                for (int k = 1; k <= resultFileRange.Columns.Count; k++)
                                {
                                    if (resultFileWsheet.Cells[j, k].value != null && expFilewsheet.Cells[j, k].value != null)
                                    {
                                        string actValue = (resultFileWsheet.Cells[j, k].value.ToString());
                                        string expValue = (expFilewsheet.Cells[j, k].value.ToString());

                                        if (actValue.Trim() == expValue.Trim())
                                        {
                                            (resultFileWsheet.Cells[j, k] as Excel.Range).Font.Bold = true;
                                            (resultFileWsheet.Cells[j, k] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbGreen;

                                        }
                                        else
                                        {
                                            (resultFileWsheet.Cells[j, k] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbRed;
                                        }
                                    }
                                    else if (resultFileWsheet.Cells[j, k].value == null && expFilewsheet.Cells[j, k].value != null)
                                    {
                                        (resultFileWsheet.Cells[j, k] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbRed;
                                    }
                                    else if (resultFileWsheet.Cells[j, k].value != null && expFilewsheet.Cells[j, k].value == null)
                                    {
                                        (resultFileWsheet.Cells[j, k] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbRed;
                                    }
                                    else
                                    {
                                        (resultFileWsheet.Cells[j, k] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbGreen;
                                    }

                                }
                            }

                            resultFilewbook.Save();
                            pathFileWbook.Save();

                            resultFilewbook.Close(false, null, null);
                            Marshal.ReleaseComObject(resultFilewbook);
                            Marshal.ReleaseComObject(resultFileWsheet);

                            Marshal.ReleaseComObject(expFilewbook);
                            Marshal.ReleaseComObject(expFilewsheet);
                        }
                        else
                        {
                            pathFileSheet.Cells[i, 4].value = "";
                            pathFileSheet.Cells[i, 5].value = "Row Count Not Matched";
                            pathFileWbook.Save();
                        }
                    }
                }
                //path file
                pathFileWbook.Close(false, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(pathFileWbook);
                Marshal.ReleaseComObject(pathFileSheet);

                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
