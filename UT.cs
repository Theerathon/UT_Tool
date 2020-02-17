using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Read_Text_2
{
    class Program
    {
        static void Main(string[] args)
        {
            string nameClass = "";
            string filepath_in = "";
            string filepath_out = "";
            string file_out = "";
            try
            {
                nameClass = args[0];
                filepath_in = args[1];
                filepath_out = args[2];
                //file_out = args[3];
            }
            catch
            {
                Console.WriteLine("Please set argument as: \n");
                Console.WriteLine("1st is the name of Class" + ", " + "2nd is the path of database" + ", " + "3rd is the path of output excel file.");
                Console.Read();
            }
            string fileMain = filepath_in + "\\" + nameClass + ".main.amd";
            string fileImplementation = filepath_in + "\\" + nameClass + ".implementation.amd";
            string[] readMain_amd = File.ReadAllLines(fileMain);
            string[] readImplementation = File.ReadAllLines(fileImplementation);
            //string[] lineData_in = File.ReadAllLines(filepath_in);
            //string[] lineData_out = File.ReadAllLines(filepath_out);
            int dem = 0, dem1 = 0, demNumber = 0, countMain = 0, countMain_increase = 0, p = 0, pp = 0, tong = 1, demType = 0, demScope = 0, demCalibrate = 0;
            //string getType = "", getCalibration = "";
            bool haveCalibration = false, haveTolerance = false;
            char space = (char)32;
            double[] number = new double[4];
            double tolerance = 0;
            int rowName = 6, columnName = 1, rowTolerance = 1, columnTolerance = 1, rowType = 2, columnType = 1, rowMax = 3;
            int columnMax = 1, rowMin = 4, columnMin = 1, rowScope = 5, columnScope = 1, rowCalibrate = 7, columnCalibrate = 1;
            List<string> listName = new List<string>();
            List<string> listType = new List<string>();
            List<string> listKind = new List<string>();
            List<string> listScope = new List<string>();
            List<string> listCalibrated = new List<string>();

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Tolerance";
            xlWorkSheet.Cells[2, 1] = "Type";
            xlWorkSheet.Cells[3, 1] = "Max";
            xlWorkSheet.Cells[4, 1] = "Min";
            xlWorkSheet.Cells[5, 1] = "Scope";
            xlWorkSheet.Cells[6, 1] = "Name of element";
            xlWorkSheet.Cells[6, 1].Columns.Autofit();
            xlWorkSheet.Cells[7, 1] = "Calibrated";

            List<string> abc = new List<string>();
            abc = readMain_amd.ToList();
            a:
            {
                foreach (string t in abc)
                {
                    if (t.Contains("<Element name="))
                    {
                        p = countMain;
                        while (!abc[p].Contains("</Element>"))
                        {
                            p++;
                            countMain_increase++;
                        }
                        break;
                    }
                    else
                    {
                        countMain++;
                        tong++;
                    }
                }
            }

            tong = countMain + countMain_increase;
            for (int i = 0; i < tong; i++)
            {
                if (abc[i].Contains("<Element name="))
                {
                    string[] a = abc[i].Split(space);
                    foreach (string k in a)
                    {
                        if (k.Contains("name="))
                        {
                            string u = k.Substring(6);
                            if (u.Contains('"'))
                            {
                                u = u.Replace('"', (char)0);
                                if (u.Contains("\0"))
                                {
                                    u = u.Replace("\0", null);
                                }
                                listName.Add(u);
                            }
                        }
                    }
                }
                else if (abc[i].Contains("<ElementAttributes"))
                {
                    string[] f = abc[i].Split(space);
                    foreach (string o in f)
                    {
                        if (o.Contains("basicModelType="))
                        {
                            string u = o.Substring(15);
                            u = u.Replace('"', (char)0);
                            if (u.Contains("\0"))
                            {
                                u = u.Replace("\0", null);
                            }
                            listType.Add(u);
                        }
                    }

                }
                else if (abc[i].Contains("<PrimitiveAttributes"))
                {
                    string[] q = abc[i].Split(space);
                    foreach (string k in q)
                    {
                        if (k.Contains("kind="))
                        {
                            string u = k.Substring(6);
                            if (u.Contains('"'))
                            {
                                listKind.Add(u.Replace('"', (char)0));
                            }
                        }
                        else if (k.Contains("scope="))
                        {
                            string u = k.Substring(7);
                            if (u.Contains('"'))
                            {
                                listScope.Add(u.Replace('"', (char)0));
                            }
                        }
                        else if (k.Contains("calibrated="))
                        {
                            string u = k.Substring(12);
                            if (u.Contains('"'))
                            {
                                listCalibrated.Add(u.Replace('"', (char)0));
                            }
                        }
                    }
                }
            }
            abc.RemoveRange(0, tong + 1);
            if (tong <= abc.Count())
            {
                tong = 1;
                countMain = 0;
                countMain_increase = 0;
                goto a;
            }

            Console.WriteLine("Name:" + "\n");
            foreach (string k in listName)
            {
                Console.WriteLine(k + "\n");
            }

            Console.WriteLine("Type:" + "\n");
            foreach (string k in listType)
            {
                Console.WriteLine(k + "\n");
            }

            Console.WriteLine("Kind:" + "\n");
            foreach (string k in listKind)
            {
                Console.WriteLine(k + "\n");
            }

            Console.WriteLine("Scope:" + "\n");
            foreach (string k in listScope)
            {
                Console.WriteLine(k + "\n");
            }

            Console.WriteLine("Calibrated:" + "\n");
            foreach (string k in listCalibrated)
            {
                Console.WriteLine(k + "\n");
            }

            foreach (string name in listName)
            {
                string removeSpaceName = name.Replace(" ", string.Empty);
                foreach (string data in readImplementation)
                {
                    if (data.Contains("<ElementImplementation elementName=\"" + removeSpaceName + "\""))
                    {
                        string[] k = data.Split(space);
                        foreach (string r in k)
                        {
                            if (r.Equals("elementName=\"" + removeSpaceName + "\"", StringComparison.Ordinal))
                            {
                                break;
                            }
                        }
                        break;
                    }
                    else
                    {
                        dem1++;
                    }
                }
                for (int i = dem1; i < dem1 + 6; i++)
                {
                    if (readImplementation[i].Contains("<PhysicalInterval"))
                    {
                        //string[] a = lineData_out[i].Split(space);
                        string[] a = readImplementation[i].Split('"');
                        foreach (string u in a)
                        {
                            if (Regex.IsMatch(u, @"\d+"))
                            {
                                double.TryParse(u, out number[demNumber]);
                                demNumber++;
                            }
                        }
                    }
                    else if (readImplementation[i].Contains("<ImplementationInterval"))
                    {
                        string[] a = readImplementation[i].Split('"');
                        foreach (string u in a)
                        {
                            if (Regex.IsMatch(u, @"\d+"))
                            {
                                double.TryParse(u, out number[demNumber]);
                                demNumber++;
                            }
                        }
                    }

                }
                if (0 != (number[3] - number[2]))
                {
                    tolerance = (number[1] - number[0]) / (number[3] - number[2]);
                    haveTolerance = true;
                }
                else
                {
                    haveTolerance = false;
                }
                decimal dtot = (decimal)tolerance;

                columnName++; columnTolerance++; columnMax++; columnMin++; columnType++; columnScope++; columnCalibrate++;

                xlWorkSheet.Cells[rowName, columnName] = removeSpaceName;
                xlWorkSheet.Cells[rowName, columnName].Columns.Autofit();

                string scope = listScope[demScope] + " " + listKind[demScope];
                if (scope.Contains("\0"))
                {
                    scope = scope.Replace("\0", null);
                }
                xlWorkSheet.Cells[rowScope, columnScope] = scope;
                demScope++;
                if (scope.Equals("local parameter"))
                {
                    xlWorkSheet.Cells[rowCalibrate, columnCalibrate] = listCalibrated[demCalibrate];
                }
                demCalibrate++;

                xlWorkSheet.Cells[rowType, columnType] = listType[demType];
                demType++;
                //set tolerance value
                if (haveTolerance)
                {
                    xlWorkSheet.Cells[rowTolerance, columnTolerance] = dtot;
                    //xlWorkSheet.Cells[rowTolerance, columnTolerance].Columns.Autofit();
                }
                else
                {
                    xlWorkSheet.Cells[rowTolerance, columnTolerance] = null;
                }

                //set range max, min
                xlWorkSheet.Cells[rowMax, columnMax] = number[1];
                xlWorkSheet.Cells[rowMin, columnMin] = number[0];

                dem = 0; dem1 = 0; demNumber = 0;
                tolerance = 0;
                haveTolerance = false;
            }

            try
            {
                xlWorkBook.SaveAs(filepath_out);
            }

            catch
            {
                //System.Runtime.InteropServices.COMException
                Console.WriteLine("Please close your opening excel file!!!");
                //Console.ReadKey();
                Console.Read();
            }
            xlWorkBook.Close(false, false, false);
            xlApp.Workbooks.Close();
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
        }
    }
}