//This program was created by Jesus Guerrero. It reads in an product list xlsx file and converts it to csv.
using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace CSVassessment
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Add Dependacies
            string path = Environment.CurrentDirectory + "/../../../ProductList.xlsx";
            _Application excel = new _Excel.Application();
            Workbook wb;
            Worksheet ws;
            excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet).SaveAs(Environment.CurrentDirectory + "/Error.xlsx");
            excel.Workbooks.Close();

            //Adding header row to Error file
            wb = excel.Workbooks.Open(Environment.CurrentDirectory + "/Error.xlsx");
            ws = wb.Worksheets[1];
            ws.Cells[1, 1].Value2 = "PID";
            ws.Cells[1, 2].Value2 = "Product ID";
            ws.Cells[1, 3].Value2 = "Mfr Name";
            ws.Cells[1, 4].Value2 = "Mfr P/N";
            ws.Cells[1, 5].Value2 = "Price";
            ws.Cells[1, 6].Value2 = "COO";
            ws.Cells[1, 7].Value2 = "Short Description";
            ws.Cells[1, 8].Value2 = "UPC";
            ws.Cells[1, 9].Value2 = "UOM";
            wb.Save();
            excel.Workbooks.Close();
            int i = 1;

            //Opening Product list and defining variables
            wb = excel.Workbooks.Open(@path);
            ws = wb.Worksheets[1];
            bool errorLine = false;
            int x = 1, y = 0, linesRead = 0;
            double pid = 0, cost = 0;
            string productId = "", mfrName = "", mfrpn = "", coo = "", description = "", upc = "", uom = "";
            string outPath = y + "outProductList.csv";

            //Adds header to first csv file
            addHeader("PID","Product ID","Mfr Name","Mfr P/N","Price","COO","Short Description","UPC","UOM", @outPath);

            Console.WriteLine("Reading Product List..");
            //Read excel file while the row is not empty
            while (x < ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row)
            {
                x++;
                errorLine = false;

                //Read PID
                if (ws.Cells[x, 1].Value2 != null)
                {
                    try
                    {
                        pid = ws.Cells[x, 1].Value2;
                    }
                    catch (Exception ex)
                    {
                        //Failed to read as a number
                        pid = 0;
                        errorLine = true;
                    }
                }
                else
                {
                    pid = 0;
                    errorLine = true;
                }

                //Read Product ID
                if (ws.Cells[x, 2].Value2 != null)
                {
                    productId = ws.Cells[x, 2].Value2;
                }
                else
                {
                    productId = ws.Cells[x, 2].Value2;
                    errorLine = true;
                }

                //Read Mfr Name
                if (ws.Cells[x, 3].Value2 != null)
                {
                    mfrName = ws.Cells[x, 3].Value2;
                }
                else
                {
                    mfrName = ws.Cells[x, 3].Value2;
                    errorLine = true;
                }

                //Read Mfr P/N
                if (ws.Cells[x, 4].Value2 != null)
                {
                    mfrpn = ws.Cells[x, 4].Value2;
                }
                else
                {
                    mfrpn = ws.Cells[x, 4].Value2;
                    errorLine = true;
                }

                //Read Cost and convert to price
                if (ws.Cells[x, 5].Value2 != null)
                {
                    try
                    {
                        cost = ws.Cells[x, 5].Value2;
                        cost = cost * 1.2;
                    }
                    catch (Exception ex)
                    {
                        //Failed to read as a number
                        cost = 0;
                        errorLine = true;
                    }
                }
                else
                {
                    cost = 0;
                    errorLine = true;
                }

                //Read COO
                if (ws.Cells[x, 6].Value2 != null)
                {
                    coo = ws.Cells[x, 6].Value2;
                }
                else
                {
                    coo = "TW";
                    errorLine = true;
                }

                //Read description
                if (ws.Cells[x, 7].Value2 != null)
                {
                    description = ws.Cells[x, 7].Value2;
                }
                else
                {
                    description = ws.Cells[x, 7].Value2;
                    errorLine = true;
                }


                upc = ws.Cells[x, 8].Value2;

                //Read UOM
                if (ws.Cells[x, 9].Value2 != null)
                {
                    uom = ws.Cells[x, 9].Value2;
                }
                else
                {
                    uom = "EA";
                    errorLine = true;
                }
                
                //If there was an error reading the row, add it to to the error file
                if (errorLine)
                {
                    i++;
                    excel.Workbooks.Close();
                    wb = excel.Workbooks.Open(Environment.CurrentDirectory + "/Error.xlsx");
                    ws = wb.Worksheets[1];
                    ws.Cells[i, 1].Value2 = pid;
                    ws.Cells[i, 2].Value2 = productId;
                    ws.Cells[i, 3].Value2 = mfrName;
                    ws.Cells[i, 4].Value2 = mfrpn;
                    ws.Cells[i, 5].Value2 = cost;
                    ws.Cells[i, 6].Value2 = coo;
                    ws.Cells[i, 7].Value2 = description;
                    ws.Cells[i, 8].Value2 = upc;
                    ws.Cells[i, 9].Value2 = uom;
                    wb.Save();
                    excel.Workbooks.Close();
                    wb = excel.Workbooks.Open(@path);
                    ws = wb.Worksheets[1];
                }
                else
                {
                    //Else add the product to csv file and create new ones if it exceeds 10,000
                    linesRead++;
                    if(linesRead > 10000)
                    {
                        y++;
                        linesRead = 0;
                        addHeader("PID", "Product ID", "Mfr Name", "Mfr P/N", "Price", "COO", "Short Description", "UPC", "UOM", y + "outProductList.csv");
                    }
                    addRecord(pid, productId, mfrName, mfrpn, cost, coo, description, upc, uom, y + "outProductList.csv");
                }
            }
            excel.Workbooks.Close();
            excel.Quit();
            Console.WriteLine("Reading Complete");
        }

        //Method to add record to csv file
        public static void addRecord(double pid, string productId, string mfrName, string mfrpn, double cost, string coo, string description, string upc, string uom, string filepath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@filepath, true))
            {
                file.WriteLine(pid + "^" + productId + "^" + mfrName + "^" + mfrpn + "^" + cost + "^" + coo + "^" + description + "^" + upc + "^" + uom);
            }
        }

        //Method to add header to csv file
        public static void addHeader(string pid, string productId, string mfrName, string mfrpn, string cost, string coo, string description, string upc, string uom, string filepath)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@filepath, true))
            {
                file.WriteLine(pid + "^" + productId + "^" + mfrName + "^" + mfrpn + "^" + cost + "^" + coo + "^" + description + "^" + upc + "^" + uom);
            }
        }
    }
}
