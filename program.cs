using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;


namespace Exceltesting
{
    class Program
    {
        static void Main(string[] args)
        {
            string filename;
            double[] intArray;
            int  i, q=0;
            double s;
            intArray = new double[25];
            for (i = 0; i < 25;i++ ) intArray[i] = 0;
            while (true)
            {
                q++;
                filename = "C:\\Testing\\" + q.ToString();
                filename += ".xlsx";
                
                XSSFWorkbook workbook = new XSSFWorkbook();
                FileStream file2 = new FileStream(@filename, FileMode.Open, FileAccess.Read);
                if (file2 == null) break;
                workbook = new XSSFWorkbook(file2);
                ISheet sheet = workbook.GetSheet("Sheet1");
                IRow row ;
                
                for (i = 1; i < 22; i++)
                {
                    row = sheet.GetRow(i);
                    if (row != null)
                    {
                        s = 0;
                        Console.WriteLine(row.GetCell(2).ToString());
                        for (int j = 3; j < row.LastCellNum; j++)
                        {
                            string temp = row.GetCell(j).ToString();

                            try
                            {
                                s += Convert.ToDouble(temp);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }
                        intArray[i] += s;
                        Console.WriteLine(intArray[i]);
                    }

                }
                
                for (i=1;i<22;i++)
                {
                    row = sheet.GetRow(i);
                    Console.Write(row.GetCell(2));
                    double t = intArray[i] / 17;
                    Console.WriteLine("{0:F1}",t);

                }
                file2.Close();
                workbook.Close();
                Console.ReadLine();
            }

        }
    }
}
