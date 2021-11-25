using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace lab_1
{
    class MyExcel
    {
        static public Excel.Application app = null;
        static public Excel.Workbook workbook = null;
        static public Excel.Worksheet worksheet = null;
        public Excel.Range workSheet_range = null;

        public MyExcel()
        {
            ReadDoc();
        }

        public Excel.Worksheet ReadDoc()
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Open("D:/Users/liza2/Desktop/7 семестр/ИС/IS_LAB1-main/IS_LAB1-main/bin/Debug/DataBase.xlsx", 
                    Type.Missing, true, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                return worksheet;
            }
            catch (Exception e)
            {
                return worksheet;
            }
        }

        public void GetRecord(List<Car> DB)
        {
            try
            {
                int line = 1;
                do
                {
                    string brand = worksheet.Cells[line, "A"].Text.ToString();
                    string name = worksheet.Cells[line, "B"].Text.ToString();
                    long price = Convert.ToInt64(worksheet.Cells[line, "C"].Text);
                    int year = Convert.ToInt32(worksheet.Cells[line, "D"].Text);
                    string transmission = worksheet.Cells[line, "E"].Text.ToString();
                    string drive = worksheet.Cells[line, "F"].Text.ToString();
                    string raw_color = worksheet.Cells[line, "G"].Text.ToString();
                    string complect = worksheet.Cells[line, "H"].Text.ToString();
                    string auto_class = worksheet.Cells[line, "I"].Text.ToString();
                    Color color = Color.Any;
                    switch (raw_color)
                    {
                        case "Black":
                            color = Color.Black;
                            break;
                        case "White":
                            color = Color.White;
                            break;
                        case "Blue":
                            color = Color.Blue;
                            break;
                        case "Green":
                            color = Color.Green;
                            break;
                        case "Red":
                            color = Color.Red;
                            break;
                        case "Grey":
                            color = Color.Grey;
                            break;
                        case "Yellow":
                            color = Color.Yellow;
                            break;
                        case "Orange":
                            color = Color.Orange;
                            break;
                    }

                    line++;

                    DB.Add(new Car(brand, name, price, year, transmission, drive, color, complect, auto_class));

                } while (worksheet.Cells[line, 1].Text != null);
            }
            catch (Exception e)
            {
               
            }
        }

    }

}
