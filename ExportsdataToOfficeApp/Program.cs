using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportsdataToOfficeApp
{
    using Excel = Microsoft.Office.Interop.Excel;
    class Program
    {
        static void Main(string[] args)
        {
            List<Car> carsInStock = new List<Car>
            {
                new Car{Color="Green",Make="VM",PetName="Mary"},
                new Car{Color="Red",Make="Camry",PetName="Pencil Light"},
                new Car{Color="Blue",Make="Nissan",PetName="BigBoy"},
                new Car{Color="Whit",Make="Volks",PetName="Bitto"},


            };
            ExportToExcel(carsInStock);
        }

        private static void ExportToExcel(List<Car> carsInTheStock)
        {
            //load excel and then make a new empty workbook
            Excel.Application excelApp = new Excel.Application();

            //making excel visible on the computer
            excelApp.Visible = true;
            excelApp.Workbooks.Add();

            //this example uses a single worksheet.
            Excel._Worksheet worksheet = excelApp.ActiveSheet;

            //establish column headings in cells
            worksheet.Cells[1, "A"] = "Make";
            worksheet.Cells[1, "B"] = "Color";
            worksheet.Cells[1, "C"] = "Pet Name";

            //now map all data in the list<Car> to the cells of spreadsheet.
            int row = 1;
            foreach(Car c in carsInTheStock)
            {
                row++;
                worksheet.Cells[row, "A"] = c.Make;
                worksheet.Cells[row, "B"] = c.Color;
                worksheet.Cells[row, "C"] = c.PetName;

            }

            //give the a nice look and feel
            worksheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            //save the file, quit and display msg to user
            worksheet.SaveAs($@"{Environment.CurrentDirectory}\Inventory.xlsx");

            excelApp.Quit();
            Console.WriteLine("the inventory.xlsx file has been saved to your app folder");
            Console.ReadLine();

        }
    }
}
