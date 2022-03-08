using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


static void GenerateRow(IWorkbook workbook, ISheet sheet1, int rowid, string firstName, string lastName, double salaryAmount, double taxRate)
{
    var row = sheet1.CreateRow(rowid);
    row.CreateCell(0).SetCellValue(firstName);  //A2
    row.CreateCell(1).SetCellValue(lastName);   //B2

    // format decimal
    var style = workbook.CreateCellStyle();
    style.DataFormat = workbook.CreateDataFormat().GetFormat("##,##0.00");
    
    var cell2 = row.CreateCell(2);
    cell2.CellStyle = style;
    cell2.SetCellValue(salaryAmount);   //C2

    var cell3 = row.CreateCell(3);
    cell3.CellStyle = style;
    cell3.SetCellValue(taxRate);        //D2

    row.CreateCell(4).SetCellFormula(string.Format("C{0}*D{0}", rowid + 1));
    row.CreateCell(5).SetCellFormula(string.Format("C{0}-E{0}", rowid + 1));
}


Console.WriteLine("Hello, World!");
var wb = new XSSFWorkbook();
var s1 = wb.CreateSheet("Monthly Salary Report");
var headerRow = s1.CreateRow(0);
headerRow.CreateCell(0).SetCellValue("First Name");
s1.SetColumnWidth(0, 20 * 256);
headerRow.CreateCell(1).SetCellValue("Last Name");
s1.SetColumnWidth(1, 20 * 256);
headerRow.CreateCell(2).SetCellValue("Salary");
headerRow.CreateCell(3).SetCellValue("Tax Rate");
headerRow.CreateCell(4).SetCellValue("Tax");
headerRow.CreateCell(5).SetCellValue("Delivery");

int row = 1;
GenerateRow(wb, s1, row++, "Bill", "Zhang", 5000.1234, 9.0 / 100);
GenerateRow(wb, s1, row++, "Amy", "Huang", 8000.1001, 11.0 / 100);
GenerateRow(wb, s1, row++, "Tomos", "Johnson", 6000.2020, 9.0 / 100);
GenerateRow(wb, s1, row++, "Macro", "Jeep", 12000.12, 15.0 / 100);
s1.ForceFormulaRecalculation = true;

var fs = File.Create("test.xlsx");
wb.Write(fs);
fs.Close();