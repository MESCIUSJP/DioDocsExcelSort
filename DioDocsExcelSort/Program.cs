// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;

Console.WriteLine("DioDocs for Excelのソート機能");

// 新しいワークブックを作成します
Workbook workbook = new();

// Excelファイルを開きます
workbook.Open("test.xlsx");
IWorksheet worksheet = workbook.Worksheets[0];

//// C列の値でソートします
//worksheet.Range["A2:F21"].Sort(worksheet.Range["C2:C21"],
//                               orientation: SortOrientation.Columns);

// 複数列（C列、F列）の値でソートします
worksheet.Range["A2:F21"].Sort(
    SortOrientation.Columns,
    false,
    new ValueSortField[] {
        new(worksheet.Range["C2:C21"], SortOrder.Ascending),
        new(worksheet.Range["F2:F21"], SortOrder.Descending)
    });

// ワークブックをExcelファイルとして保存します
workbook.Save("result.xlsx");
