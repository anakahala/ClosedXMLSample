using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ClosedXMLSample
{
    class Program
    {
        /// <summary>
        /// 野菜の一覧を作成します
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            // Workbookを作成する
            using (XLWorkbook workbook = new XLWorkbook())
            {
                // Worksheetを追加
                IXLWorksheet worksheet = workbook.AddWorksheet("Sheet1");

                int row = 1;

                // 1行目にヘッダー部を出力
                worksheet.Cell(row, 1).SetValue("No");
                worksheet.Cell(row, 2).SetValue("Yasai Name");
                worksheet.Cell(row, 3).SetValue("Price");
                row++;

                // 2行目以降に野菜一覧作成
                worksheet.Cell(row, 1).SetValue(row - 1);
                worksheet.Cell(row, 2).SetValue("Carrot");
                worksheet.Cell(row, 3).SetValue(120);
                row++;

                worksheet.Cell(row, 1).SetValue(row - 1);
                worksheet.Cell(row, 2).SetValue("Tomato");
                worksheet.Cell(row, 3).SetValue(220);
                row++;

                worksheet.Cell(row, 1).SetValue(row - 1);
                worksheet.Cell(row, 2).SetValue("Cabbage");
                worksheet.Cell(row, 3).SetValue(100);

                //出力した表の内側に罫線を引く
                worksheet.Range(1, 1, row, 3).Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                //出力した表の外側に罫線を引く
                worksheet.Range(1, 1, row, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                // ヘッダー(1行目)に背景色を付ける
                worksheet.Range(1, 1, 1, 3).Style.Fill.BackgroundColor = XLColor.SkyBlue;

                // 名前をつけてブックを保存
                workbook.SaveAs("ClosedXMLSample.xlsx");
            }
        }
    }
}