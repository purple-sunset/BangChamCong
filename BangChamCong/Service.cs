using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace BangChamCong
{
    public static class Service
    {
        public static void ProcessData(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var listSheets = package.Workbook.Worksheets;
                var inOutSheet = listSheets.First(x => x.Name.Equals(ExcelConstant.InOutSheet));
                var inOutData = ProcessInOutData(inOutSheet);
            }
        }

        public static Dictionary<string, List<InOutData>> ProcessInOutData(ExcelWorksheet inOutSheet)
        {
            var inOutDataRange = GetInOutDataRange(inOutSheet);
            var result = new Dictionary<string, List<InOutData>>();
            for (int row = inOutDataRange.StartRow; row <= inOutDataRange.EndRow; row++)
            {
                var day = inOutSheet.Cells[$"B{row}"].Text;
                var employee = inOutSheet.Cells[$"D{row}"].Text.ToLower();
                var timeIn = inOutSheet.Cells[$"F{row}"].Text;
                var timeOut = inOutSheet.Cells[$"G{row}"].Text;

                if (!string.IsNullOrWhiteSpace(employee) && !string.IsNullOrWhiteSpace(day))
                {
                    if (result.ContainsKey(employee))
                    {
                        var data = result[employee];
                    }
                    else
                    {
                        var data = new List<InOutData>();

                        result.Add(employee, data);
                    }
                }
            }

            return result;
        }

        public static Range GetInOutDataRange(ExcelWorksheet inOutSheet)
        {
            var result = new Range();
            for (int row = 0; row < inOutSheet.Dimension.End.Row; row++)
            {
                var value = inOutSheet.Cells[$"A{row}"].Text;
                if (value.Contains(AppSettingConstant.InOutStart))
                {
                    result.StartRow = row + 1;
                }

                if (result.StartRow > 0 && value.Contains(AppSettingConstant.InOutEnd))
                {
                    result.EndRow = row - 1;
                    break;
                }
            }

            return result;
        }

        public static InOutData CreateInOutData(string employee, string day, string timeIn, string timeOut)
        {
            var result = new InOutData();

            return result;
        }
    }
}
