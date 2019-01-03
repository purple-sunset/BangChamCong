using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace BangChamCong
{
    public static class Service
    {
        private static int _currYear;
        private static int _currMonth;
        private static int _totalDay;
        public static void ProcessData(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheets listSheets = package.Workbook.Worksheets;

                // Tính thời gian vào ra
                ExcelWorksheet inOutSheet = listSheets.First(x => x.Name.Equals(AppSettingConstant.InOutSheet));
                Dictionary<string, List<InOutData>> inOutData = ProcessInOutData(inOutSheet);
                _currYear = inOutData.First().Value.First().InTime.Month;
                _currMonth = inOutData.First().Value.First().InTime.Month;
                _totalDay = DateTime.DaysInMonth(_currYear, _currMonth);

                // Ghi vào sheet BCC
                ExcelWorksheet salarySheet = listSheets.First(x => x.Name.Equals(AppSettingConstant.SalarySheet));
                WriteSalaryInformation(salarySheet, inOutData);

                // Save
                package.Save();
            }
        }

        public static Dictionary<string, List<InOutData>> ProcessInOutData(ExcelWorksheet inOutSheet)
        {
            List<Range> inOutDataRange = GetInOutDataRange(inOutSheet);
            var result = new Dictionary<string, List<InOutData>>();
            for (var row = 1; row <= inOutSheet.Dimension.End.Row; row++)
            {
                if (inOutDataRange.Any(x => x.StartRow <= row && x.EndRow >= row))
                {
                    string day = inOutSheet.Cells[$"B{row}"].Text;
                    string employee = inOutSheet.Cells[$"D{row}"].Text.ToLower();
                    string timeIn = inOutSheet.Cells[$"F{row}"].Text;
                    string timeOut = inOutSheet.Cells[$"G{row}"].Text;

                    if (!string.IsNullOrWhiteSpace(employee) && !string.IsNullOrWhiteSpace(day))
                    {
                        var inOutItem = CreateInOutData(employee, day, timeIn, timeOut);
                        if (result.ContainsKey(employee))
                        {
                            List<InOutData> data = result[employee];
                            data.Add(inOutItem);
                        }
                        else
                        {
                            var data = new List<InOutData>();
                            data.Add(inOutItem);
                            result.Add(employee, data);
                        }
                    }
                }
            }

            return result;
        }

        public static List<Range> GetInOutDataRange(ExcelWorksheet inOutSheet)
        {
            var result = new List<Range>();
            for (var row = 1; row <= inOutSheet.Dimension.End.Row; row++)
            {
                string value = inOutSheet.Cells[$"A{row}"].Text.ToLower();
                if (AppSettingConstant.InOutStart.Contains(value))
                {
                    var lastItem = result.LastOrDefault();
                    if (lastItem != null && lastItem.EndRow == 0)
                    {
                        lastItem.EndRow = row - 1;
                    }

                    var item = new Range();
                    item.StartRow = row + 1;
                    result.Add(item);
                }
                else if (value.Contains(AppSettingConstant.InOutEnd))
                {
                    var item = result.LastOrDefault();
                    if (item != null && item.EndRow == 0)
                    {
                        item.EndRow = row - 1;
                    }
                }
            }

            return result;
        }

        public static InOutData CreateInOutData(string employee, string dayString, string timeInString,
                                                string timeOutString)
        {
            var result = new InOutData();
            result.InTime = DateTime.ParseExact($"{dayString} {timeInString}", AppSettingConstant.DateTimeFormat,
                                                CultureInfo.InvariantCulture);
            result.OutTime = DateTime.ParseExact($"{dayString} {timeOutString}", AppSettingConstant.DateTimeFormat,
                                                 CultureInfo.InvariantCulture);
            DateTime normalInTime = DateTime.ParseExact($"{dayString} {AppSettingConstant.NormalInTime}",
                                                        AppSettingConstant.DateTimeFormat,
                                                        CultureInfo.InvariantCulture);
            DateTime normalMorningOutTime =
                DateTime.ParseExact($"{dayString} {AppSettingConstant.NormalMorningOutTime}",
                                    AppSettingConstant.DateTimeFormat, CultureInfo.InvariantCulture);
            DateTime normalAfternoonInTime =
                DateTime.ParseExact($"{dayString} {AppSettingConstant.NormalAfternoonInTime}",
                                    AppSettingConstant.DateTimeFormat, CultureInfo.InvariantCulture);
            DateTime normalOutTime = DateTime.ParseExact($"{dayString} {AppSettingConstant.NormalOutTime}",
                                                         AppSettingConstant.DateTimeFormat,
                                                         CultureInfo.InvariantCulture);
            result.IsOTDay = result.InTime.DayOfWeek == DayOfWeek.Sunday;

            // Chủ nhật
            if (result.IsOTDay)
            {
                // Đi làm vào buổi sáng
                if (result.InTime < normalMorningOutTime)
                {
                    // Làm việc đến chiều
                    if (result.OutTime > normalAfternoonInTime)
                    {
                        result.OTHour += (normalMorningOutTime - result.InTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {result.InTime:hh:mm}-{normalMorningOutTime:hh:mm}; ";
                        result.OTHour += (result.OutTime - normalAfternoonInTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {normalAfternoonInTime:hh:mm}-{result.OutTime:hh:mm}; ";
                    }
                    // Chỉ làm buổi sáng
                    else
                    {
                        DateTime outTime = result.OutTime < normalMorningOutTime
                            ? result.OutTime
                            : normalMorningOutTime;
                        result.OTHour += (outTime - result.InTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {result.InTime:hh:mm}-{outTime:hh:mm}; ";
                    }
                }
                // Chỉ làm buổi chiều
                else
                {
                    DateTime inTime = result.InTime > normalAfternoonInTime ? result.InTime : normalAfternoonInTime;
                    result.OTHour += (result.OutTime - inTime).TotalHours;
                    result.Note += $"ngày {result.InTime.Day:D2}: {inTime:hh:mm}-{result.OutTime:hh:mm}; ";
                }
            }
            // Ngày thường
            else
            {
                // Check đi muộn, về sớm
                if (result.InTime > normalMorningOutTime)
                {
                    result.Comment += "Nghỉ sáng \n";
                    result.WorkDay -= 0.5;
                }
                else if (result.InTime > normalInTime.AddMinutes(AppSettingConstant.LateGapMinute))
                {
                    result.Comment += "Đi muộn \n";
                    //result.WorkDay -= (result.InTime - normalInTime).TotalHours / 8;
                }

                if (result.OutTime < normalAfternoonInTime)
                {
                    result.Comment += "Nghỉ chiều \n";
                    result.WorkDay -= 0.5;
                }
                else if (result.OutTime < normalOutTime.AddMinutes(-AppSettingConstant.EarlyGapMinute))
                {
                    result.Comment += "Về sớm \n";
                    result.WorkDay -= (normalOutTime - result.OutTime).TotalHours / 8;
                }

                // Tính thời gian OT
                if (result.OutTime > normalOutTime.AddMinutes(AppSettingConstant.OTGapMinute))
                {
                    // Quản lý
                    if (AppSettingConstant.ManageList.Contains(employee))
                    {
                    }
                    // Part time
                    else if (AppSettingConstant.InternList.Contains(employee))
                    {
                        //result.OTHour += (result.OutTime - normalOutTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {normalOutTime:hh:mm}-{result.OutTime:hh:mm}; ";
                    }
                    else
                    {
                        result.OTHour += (result.OutTime - normalOutTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {normalOutTime:hh:mm}-{result.OutTime:hh:mm}; ";
                    }
                }
            }

            return result;
        }

        public static void WriteSalaryInformation(ExcelWorksheet salarySheet, Dictionary<string, List<InOutData>> inOutData)
        {
            for (var row = 1; row <= salarySheet.Dimension.End.Row; row++)
            {
                string employee = salarySheet.Cells[row, AppSettingConstant.NameColumn].Text.ToLower();
                if (inOutData.ContainsKey(employee))
                {
                    var data = inOutData[employee];

                    // Ghi ngày làm việc
                    for (var day = 1; day <= _totalDay; day++)
                    {
                        var dataItem = data.FirstOrDefault(x => x.InTime.Day == day);
                        var column = AppSettingConstant.StartDayColumn + day - 1;

                        // Có dữ liệu quẹt thẻ
                        if (dataItem != null)
                        {
                            salarySheet.Cells[row, column].Value = Math.Round(dataItem.WorkDay, 1, MidpointRounding.AwayFromZero);
                            if (string.IsNullOrWhiteSpace(dataItem.Comment))
                            {
                                salarySheet.Cells[row, column].Comment.Author = AppSettingConstant.Author;
                                salarySheet.Cells[row, column].Comment.Text = dataItem.Comment;
                            }
                        }
                        // Không có dữ liệu
                        else
                        {
                            var date = new DateTime(_currYear, _currMonth, day);
                            if (date.DayOfWeek != DayOfWeek.Sunday)
                            {
                                // Check không phải part time
                                if (!AppSettingConstant.InternList.Contains(employee))
                                {
                                    salarySheet.Cells[row, column].Value = 0;
                                    salarySheet.Cells[row, column].Comment.Author = AppSettingConstant.Author;
                                    salarySheet.Cells[row, column].Comment.Text = "Nghỉ có phép";
                                }
                            }
                        }
                    }

                    // Ghi thông tin OT
                    var otColumn = AppSettingConstant.StartDayColumn + _totalDay + 6;
                    var noteColumn = AppSettingConstant.StartDayColumn + _totalDay + 8;

                    salarySheet.Cells[row, otColumn].Value = Math.Round(data.Sum(x => x.OTHour), 2, MidpointRounding.AwayFromZero);
                    if (data.Any(x => string.IsNullOrWhiteSpace(x.Note)))
                    {
                        var note = new StringBuilder("Tăng ca ");
                        foreach (var dataItem in data)
                        {
                            if (!string.IsNullOrWhiteSpace(dataItem.Note))
                            {
                                note.Append(dataItem.Note);
                            }
                        }

                        salarySheet.Cells[row, noteColumn].Value = note.ToString();
                    }
                }
            }
        }
    }
}