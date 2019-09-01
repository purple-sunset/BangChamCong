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
                ExcelWorksheet inOutSheet = listSheets.First(x => x.Name.Trim().ToLower().Equals(AppSettingConstant.InOutSheet));
                Dictionary<string, List<InOutData>> inOutData = ProcessInOutData(inOutSheet);
                _currYear = inOutData.First().Value.First().InTime.Month;
                _currMonth = inOutData.First().Value.First().InTime.Month;
                _totalDay = DateTime.DaysInMonth(_currYear, _currMonth);

                // Ghi vào sheet BCC
                ExcelWorksheet salarySheet = listSheets.First(x => x.Name.Trim().ToLower().Equals(AppSettingConstant.SalarySheet));
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
                    string day = inOutSheet.Cells[$"B{row}"].GetDateString();
                    string employee = inOutSheet.Cells[$"D{row}"].GetLowerString();
                    string timeIn = inOutSheet.Cells[$"F{row}"].GetLowerString();
                    string timeOut = inOutSheet.Cells[$"G{row}"].GetLowerString();

                    if (!string.IsNullOrWhiteSpace(employee) && !string.IsNullOrWhiteSpace(day) && !string.IsNullOrWhiteSpace(timeIn) && !string.IsNullOrWhiteSpace(timeOut))
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
                string value = inOutSheet.Cells[$"A{row}"].GetLowerString();
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
                result.OutTime = NormalizeOTHour(result.OutTime);
                // Đi làm vào buổi sáng
                if (result.InTime < normalMorningOutTime)
                {
                    // Làm việc đến chiều
                    if (result.OutTime > normalAfternoonInTime)
                    {
                        result.OTHour += (normalMorningOutTime - result.InTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {result.InTime:HH\\hmm}-{normalMorningOutTime:HH\\hmm}, ";
                        result.OTHour += (result.OutTime - normalAfternoonInTime).TotalHours;
                        result.Note += $"{normalAfternoonInTime:HH\\hmm}-{result.OutTime:HH\\hmm}; ";
                    }
                    // Chỉ làm buổi sáng
                    else
                    {
                        DateTime outTime = result.OutTime < normalMorningOutTime
                            ? result.OutTime
                            : normalMorningOutTime;
                        result.OTHour += (outTime - result.InTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {result.InTime:HH\\hmm}-{outTime:HH\\hmm}; ";
                    }
                }
                // Chỉ làm buổi chiều
                else
                {
                    DateTime inTime = result.InTime > normalAfternoonInTime ? result.InTime : normalAfternoonInTime;
                    result.OTHour += (result.OutTime - inTime).TotalHours;
                    result.Note += $"ngày {result.InTime.Day:D2}: {inTime:HH\\hmm}-{result.OutTime:HH\\hmm}; ";
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
                if (NormalizeOTHour(result.OutTime) > normalOutTime.AddMinutes(AppSettingConstant.OTGapMinute))
                {
                    result.OutTime = NormalizeOTHour(result.OutTime);
                    // Quản lý
                    if (AppSettingConstant.ManageList.Contains(employee))
                    {
                    }
                    // Part time
                    else if (AppSettingConstant.InternList.Contains(employee))
                    {
                        //result.OTHour += (result.OutTime - normalOutTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {normalOutTime:HH\\hmm}-{result.OutTime:HH\\hmm}; ";
                    }
                    else
                    {
                        result.OTHour += (result.OutTime - normalOutTime).TotalHours;
                        result.Note += $"ngày {result.InTime.Day:D2}: {normalOutTime:HH\\hmm}-{result.OutTime:HH\\hmm}; ";
                    }
                }
            }

            return result;
        }

        public static void WriteSalaryInformation(ExcelWorksheet salarySheet, Dictionary<string, List<InOutData>> inOutData)
        {
            for (var row = 1; row <= salarySheet.Dimension.End.Row; row++)
            {
                string employee = salarySheet.Cells[row, AppSettingConstant.NameColumn].GetLowerString();
                if (inOutData.ContainsKey(employee))
                {
                    var data = inOutData[employee];

                    // Ghi ngày làm việc
                    for (var day = 1; day <= _totalDay; day++)
                    {
                        var date = new DateTime(_currYear, _currMonth, day);
                        if (date.DayOfWeek == DayOfWeek.Sunday)
                        {
                            continue;
                        }
                        var dataItem = data.FirstOrDefault(x => x.InTime.Day == day);
                        var column = AppSettingConstant.StartDayColumn + day - 1;

                        // Lỗi dữ liệu
                        if (data.Count(x => x.InTime.Day == day) > 1)
                        {
                            salarySheet.Cells[row, column].Value = 0;
                            if (salarySheet.Cells[row, column].Comment != null)
                            {
                                var oldComment = salarySheet.Cells[row, column].Comment.Text;
                                salarySheet.Comments.Remove(salarySheet.Cells[row, column].Comment);
                                var newComment = $"{oldComment}\nLỗi";
                                salarySheet.Cells[row, column].AddComment(newComment, AppSettingConstant.Author);
                            }
                            else
                            {
                                salarySheet.Cells[row, column].AddComment("Lỗi", AppSettingConstant.Author);
                            }
                        }
                        // Có dữ liệu quẹt thẻ
                        else if (dataItem != null)
                        {
                            salarySheet.Cells[row, column].Value = Math.Round(dataItem.WorkDay, 1, MidpointRounding.AwayFromZero);
                            if (!string.IsNullOrWhiteSpace(dataItem.Comment))
                            {
                                if (salarySheet.Cells[row, column].Comment != null)
                                {
                                    var oldComment = salarySheet.Cells[row, column].Comment.Text;
                                    salarySheet.Comments.Remove(salarySheet.Cells[row, column].Comment);
                                    var newComment = $"{oldComment}\n{dataItem.Comment}";
                                    salarySheet.Cells[row, column].AddComment(newComment, AppSettingConstant.Author);
                                }
                                else
                                {
                                    salarySheet.Cells[row, column].AddComment(dataItem.Comment, AppSettingConstant.Author);
                                }
                            }
                        }
                        // Không có dữ liệu
                        else
                        {
                            // Check không phải part time
                            if (!AppSettingConstant.InternList.Contains(employee))
                            {
                                salarySheet.Cells[row, column].Value = 0;
                                if (salarySheet.Cells[row, column].Comment != null)
                                {
                                    var oldComment = salarySheet.Cells[row, column].Comment.Text;
                                    salarySheet.Comments.Remove(salarySheet.Cells[row, column].Comment);
                                    var newComment = $"{oldComment}\nNghỉ có phép";
                                    salarySheet.Cells[row, column].AddComment(newComment, AppSettingConstant.Author);
                                }
                                else
                                {
                                    salarySheet.Cells[row, column].AddComment("Nghỉ có phép", AppSettingConstant.Author);
                                }
                            }
                        }
                    }

                    // Ghi thông tin OT
                    var otColumn = AppSettingConstant.StartDayColumn + _totalDay + 8;
                    var noteColumn = AppSettingConstant.StartDayColumn + _totalDay + 10;

                    var totalOTHour = Math.Round(data.Sum(x => x.OTHour), 2, MidpointRounding.AwayFromZero);
                    salarySheet.Cells[row, otColumn].Value = totalOTHour;
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

                        note.Append($"Tổng: {totalOTHour}h.");
                        salarySheet.Cells[row, noteColumn].Value = note.ToString();
                    }
                }
            }
        }

        public static DateTime NormalizeOTHour(DateTime outTime)
        {
            int minute = outTime.Minute;
            int normalizeMinute = minute % AppSettingConstant.NormalizeOTMinute;
            var newOutTime = outTime.AddMinutes(-normalizeMinute);
            return newOutTime;
        }

        #region Extension Method

        public static string GetDateString(this ExcelRangeBase cell)
        {
            var value = cell.Value;
            if (value != null)
            {
                if (value is DateTime)
                {
                    return (value as DateTime?).Value.ToString("dd/MM/yyyy");
                }
                else
                {
                    return value as string;
                }
            }
            else
            {
                return String.Empty;
            }
        }

        public static string GetLowerString(this ExcelRangeBase cell)
        {
            var value = cell.Value;
            if (value != null)
            {
                if (value is string)
                {
                    return (value as string).ToLower();
                }
                else
                {
                    return value.ToString().ToLower();
                }
            }
            else
            {
                return String.Empty;
            }
        }

        #endregion
    }
}