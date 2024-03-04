using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace UW04Mar
{
    internal class Program
    {
        static void Main()
        {
            string excelFilePath = "E:\\VSV\\Upwork\\11.xlsx";

            List<RoomEntry> roomEntries = ReadExcelFile(excelFilePath);

            if (roomEntries != null && roomEntries.Any())
            {
                int mostOccupiedRoom = FindMostOccupiedRoom(roomEntries);
                Console.WriteLine($"Самая занятая комната: {mostOccupiedRoom}");
            }
            else
            {
                Console.WriteLine("Не удалось прочитать данные из файла Excel.");
            }
        }

        static List<RoomEntry> ReadExcelFile(string filePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Установите контекст лицензирования

                FileInfo file = new FileInfo(filePath);
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    List<RoomEntry> roomEntries = new List<RoomEntry>();

                    for (int row = 2; row < rowCount; row++) // Начинаем с 2 строки, предполагая, что первая строка - заголовок
                    {
                        int roomNumber = int.Parse(worksheet.Cells[row, 1].Value.ToString());
                        DateTime entryTime = DateTime.Parse(worksheet.Cells[row, 2].Value.ToString());
                        DateTime exitTime = DateTime.Parse(worksheet.Cells[row, 3].Value.ToString());

                        RoomEntry entry = new RoomEntry
                        {
                            RoomNumber = roomNumber,
                            EntryTime = entryTime,
                            ExitTime = exitTime
                        };

                        roomEntries.Add(entry);
                    }

                    return roomEntries;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при чтении файла Excel: {ex.Message}");
                return null;
            }
        }

        static int FindMostOccupiedRoom(List<RoomEntry> roomEntries)
        {
            var roomOccupancy = new Dictionary<int, int>();

            foreach (var entry in roomEntries)
            {
                var numMinutes = (int)(entry.ExitTime - entry.EntryTime).TotalMinutes;
                if (roomOccupancy.ContainsKey(entry.RoomNumber))
                    roomOccupancy[entry.RoomNumber] = +numMinutes;
                else
                    roomOccupancy[entry.RoomNumber] = numMinutes;

                //for (DateTime currentTime = entry.EntryTime; currentTime < entry.ExitTime; currentTime = currentTime.AddMinutes(1))
                //{
                //    if (roomOccupancy.ContainsKey(entry.RoomNumber))
                //        roomOccupancy[entry.RoomNumber]++;
                //    else
                //        roomOccupancy[entry.RoomNumber] = 1;
                //}
            }

            return roomOccupancy.OrderByDescending(kv => kv.Value).FirstOrDefault().Key;
        }
    }

    class RoomEntry
    {
        public int RoomNumber { get; set; }
        public DateTime EntryTime { get; set; }
        public DateTime ExitTime { get; set; }
    }
}
