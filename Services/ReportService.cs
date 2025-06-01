using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Uspevaemost_API.Services
{
    public class ReportService
    {
        private readonly IConfiguration _config;
     
        public ReportService(IConfiguration config)
        {
            _config = config;
        }
        public ReportService()
        {

        }
        public async Task<byte[]> GenerateExcelReportAsync(Models.ReportRequest request,string uchps)
        {
            var year = request.Goda.Select(y => $"'{y}'").ToList();
            var sem = request.Semestry.Select(s => $"'{s}'").ToList();
            var uo = request.Urovni.Select(u => $"'{u}'").ToList();
            var fo = request.FormyObucheniya.Select(f => $"'{f}'").ToList();
            var curs = request.Kurs.Select(c => $"'{c}'").ToList();

            Models.Requests req = new(uchps);


            var list = req.getData2(year,sem,uo,fo,curs);
            // Пример создания Excel
            using var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("Успеваемость");
            for (int i = 1; i < 10; i++)
            {
                sheet.Cells[1, i, 2, i].Merge = true;
            }

            var list_t = list;
            // основные данные
            {

                sheet.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[1, 1].Value = "УчП";
                sheet.Cells[1, 2].Value = "Группа";
                sheet.Cells[1, 3].Value = "Курс";
                sheet.Cells[1, 4].Value = "Форма обучения";
                sheet.Cells[1, 5].Value = "Уровень образования";
                sheet.Cells[1, 6].Value = "ФИО студента";
                sheet.Cells[1, 7].Value = "Гражданство";
                sheet.Cells[1, 8].Value = "Финансирование";
                sheet.Cells[1, 9].Value = "Льготы";
                sheet.Cells[2, 10].Value = "Экзаменов";
                sheet.Cells[2, 11].Value = "Зачетов с оценкой";
                sheet.Cells[2, 12].Value = "Зачетов";
                sheet.Cells[2, 13].Value = "Курсовых работ";
                sheet.Cells[2, 14].Value = "Курсовых проектов";
                sheet.Cells[2, 15].Value = "Отл (5)";
                sheet.Cells[2, 16].Value = "Хор (4)";
                sheet.Cells[2, 17].Value = "Удовл (3)";
                sheet.Cells[2, 18].Value = "Зачтено";
                sheet.Cells[2, 19].Value = "Неуд (2)";
                sheet.Cells[2, 20].Value = "Незачет";

                sheet.Cells[2, 21].Value = "Абсолютная";
                sheet.Cells[2, 22].Value = "Качественная";

                sheet.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                foreach (string[] s in list)
                {
                    int column = 1;
                    foreach (string s2 in s)
                    {
                        if (double.TryParse(s2, out double numericValue))
                        {

                            sheet.Cells[row, column].Value = numericValue;

                        }
                        else
                        {

                            sheet.Cells[row, column].Value = s2;
                        }

                        column++;
                    }
                    row++;
                }
                sheet.Cells[2, 1, 2, 36].AutoFilter = true;
                sheet.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                sheet.Cells[1, 1, row, 36].Style.Font.Size = 10;
                sheet.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    sheet.Column(i).Width = 9;
                }
                sheet.Cells[1, 36, row, 36].AutoFitColumns();
                sheet.Cells[1, 1, row, 8].AutoFitColumns();
                sheet.Column(9).Width = 30;
                sheet.Cells[1, 19, row, 19].AutoFitColumns();
                sheet.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                sheet.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                sheet.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                sheet.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                sheet.Cells[1, 10, 1, 14].Merge = true;
                sheet.Cells[1, 10].Value = "Количество";
                sheet.Cells[1, 15, 1, 20].Merge = true;
                sheet.Cells[1, 15].Value = "Количество сданных на:";
                sheet.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                sheet.Cells[1, 1, 2, 8].Style.WrapText = true;
                sheet.Cells[1, 21, 2, 36].Style.WrapText = true;
                sheet.Cells[1, 21, 1, 22].Merge = true;
                sheet.Cells[1, 21].Value = "Успеваемость (в %)";
                sheet.Cells[1, 23, 2, 23].Merge = true;
                sheet.Cells[1, 24, 2, 24].Merge = true;
                sheet.Cells[1, 24].Value = "Статус успеваемости";
                sheet.Cells[1, 25, 2, 25].Merge = true;
                sheet.Cells[1, 25].Value = "Сумма баллов";
                sheet.Cells[1, 26, 2, 26].Merge = true;
                sheet.Cells[1, 26].Value = "Количество дисциплин";
                sheet.Cells[1, 27, 2, 27].Merge = true;
                sheet.Cells[1, 27].Value = "Средний балл";
                sheet.Cells[1, 28, 2, 28].Merge = true;
                sheet.Cells[1, 28].Value = "Сумма оценок";
                sheet.Cells[1, 29, 2, 29].Merge = true;
                sheet.Cells[1, 29].Value = "Количество оценок";
                sheet.Cells[1, 30, 2, 30].Merge = true;
                sheet.Cells[1, 30].Value = "Средняя оценка";
                sheet.Cells[1, 31, 2, 31].Merge = true;
                sheet.Cells[1, 31].Value = "Количество АЗ после сессии";
                sheet.Cells[1, 32, 2, 32].Merge = true;
                sheet.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                sheet.Cells[1, 33, 2, 33].Merge = true;
                sheet.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                sheet.Cells[1, 34, 2, 34].Merge = true;
                sheet.Cells[1, 34].Value = "Результат пересдач";
                sheet.Cells[1, 35, 2, 35].Merge = true;
                sheet.Cells[1, 35].Value = "Сессия продлена до";
                sheet.Cells[1, 36, 2, 36].Merge = true;
                sheet.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    sheet.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    sheet.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (sheet.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(sheet.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(sheet.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(sheet.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(sheet.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(sheet.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(sheet.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            sheet.Cells[row, 21].Value = absol;
                            sheet.Cells[row, 22].Value = kach;
                            sheet.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            sheet.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(sheet.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(sheet.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                sheet.Cells[row, 27].Value = sum / kol;
                                sheet.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                sheet.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(sheet.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(sheet.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                sheet.Cells[row, 30].Value = sumo / kolo;
                                sheet.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                sheet.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }
            // по группам
            var svodG = package.Workbook.Worksheets.Add("по группам");
            {
                List<string[]> groups = req.getGroups(year, sem, uo, fo, curs);

                svodG.Cells[1, 1].Value = "УчП";
                svodG.Cells[1, 2].Value = "Группа";
                svodG.Cells[1, 3].Value = "Отличники";
                svodG.Cells[1, 4].Value = "Хорошисты";
                svodG.Cells[1, 5].Value = "Троечники";
                svodG.Cells[1, 6].Value = "Неуспевающие";
                svodG.Cells[1, 7].Value = "Качественная успеваемость (в %)";
                svodG.Cells[1, 8].Value = "Абсолютная успеваемость (в %)";
                svodG.Cells[1, 9].Value = "Средний балл (100)";
                svodG.Cells[1, 10].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 7)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodG.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodG.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 9)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodG.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodG.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j >= 0 && j <= 1)
                        {
                            svodG.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 2 && j <= 5)
                        {
                            int.TryParse(data[j], out int k);
                            svodG.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]) + Int32.Parse(data[5]);
                    double kach = Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    double absol = Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    if (total > 0)
                    {
                        svodG.Cells[i + 2, 7].Value = kach / total;
                        svodG.Cells[i + 2, 8].Value = absol / total;
                    }
                    else
                    {
                        svodG.Cells[i + 2, 7].Value = 0;
                        svodG.Cells[i + 2, 8].Value = 0;
                    }
                    svodG.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                    svodG.Cells[i + 2, 8].Style.Numberformat.Format = "0.00%";
                }

                svodG.Cells[1, 1, row, 10].Style.Font.Name = "Times New Roman";
                svodG.Cells[1, 1, row, 10].Style.Font.Size = 10;
                {
                    svodG.Cells[1, 1, row, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodG.Cells[1, 1, row, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodG.Cells[1, 1, row, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodG.Cells[1, 1, row, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodG.Cells[1, 1, 1, 10].Style.Font.Bold = true;
                    svodG.Cells[2, 9, row, 10].Style.Numberformat.Format = "0.00";
                    svodG.Cells[1, 1, row, 10].AutoFitColumns();
                }
            }

            // по учп
            var svogUch = package.Workbook.Worksheets.Add("по УчП");
            {
                List<string[]> groups = req.getbyUchp(year, sem, uo, fo, curs);

                svogUch.Cells[1, 1].Value = "УчП";
                svogUch.Cells[1, 2].Value = "Отличники";
                svogUch.Cells[1, 3].Value = "Хорошисты";
                svogUch.Cells[1, 4].Value = "Троечники";
                svogUch.Cells[1, 5].Value = "Неуспевающие";
                svogUch.Cells[1, 6].Value = "Качественная успеваемость (в %)";
                svogUch.Cells[1, 7].Value = "Абсолютная успеваемость (в %)";
                svogUch.Cells[1, 8].Value = "Средний балл (100)";
                svogUch.Cells[1, 9].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 6)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svogUch.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svogUch.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 8)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svogUch.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svogUch.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j == 0)
                        {
                            svogUch.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 1 && j <= 4)
                        {
                            int.TryParse(data[j], out int k);
                            svogUch.Cells[i + 2, j + 1].Value = k;
                        }

                    }
                    double total = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    double kach = Int32.Parse(data[1]) + Int32.Parse(data[2]);
                    double absol = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    if (total > 0)
                    {
                        svogUch.Cells[i + 2, 6].Value = kach / total;
                        svogUch.Cells[i + 2, 7].Value = absol / total;
                    }
                    else
                    {
                        svogUch.Cells[i + 2, 6].Value = 0;
                        svogUch.Cells[i + 2, 7].Value = 0;
                    }
                    svogUch.Cells[i + 2, 6].Style.Numberformat.Format = "0.00%";
                    svogUch.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                }

                svogUch.Cells[1, 1, row, 9].Style.Font.Name = "Times New Roman";
                svogUch.Cells[1, 1, row, 9].Style.Font.Size = 10;
                {
                    svogUch.Cells[1, 1, row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svogUch.Cells[1, 1, row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svogUch.Cells[1, 1, row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svogUch.Cells[1, 1, row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svogUch.Cells[1, 1, 1, 9].Style.Font.Bold = true;
                    svogUch.Cells[2, 8, row, 9].Style.Numberformat.Format = "0.00";
                    svogUch.Cells[1, 1, row, 9].AutoFitColumns();
                }
            }

            // по уровню
            var svodUO = package.Workbook.Worksheets.Add("по УО");
            {
                List<string[]> groups = req.getbyUO(year, sem, uo, fo, curs);

                svodUO.Cells[1, 1].Value = "Уровень образования";
                svodUO.Cells[1, 2].Value = "Отличник";
                svodUO.Cells[1, 3].Value = "Хорошисты";
                svodUO.Cells[1, 4].Value = "Троечники";
                svodUO.Cells[1, 5].Value = "Неуспевающие";
                svodUO.Cells[1, 6].Value = "Качественная успеваемость (в %)";
                svodUO.Cells[1, 7].Value = "Абсолютная успеваемость (в %)";
                svodUO.Cells[1, 8].Value = "Средний балл (100)";
                svodUO.Cells[1, 9].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 6)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodUO.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodUO.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 8)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodUO.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodUO.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j == 0)
                        {
                            svodUO.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 1 && j <= 4)
                        {
                            int.TryParse(data[j], out int k);
                            svodUO.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    double kach = Int32.Parse(data[1]) + Int32.Parse(data[2]);
                    double absol = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    if (total > 0)
                    {
                        svodUO.Cells[i + 2, 6].Value = kach / total;
                        svodUO.Cells[i + 2, 7].Value = absol / total;

                    }
                    else
                    {
                        svodUO.Cells[i + 2, 6].Value = 0;
                        svodUO.Cells[i + 2, 7].Value = 0;
                    }
                    svodUO.Cells[i + 2, 6].Style.Numberformat.Format = "0.00%";
                    svodUO.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                }

                svodUO.Cells[1, 1, row, 9].Style.Font.Name = "Times New Roman";
                svodUO.Cells[1, 1, row, 9].Style.Font.Size = 10;
                {
                    svodUO.Cells[1, 1, row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodUO.Cells[1, 1, row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodUO.Cells[1, 1, row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodUO.Cells[1, 1, row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodUO.Cells[1, 1, 1, 9].Style.Font.Bold = true;
                    svodUO.Cells[2, 8, row, 9].Style.Numberformat.Format = "0.00";
                    svodUO.Cells[1, 1, row, 9].AutoFitColumns();
                }
            }

            // по курсу
            var svodCURS = package.Workbook.Worksheets.Add("по курсу");
            {
                List<string[]> groups = req.getbyCURS(year, sem, uo, fo, curs);

                svodCURS.Cells[1, 1].Value = "Курс";
                svodCURS.Cells[1, 2].Value = "Отличник";
                svodCURS.Cells[1, 3].Value = "Хорошисты";
                svodCURS.Cells[1, 4].Value = "Троечники";
                svodCURS.Cells[1, 5].Value = "Неуспевающие";
                svodCURS.Cells[1, 6].Value = "Качественная успеваемость (в %)";
                svodCURS.Cells[1, 7].Value = "Абсолютная успеваемость (в %)";
                svodCURS.Cells[1, 8].Value = "Средний балл (100)";
                svodCURS.Cells[1, 9].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 6)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodCURS.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodCURS.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 8)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodCURS.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodCURS.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j == 0)
                        {
                            svodCURS.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 1 && j <= 4)
                        {
                            int.TryParse(data[j], out int k);
                            svodCURS.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    double kach = Int32.Parse(data[1]) + Int32.Parse(data[2]);
                    double absol = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    if (total > 0)
                    {
                        svodCURS.Cells[i + 2, 6].Value = kach / total;
                        svodCURS.Cells[i + 2, 7].Value = absol / total;
                    }
                    else
                    {
                        svodCURS.Cells[i + 2, 6].Value = 0;
                        svodCURS.Cells[i + 2, 7].Value = 0;
                    }
                    svodCURS.Cells[i + 2, 6].Style.Numberformat.Format = "0.00%";
                    svodCURS.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                }

                svodCURS.Cells[1, 1, row, 9].Style.Font.Name = "Times New Roman";
                svodCURS.Cells[1, 1, row, 9].Style.Font.Size = 10;
                {
                    svodCURS.Cells[1, 1, row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodCURS.Cells[1, 1, row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodCURS.Cells[1, 1, row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodCURS.Cells[1, 1, row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodCURS.Cells[1, 1, 1, 9].Style.Font.Bold = true;
                    svodCURS.Cells[2, 8, row, 9].Style.Numberformat.Format = "0.00";
                    svodCURS.Cells[1, 1, row, 9].AutoFitColumns();
                }
            }

            // по курсу и УО
            var svodUOCURS = package.Workbook.Worksheets.Add("по курсу и УО");
            {
                List<string[]> groups = req.getbyUOCurs(year, sem, uo, fo, curs);

                svodUOCURS.Cells[1, 1].Value = "Курс";
                svodUOCURS.Cells[1, 2].Value = "Уровень образования";
                svodUOCURS.Cells[1, 3].Value = "Отличник";
                svodUOCURS.Cells[1, 4].Value = "Хорошисты";
                svodUOCURS.Cells[1, 5].Value = "Троечники";
                svodUOCURS.Cells[1, 6].Value = "Неуспевающие";
                svodUOCURS.Cells[1, 7].Value = "Качественная успеваемость (в %)";
                svodUOCURS.Cells[1, 8].Value = "Абсолютная успеваемость (в %)";
                svodUOCURS.Cells[1, 9].Value = "Средний балл (100)";
                svodUOCURS.Cells[1, 10].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                foreach (var s in groups)
                {
                    foreach (var k in s)
                    {
                        Console.Write(k + " ");
                    }
                    Console.WriteLine();
                }
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 7)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodUOCURS.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodUOCURS.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 9)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodUOCURS.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodUOCURS.Cells[i + 2, j + 1].Value = 0;
                            }
                        }
                        if (j < 2)
                        {
                            svodUOCURS.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 2 && j <= 5)
                        {
                            int.TryParse(data[j], out int k);
                            svodUOCURS.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]) + Int32.Parse(data[5]);
                    double kach = Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    double absol = Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    if (total > 0)
                    {
                        svodUOCURS.Cells[i + 2, 7].Value = kach / total;
                        svodUOCURS.Cells[i + 2, 8].Value = absol / total;
                    }
                    else
                    {
                        svodUOCURS.Cells[i + 2, 7].Value = 0;
                        svodUOCURS.Cells[i + 2, 8].Value = 0;
                    }
                    svodUOCURS.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                    svodUOCURS.Cells[i + 2, 8].Style.Numberformat.Format = "0.00%";
                }

                svodUOCURS.Cells[1, 1, row, 10].Style.Font.Name = "Times New Roman";
                svodUOCURS.Cells[1, 1, row, 10].Style.Font.Size = 10;
                {
                    svodUOCURS.Cells[1, 1, row, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodUOCURS.Cells[1, 1, row, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodUOCURS.Cells[1, 1, row, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodUOCURS.Cells[1, 1, row, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodUOCURS.Cells[1, 1, 1, 10].Style.Font.Bold = true;
                    svodUOCURS.Cells[2, 8, row, 10].Style.Numberformat.Format = "0.00";
                    svodUOCURS.Cells[1, 1, row, 10].AutoFitColumns();
                }
            }


            // по форме
            var svodFO = package.Workbook.Worksheets.Add("по ФО");
            {
                List<string[]> groups = req.getbyFO(year, sem, uo, fo, curs);

                svodFO.Cells[1, 1].Value = "Форма образования";
                svodFO.Cells[1, 2].Value = "Отличник";
                svodFO.Cells[1, 3].Value = "Хорошисты";
                svodFO.Cells[1, 4].Value = "Троечники";
                svodFO.Cells[1, 5].Value = "Неуспевающие";
                svodFO.Cells[1, 6].Value = "Качественная успеваемость (в %)";
                svodFO.Cells[1, 7].Value = "Абсолютная успеваемость (в %)";
                svodFO.Cells[1, 8].Value = "Средний балл (100)";
                svodFO.Cells[1, 9].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 6)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodFO.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodFO.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 8)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodFO.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodFO.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j == 0)
                        {
                            svodFO.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 1 && j <= 4)
                        {
                            int.TryParse(data[j], out int k);
                            svodFO.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    double kach = Int32.Parse(data[1]) + Int32.Parse(data[2]);
                    double absol = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    if (total > 0)
                    {
                        svodFO.Cells[i + 2, 6].Value = kach / total;
                        svodFO.Cells[i + 2, 7].Value = absol / total;
                    }

                    else
                    {
                        svodFO.Cells[i + 2, 6].Value = 0;
                        svodFO.Cells[i + 2, 7].Value = 0;
                    }
                    svodFO.Cells[i + 2, 6].Style.Numberformat.Format = "0.00%";
                    svodFO.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                }

                svodFO.Cells[1, 1, row, 9].Style.Font.Name = "Times New Roman";
                svodFO.Cells[1, 1, row, 9].Style.Font.Size = 10;
                {
                    svodFO.Cells[1, 1, row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodFO.Cells[1, 1, row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodFO.Cells[1, 1, row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodFO.Cells[1, 1, row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodFO.Cells[1, 1, 1, 9].Style.Font.Bold = true;
                    svodFO.Cells[2, 8, row, 9].Style.Numberformat.Format = "0.00";
                    svodFO.Cells[1, 1, row, 9].AutoFitColumns();
                }
            }



            // по квоте
            var svodQuote = package.Workbook.Worksheets.Add("по квоте");
            {
                List<string[]> groups = req.getbyQuote(year, sem, uo, fo, curs);

                svodQuote.Cells[1, 1].Value = "Квота";
                svodQuote.Cells[1, 2].Value = "Отличник";
                svodQuote.Cells[1, 3].Value = "Хорошисты";
                svodQuote.Cells[1, 4].Value = "Троечники";
                svodQuote.Cells[1, 5].Value = "Неуспевающие";
                svodQuote.Cells[1, 6].Value = "Качественная успеваемость (в %)";
                svodQuote.Cells[1, 7].Value = "Абсолютная успеваемость (в %)";
                svodQuote.Cells[1, 8].Value = "Средний балл (100)";
                svodQuote.Cells[1, 9].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 6)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodQuote.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodQuote.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 8)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodQuote.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodQuote.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j == 0)
                        {
                            svodQuote.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 1 && j <= 4)
                        {
                            int.TryParse(data[j], out int k);
                            svodQuote.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    double kach = Int32.Parse(data[1]) + Int32.Parse(data[2]);
                    double absol = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    if (total > 0)
                    {
                        svodQuote.Cells[i + 2, 6].Value = kach / total;
                        svodQuote.Cells[i + 2, 7].Value = absol / total;
                    }
                    else
                    {
                        svodQuote.Cells[i + 2, 6].Value = 0;
                        svodQuote.Cells[i + 2, 7].Value = 0;
                    }
                    svodQuote.Cells[i + 2, 6].Style.Numberformat.Format = "0.00%";
                    svodQuote.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                }

                svodQuote.Cells[1, 1, row, 9].Style.Font.Name = "Times New Roman";
                svodQuote.Cells[1, 1, row, 9].Style.Font.Size = 10;
                {
                    svodQuote.Cells[1, 1, row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodQuote.Cells[1, 1, row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodQuote.Cells[1, 1, row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodQuote.Cells[1, 1, row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodQuote.Cells[1, 1, 1, 9].Style.Font.Bold = true;
                    svodQuote.Cells[2, 8, row, 9].Style.Numberformat.Format = "0.00";
                    svodQuote.Cells[1, 1, row, 9].AutoFitColumns();
                }
            }

            // по гражданству
            var svodCountry = package.Workbook.Worksheets.Add("по гражданству");
            {
                List<string[]> groups = req.getbyCountry(year, sem, uo, fo, curs);

                svodCountry.Cells[1, 1].Value = "Гражданство";
                svodCountry.Cells[1, 2].Value = "Отличник";
                svodCountry.Cells[1, 3].Value = "Хорошисты";
                svodCountry.Cells[1, 4].Value = "Троечники";
                svodCountry.Cells[1, 5].Value = "Неуспевающие";
                svodCountry.Cells[1, 6].Value = "Качественная успеваемость (в %)";
                svodCountry.Cells[1, 7].Value = "Абсолютная успеваемость (в %)";
                svodCountry.Cells[1, 8].Value = "Средний балл (100)";
                svodCountry.Cells[1, 9].Value = "Средняя оценка (5)";

                List<double[]> ints = new List<double[]>();
                //ввод данных в таблицу
                int row = groups.Count() + 1;
                for (int i = 0; i < groups.Count(); i++)
                {
                    string[] data = groups[i];
                    for (int j = 0; j < data.Length; j++)
                    {
                        if (j == 6)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {

                                svodCountry.Cells[i + 2, j + 2].Value = l / k;
                            }
                            else
                            {
                                svodCountry.Cells[i + 2, j + 2].Value = 0;
                            }
                        }
                        if (j == 8)
                        {
                            if (double.TryParse(data[j], out var k) && double.TryParse(data[j - 1], out var l))
                            {
                                svodCountry.Cells[i + 2, j + 1].Value = l / k;
                            }
                            else
                            {
                                svodCountry.Cells[i + 2, j + 1].Value = 0;
                            }


                        }
                        if (j == 0)
                        {
                            svodCountry.Cells[i + 2, j + 1].Value = data[j];
                        }
                        if (j >= 1 && j <= 4)
                        {
                            int.TryParse(data[j], out int k);
                            svodCountry.Cells[i + 2, j + 1].Value = k;
                        }
                    }
                    double total = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]) + Int32.Parse(data[4]);
                    double kach = Int32.Parse(data[1]) + Int32.Parse(data[2]);
                    double absol = Int32.Parse(data[1]) + Int32.Parse(data[2]) + Int32.Parse(data[3]);
                    if (total > 0)
                    {
                        svodCountry.Cells[i + 2, 6].Value = kach / total;
                        svodCountry.Cells[i + 2, 7].Value = absol / total;
                    }
                    else
                    {
                        svodCountry.Cells[i + 2, 6].Value = 0;
                        svodCountry.Cells[i + 2, 7].Value = 0;
                    }
                    svodCountry.Cells[i + 2, 6].Style.Numberformat.Format = "0.00%";
                    svodCountry.Cells[i + 2, 7].Style.Numberformat.Format = "0.00%";
                }

                svodCountry.Cells[1, 1, row, 9].Style.Font.Name = "Times New Roman";
                svodCountry.Cells[1, 1, row, 9].Style.Font.Size = 10;
                {
                    svodCountry.Cells[1, 1, row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    svodCountry.Cells[1, 1, row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    svodCountry.Cells[1, 1, row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    svodCountry.Cells[1, 1, row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    svodCountry.Cells[1, 1, 1, 9].Style.Font.Bold = true;
                    svodCountry.Cells[2, 8, row, 9].Style.Numberformat.Format = "0.00";
                    svodCountry.Cells[1, 1, row, 9].AutoFitColumns();
                }
            }

            // по студентам с ООП
            var invData = package.Workbook.Worksheets.Add("Студенты с ООП");
            {
                list = req.getDataInv(year, sem, uo, fo, curs);
                for (int i = 1; i < 10; i++)
                {
                    invData.Cells[1, i, 2, i].Merge = true;
                }
                invData.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                invData.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                invData.Cells[1, 1].Value = "УчП";
                invData.Cells[1, 2].Value = "Группа";
                invData.Cells[1, 3].Value = "Курс";
                invData.Cells[1, 4].Value = "Форма обучения";
                invData.Cells[1, 5].Value = "Уровень образования";
                invData.Cells[1, 6].Value = "ФИО студента";
                invData.Cells[1, 7].Value = "Гражданство";
                invData.Cells[1, 8].Value = "Финансирование";
                invData.Cells[1, 9].Value = "Льготы";
                invData.Cells[2, 10].Value = "Экзаменов";
                invData.Cells[2, 11].Value = "Зачетов с оценкой";
                invData.Cells[2, 12].Value = "Зачетов";
                invData.Cells[2, 13].Value = "Курсовых работ";
                invData.Cells[2, 14].Value = "Курсовых проектов";
                invData.Cells[2, 15].Value = "Отл (5)";
                invData.Cells[2, 16].Value = "Хор (4)";
                invData.Cells[2, 17].Value = "Удовл (3)";
                invData.Cells[2, 18].Value = "Зачтено";
                invData.Cells[2, 19].Value = "Неуд (2)";
                invData.Cells[2, 20].Value = "Незачет";

                invData.Cells[2, 21].Value = "Абсолютная";
                invData.Cells[2, 22].Value = "Качественная";

                invData.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                foreach (string[] s in list)
                {
                    int column = 1;
                    foreach (string s2 in s)
                    {
                        if (double.TryParse(s2, out double numericValue))
                        {

                            invData.Cells[row, column].Value = numericValue;

                        }
                        else
                        {

                            invData.Cells[row, column].Value = s2;
                        }

                        column++;
                    }
                    row++;
                }
                invData.Cells[2, 1, 2, 36].AutoFilter = true;
                invData.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                invData.Cells[1, 1, row, 36].Style.Font.Size = 10;
                invData.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    invData.Column(i).Width = 9;
                }
                invData.Cells[1, 36, row, 36].AutoFitColumns();
                invData.Cells[1, 1, row, 8].AutoFitColumns();
                invData.Column(9).Width = 30;
                invData.Cells[1, 19, row, 19].AutoFitColumns();
                invData.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                invData.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                invData.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                invData.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                invData.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                invData.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                invData.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                invData.Cells[1, 10, 1, 14].Merge = true;
                invData.Cells[1, 10].Value = "Количество";
                invData.Cells[1, 15, 1, 20].Merge = true;
                invData.Cells[1, 15].Value = "Количество сданных на:";
                invData.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                invData.Cells[1, 1, 2, 8].Style.WrapText = true;
                invData.Cells[1, 21, 2, 36].Style.WrapText = true;
                invData.Cells[1, 21, 1, 22].Merge = true;
                invData.Cells[1, 21].Value = "Успеваемость (в %)";
                invData.Cells[1, 23, 2, 23].Merge = true;
                invData.Cells[1, 24, 2, 24].Merge = true;
                invData.Cells[1, 24].Value = "Статус успеваемости";
                invData.Cells[1, 25, 2, 25].Merge = true;
                invData.Cells[1, 25].Value = "Сумма баллов";
                invData.Cells[1, 26, 2, 26].Merge = true;
                invData.Cells[1, 26].Value = "Количество дисциплин";
                invData.Cells[1, 27, 2, 27].Merge = true;
                invData.Cells[1, 27].Value = "Средний балл";
                invData.Cells[1, 28, 2, 28].Merge = true;
                invData.Cells[1, 28].Value = "Сумма оценок";
                invData.Cells[1, 29, 2, 29].Merge = true;
                invData.Cells[1, 29].Value = "Количество оценок";
                invData.Cells[1, 30, 2, 30].Merge = true;
                invData.Cells[1, 30].Value = "Средняя оценка";
                invData.Cells[1, 31, 2, 31].Merge = true;
                invData.Cells[1, 31].Value = "Количество АЗ после сессии";
                invData.Cells[1, 32, 2, 32].Merge = true;
                invData.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                invData.Cells[1, 33, 2, 33].Merge = true;
                invData.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                invData.Cells[1, 34, 2, 34].Merge = true;
                invData.Cells[1, 34].Value = "Результат пересдач";
                invData.Cells[1, 35, 2, 35].Merge = true;
                invData.Cells[1, 35].Value = "Сессия продлена до";
                invData.Cells[1, 36, 2, 36].Merge = true;
                invData.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    invData.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    invData.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (invData.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(invData.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(invData.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(invData.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(invData.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(invData.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(invData.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            invData.Cells[row, 21].Value = absol;
                            invData.Cells[row, 22].Value = kach;
                            invData.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            invData.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(invData.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(invData.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                invData.Cells[row, 27].Value = sum / kol;
                                invData.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                invData.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(invData.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(invData.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                invData.Cells[row, 30].Value = sumo / kolo;
                                invData.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                invData.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }


            // по АЗ по студентам с ООП
            var invDolgi = package.Workbook.Worksheets.Add("АЗ студентов с ООП");
            {
                list = req.getInvDolgi(year, sem, uo, fo, curs);

                invDolgi.Cells[1, 1, 1, 21].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                invDolgi.Cells[1, 1, 1, 21].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                invDolgi.Cells[1, 1].Value = "УчП";
                invDolgi.Cells[1, 2].Value = "Группа";
                invDolgi.Cells[1, 3].Value = "ФИО студента";
                invDolgi.Cells[1, 4].Value = "Льготы";
                invDolgi.Cells[1, 5].Value = "Уровень образования";
                invDolgi.Cells[1, 6].Value = "Форма образования";
                invDolgi.Cells[1, 7].Value = "Курс";
                invDolgi.Cells[1, 8].Value = "№ ведомости";
                invDolgi.Cells[1, 9].Value = "Дисциплина";
                invDolgi.Cells[1, 10].Value = "Форма контроля";
                invDolgi.Cells[1, 11].Value = "Преподаватель";
                invDolgi.Cells[1, 12].Value = "Срез 1";
                invDolgi.Cells[1, 13].Value = "Срез 2";
                invDolgi.Cells[1, 14].Value = "Рубежный срез";
                invDolgi.Cells[1, 15].Value = "Премиальные баллы";
                invDolgi.Cells[1, 16].Value = "Экзамен";
                invDolgi.Cells[1, 17].Value = "Балл по итогу сессии";
                invDolgi.Cells[1, 18].Value = "Оценка по итогу сессии";
                invDolgi.Cells[1, 19].Value = "Пересдачи";
                invDolgi.Cells[1, 20].Value = "Баллы после пересдачи";
                invDolgi.Cells[1, 21].Value = "Оценка после пересдачи";

                int row = 2;
                foreach (string[] s in list)
                {
                    int column = 1;
                    foreach (string s2 in s)
                    {
                        if (double.TryParse(s2, out double numericValue))
                        {

                            invDolgi.Cells[row, column].Value = numericValue;

                        }
                        else
                        {

                            invDolgi.Cells[row, column].Value = s2;
                        }

                        column++;
                    }
                    row++;
                }
                invDolgi.Cells[1, 1, 1, 21].AutoFilter = true;
                invDolgi.Cells[1, 1, row, 21].Style.Font.Name = "Times New Roman";
                invDolgi.Cells[1, 1, row, 21].Style.Font.Size = 10;

                invDolgi.Cells[1, 1, row, 21].AutoFitColumns();

                invDolgi.Cells[1, 1, row - 1, 21].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                invDolgi.Cells[1, 1, row - 1, 21].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                invDolgi.Cells[1, 1, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                invDolgi.Cells[1, 1, row - 1, 21].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                invDolgi.Cells[1, 1, 1, 21].Style.Font.Bold = true;
                invDolgi.Cells[1, 1, 1, 21].Style.WrapText = true;

                invDolgi.Column(9).Width = 30;
                invDolgi.Column(11).Width = 30;

                invDolgi.Column(4).Width = 30;
                invDolgi.Column(19).Width = 30;
            }

            // Отличники
            list = list_t;
            var Otlichniki = package.Workbook.Worksheets.Add("Отличники");
            {
                for (int i = 1; i < 10; i++)
                {
                    Otlichniki.Cells[1, i, 2, i].Merge = true;
                }
                Otlichniki.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Otlichniki.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Otlichniki.Cells[1, 1].Value = "УчП";
                Otlichniki.Cells[1, 2].Value = "Группа";
                Otlichniki.Cells[1, 3].Value = "Курс";
                Otlichniki.Cells[1, 4].Value = "Форма обучения";
                Otlichniki.Cells[1, 5].Value = "Уровень образования";
                Otlichniki.Cells[1, 6].Value = "ФИО студента";
                Otlichniki.Cells[1, 7].Value = "Гражданство";
                Otlichniki.Cells[1, 8].Value = "Финансирование";
                Otlichniki.Cells[1, 9].Value = "Льготы";
                Otlichniki.Cells[2, 10].Value = "Экзаменов";
                Otlichniki.Cells[2, 11].Value = "Зачетов с оценкой";
                Otlichniki.Cells[2, 12].Value = "Зачетов";
                Otlichniki.Cells[2, 13].Value = "Курсовых работ";
                Otlichniki.Cells[2, 14].Value = "Курсовых проектов";
                Otlichniki.Cells[2, 15].Value = "Отл (5)";
                Otlichniki.Cells[2, 16].Value = "Хор (4)";
                Otlichniki.Cells[2, 17].Value = "Удовл (3)";
                Otlichniki.Cells[2, 18].Value = "Зачтено";
                Otlichniki.Cells[2, 19].Value = "Неуд (2)";
                Otlichniki.Cells[2, 20].Value = "Незачет";

                Otlichniki.Cells[2, 21].Value = "Абсолютная";
                Otlichniki.Cells[2, 22].Value = "Качественная";

                Otlichniki.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                int counter = 0;
                foreach (string[] s in list)
                {

                    if (list[counter][23] == "Отличник")
                    {
                        int column = 1;
                        foreach (string s2 in s)
                        {
                            if (double.TryParse(s2, out double numericValue))
                            {

                                Otlichniki.Cells[row, column].Value = numericValue;

                            }
                            else
                            {

                                Otlichniki.Cells[row, column].Value = s2;
                            }

                            column++;
                        }
                        row++;
                    }
                    counter++;
                }
                Otlichniki.Cells[2, 1, 2, 36].AutoFilter = true;
                Otlichniki.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                Otlichniki.Cells[1, 1, row, 36].Style.Font.Size = 10;
                Otlichniki.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    Otlichniki.Column(i).Width = 9;
                }
                Otlichniki.Cells[1, 36, row, 36].AutoFitColumns();
                Otlichniki.Cells[1, 1, row, 8].AutoFitColumns();
                Otlichniki.Column(9).Width = 30;
                Otlichniki.Cells[1, 19, row, 19].AutoFitColumns();
                Otlichniki.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Otlichniki.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Otlichniki.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Otlichniki.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                Otlichniki.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                Otlichniki.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                Otlichniki.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                Otlichniki.Cells[1, 10, 1, 14].Merge = true;
                Otlichniki.Cells[1, 10].Value = "Количество";
                Otlichniki.Cells[1, 15, 1, 20].Merge = true;
                Otlichniki.Cells[1, 15].Value = "Количество сданных на:";
                Otlichniki.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                Otlichniki.Cells[1, 1, 2, 8].Style.WrapText = true;
                Otlichniki.Cells[1, 21, 2, 36].Style.WrapText = true;
                Otlichniki.Cells[1, 21, 1, 22].Merge = true;
                Otlichniki.Cells[1, 21].Value = "Успеваемость (в %)";
                Otlichniki.Cells[1, 23, 2, 23].Merge = true;
                Otlichniki.Cells[1, 24, 2, 24].Merge = true;
                Otlichniki.Cells[1, 24].Value = "Статус успеваемости";
                Otlichniki.Cells[1, 25, 2, 25].Merge = true;
                Otlichniki.Cells[1, 25].Value = "Сумма баллов";
                Otlichniki.Cells[1, 26, 2, 26].Merge = true;
                Otlichniki.Cells[1, 26].Value = "Количество дисциплин";
                Otlichniki.Cells[1, 27, 2, 27].Merge = true;
                Otlichniki.Cells[1, 27].Value = "Средний балл";
                Otlichniki.Cells[1, 28, 2, 28].Merge = true;
                Otlichniki.Cells[1, 28].Value = "Сумма оценок";
                Otlichniki.Cells[1, 29, 2, 29].Merge = true;
                Otlichniki.Cells[1, 29].Value = "Количество оценок";
                Otlichniki.Cells[1, 30, 2, 30].Merge = true;
                Otlichniki.Cells[1, 30].Value = "Средняя оценка";
                Otlichniki.Cells[1, 31, 2, 31].Merge = true;
                Otlichniki.Cells[1, 31].Value = "Количество АЗ после сессии";
                Otlichniki.Cells[1, 32, 2, 32].Merge = true;
                Otlichniki.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                Otlichniki.Cells[1, 33, 2, 33].Merge = true;
                Otlichniki.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                Otlichniki.Cells[1, 34, 2, 34].Merge = true;
                Otlichniki.Cells[1, 34].Value = "Результат пересдач";
                Otlichniki.Cells[1, 35, 2, 35].Merge = true;
                Otlichniki.Cells[1, 35].Value = "Сессия продлена до";
                Otlichniki.Cells[1, 36, 2, 36].Merge = true;
                Otlichniki.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    Otlichniki.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    Otlichniki.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (Otlichniki.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(Otlichniki.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(Otlichniki.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(Otlichniki.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(Otlichniki.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(Otlichniki.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(Otlichniki.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            Otlichniki.Cells[row, 21].Value = absol;
                            Otlichniki.Cells[row, 22].Value = kach;
                            Otlichniki.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            Otlichniki.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(Otlichniki.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(Otlichniki.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                Otlichniki.Cells[row, 27].Value = sum / kol;
                                Otlichniki.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                Otlichniki.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(Otlichniki.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(Otlichniki.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                Otlichniki.Cells[row, 30].Value = sumo / kolo;
                                Otlichniki.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                Otlichniki.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }

            // Хорошисты
            var Horosh = package.Workbook.Worksheets.Add("Хорошисты");
            {
                for (int i = 1; i < 10; i++)
                {
                    Horosh.Cells[1, i, 2, i].Merge = true;
                }
                Horosh.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Horosh.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Horosh.Cells[1, 1].Value = "УчП";
                Horosh.Cells[1, 2].Value = "Группа";
                Horosh.Cells[1, 3].Value = "Курс";
                Horosh.Cells[1, 4].Value = "Форма обучения";
                Horosh.Cells[1, 5].Value = "Уровень образования";
                Horosh.Cells[1, 6].Value = "ФИО студента";
                Horosh.Cells[1, 7].Value = "Гражданство";
                Horosh.Cells[1, 8].Value = "Финансирование";
                Horosh.Cells[1, 9].Value = "Льготы";
                Horosh.Cells[2, 10].Value = "Экзаменов";
                Horosh.Cells[2, 11].Value = "Зачетов с оценкой";
                Horosh.Cells[2, 12].Value = "Зачетов";
                Horosh.Cells[2, 13].Value = "Курсовых работ";
                Horosh.Cells[2, 14].Value = "Курсовых проектов";
                Horosh.Cells[2, 15].Value = "Отл (5)";
                Horosh.Cells[2, 16].Value = "Хор (4)";
                Horosh.Cells[2, 17].Value = "Удовл (3)";
                Horosh.Cells[2, 18].Value = "Зачтено";
                Horosh.Cells[2, 19].Value = "Неуд (2)";
                Horosh.Cells[2, 20].Value = "Незачет";

                Horosh.Cells[2, 21].Value = "Абсолютная";
                Horosh.Cells[2, 22].Value = "Качественная";

                Horosh.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                int counter = 0;
                foreach (string[] s in list)
                {

                    if (list[counter][23] == "Хорошист")
                    {
                        int column = 1;
                        foreach (string s2 in s)
                        {
                            if (double.TryParse(s2, out double numericValue))
                            {

                                Horosh.Cells[row, column].Value = numericValue;

                            }
                            else
                            {

                                Horosh.Cells[row, column].Value = s2;
                            }

                            column++;
                        }
                        row++;
                    }
                    counter++;
                }
                Horosh.Cells[2, 1, 2, 36].AutoFilter = true;
                Horosh.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                Horosh.Cells[1, 1, row, 36].Style.Font.Size = 10;
                Horosh.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    Horosh.Column(i).Width = 9;
                }
                Horosh.Cells[1, 36, row, 36].AutoFitColumns();
                Horosh.Cells[1, 1, row, 8].AutoFitColumns();
                Horosh.Column(9).Width = 30;
                Horosh.Cells[1, 19, row, 19].AutoFitColumns();
                Horosh.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                Horosh.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                Horosh.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                Horosh.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                Horosh.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                Horosh.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                Horosh.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                Horosh.Cells[1, 10, 1, 14].Merge = true;
                Horosh.Cells[1, 10].Value = "Количество";
                Horosh.Cells[1, 15, 1, 20].Merge = true;
                Horosh.Cells[1, 15].Value = "Количество сданных на:";
                Horosh.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                Horosh.Cells[1, 1, 2, 8].Style.WrapText = true;
                Horosh.Cells[1, 21, 2, 36].Style.WrapText = true;
                Horosh.Cells[1, 21, 1, 22].Merge = true;
                Horosh.Cells[1, 21].Value = "Успеваемость (в %)";
                Horosh.Cells[1, 23, 2, 23].Merge = true;
                Horosh.Cells[1, 24, 2, 24].Merge = true;
                Horosh.Cells[1, 24].Value = "Статус успеваемости";
                Horosh.Cells[1, 25, 2, 25].Merge = true;
                Horosh.Cells[1, 25].Value = "Сумма баллов";
                Horosh.Cells[1, 26, 2, 26].Merge = true;
                Horosh.Cells[1, 26].Value = "Количество дисциплин";
                Horosh.Cells[1, 27, 2, 27].Merge = true;
                Horosh.Cells[1, 27].Value = "Средний балл";
                Horosh.Cells[1, 28, 2, 28].Merge = true;
                Horosh.Cells[1, 28].Value = "Сумма оценок";
                Horosh.Cells[1, 29, 2, 29].Merge = true;
                Horosh.Cells[1, 29].Value = "Количество оценок";
                Horosh.Cells[1, 30, 2, 30].Merge = true;
                Horosh.Cells[1, 30].Value = "Средняя оценка";
                Horosh.Cells[1, 31, 2, 31].Merge = true;
                Horosh.Cells[1, 31].Value = "Количество АЗ после сессии";
                Horosh.Cells[1, 32, 2, 32].Merge = true;
                Horosh.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                Horosh.Cells[1, 33, 2, 33].Merge = true;
                Horosh.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                Horosh.Cells[1, 34, 2, 34].Merge = true;
                Horosh.Cells[1, 34].Value = "Результат пересдач";
                Horosh.Cells[1, 35, 2, 35].Merge = true;
                Horosh.Cells[1, 35].Value = "Сессия продлена до";
                Horosh.Cells[1, 36, 2, 36].Merge = true;
                Horosh.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    Horosh.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    Horosh.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (Horosh.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(Horosh.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(Horosh.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(Horosh.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(Horosh.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(Horosh.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(Horosh.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            Horosh.Cells[row, 21].Value = absol;
                            Horosh.Cells[row, 22].Value = kach;
                            Horosh.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            Horosh.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(Horosh.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(Horosh.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                Horosh.Cells[row, 27].Value = sum / kol;
                                Horosh.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                Horosh.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(Horosh.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(Horosh.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                Horosh.Cells[row, 30].Value = sumo / kolo;
                                Horosh.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                Horosh.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }

            // Одна 3
            var OneThree = package.Workbook.Worksheets.Add("Одна тройка");
            {
                for (int i = 1; i < 10; i++)
                {
                    OneThree.Cells[1, i, 2, i].Merge = true;
                }
                OneThree.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                OneThree.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                OneThree.Cells[1, 1].Value = "УчП";
                OneThree.Cells[1, 2].Value = "Группа";
                OneThree.Cells[1, 3].Value = "Курс";
                OneThree.Cells[1, 4].Value = "Форма обучения";
                OneThree.Cells[1, 5].Value = "Уровень образования";
                OneThree.Cells[1, 6].Value = "ФИО студента";
                OneThree.Cells[1, 7].Value = "Гражданство";
                OneThree.Cells[1, 8].Value = "Финансирование";
                OneThree.Cells[1, 9].Value = "Льготы";
                OneThree.Cells[2, 10].Value = "Экзаменов";
                OneThree.Cells[2, 11].Value = "Зачетов с оценкой";
                OneThree.Cells[2, 12].Value = "Зачетов";
                OneThree.Cells[2, 13].Value = "Курсовых работ";
                OneThree.Cells[2, 14].Value = "Курсовых проектов";
                OneThree.Cells[2, 15].Value = "Отл (5)";
                OneThree.Cells[2, 16].Value = "Хор (4)";
                OneThree.Cells[2, 17].Value = "Удовл (3)";
                OneThree.Cells[2, 18].Value = "Зачтено";
                OneThree.Cells[2, 19].Value = "Неуд (2)";
                OneThree.Cells[2, 20].Value = "Незачет";

                OneThree.Cells[2, 21].Value = "Абсолютная";
                OneThree.Cells[2, 22].Value = "Качественная";

                OneThree.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                int counter = 0;
                foreach (string[] s in list)
                {

                    if (list[counter][23] == "Успевающий с удовлетворительными оценками" && list[counter][16] == "1")
                    {
                        int column = 1;
                        foreach (string s2 in s)
                        {
                            if (double.TryParse(s2, out double numericValue))
                            {

                                OneThree.Cells[row, column].Value = numericValue;

                            }
                            else
                            {

                                OneThree.Cells[row, column].Value = s2;
                            }

                            column++;
                        }
                        row++;
                    }
                    counter++;
                }
                OneThree.Cells[2, 1, 2, 36].AutoFilter = true;
                OneThree.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                OneThree.Cells[1, 1, row, 36].Style.Font.Size = 10;
                OneThree.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    OneThree.Column(i).Width = 9;
                }
                OneThree.Cells[1, 36, row, 36].AutoFitColumns();
                OneThree.Cells[1, 1, row, 8].AutoFitColumns();
                OneThree.Column(9).Width = 30;
                OneThree.Cells[1, 19, row, 19].AutoFitColumns();
                OneThree.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                OneThree.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                OneThree.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                OneThree.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                OneThree.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                OneThree.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                OneThree.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                OneThree.Cells[1, 10, 1, 14].Merge = true;
                OneThree.Cells[1, 10].Value = "Количество";
                OneThree.Cells[1, 15, 1, 20].Merge = true;
                OneThree.Cells[1, 15].Value = "Количество сданных на:";
                OneThree.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                OneThree.Cells[1, 1, 2, 8].Style.WrapText = true;
                OneThree.Cells[1, 21, 2, 36].Style.WrapText = true;
                OneThree.Cells[1, 21, 1, 22].Merge = true;
                OneThree.Cells[1, 21].Value = "Успеваемость (в %)";
                OneThree.Cells[1, 23, 2, 23].Merge = true;
                OneThree.Cells[1, 24, 2, 24].Merge = true;
                OneThree.Cells[1, 24].Value = "Статус успеваемости";
                OneThree.Cells[1, 25, 2, 25].Merge = true;
                OneThree.Cells[1, 25].Value = "Сумма баллов";
                OneThree.Cells[1, 26, 2, 26].Merge = true;
                OneThree.Cells[1, 26].Value = "Количество дисциплин";
                OneThree.Cells[1, 27, 2, 27].Merge = true;
                OneThree.Cells[1, 27].Value = "Средний балл";
                OneThree.Cells[1, 28, 2, 28].Merge = true;
                OneThree.Cells[1, 28].Value = "Сумма оценок";
                OneThree.Cells[1, 29, 2, 29].Merge = true;
                OneThree.Cells[1, 29].Value = "Количество оценок";
                OneThree.Cells[1, 30, 2, 30].Merge = true;
                OneThree.Cells[1, 30].Value = "Средняя оценка";
                OneThree.Cells[1, 31, 2, 31].Merge = true;
                OneThree.Cells[1, 31].Value = "Количество АЗ после сессии";
                OneThree.Cells[1, 32, 2, 32].Merge = true;
                OneThree.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                OneThree.Cells[1, 33, 2, 33].Merge = true;
                OneThree.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                OneThree.Cells[1, 34, 2, 34].Merge = true;
                OneThree.Cells[1, 34].Value = "Результат пересдач";
                OneThree.Cells[1, 35, 2, 35].Merge = true;
                OneThree.Cells[1, 35].Value = "Сессия продлена до";
                OneThree.Cells[1, 36, 2, 36].Merge = true;
                OneThree.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    OneThree.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    OneThree.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (OneThree.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(OneThree.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(OneThree.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(OneThree.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(OneThree.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(OneThree.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(OneThree.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            OneThree.Cells[row, 21].Value = absol;
                            OneThree.Cells[row, 22].Value = kach;
                            OneThree.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            OneThree.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(OneThree.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(OneThree.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                OneThree.Cells[row, 27].Value = sum / kol;
                                OneThree.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                OneThree.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(OneThree.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(OneThree.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                OneThree.Cells[row, 30].Value = sumo / kolo;
                                OneThree.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                OneThree.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }

            // Троек больше 1
            var ManyThree = package.Workbook.Worksheets.Add("Троек больше 1");
            {
                for (int i = 1; i < 10; i++)
                {
                    ManyThree.Cells[1, i, 2, i].Merge = true;
                }
                ManyThree.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ManyThree.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ManyThree.Cells[1, 1].Value = "УчП";
                ManyThree.Cells[1, 2].Value = "Группа";
                ManyThree.Cells[1, 3].Value = "Курс";
                ManyThree.Cells[1, 4].Value = "Форма обучения";
                ManyThree.Cells[1, 5].Value = "Уровень образования";
                ManyThree.Cells[1, 6].Value = "ФИО студента";
                ManyThree.Cells[1, 7].Value = "Гражданство";
                ManyThree.Cells[1, 8].Value = "Финансирование";
                ManyThree.Cells[1, 9].Value = "Льготы";
                ManyThree.Cells[2, 10].Value = "Экзаменов";
                ManyThree.Cells[2, 11].Value = "Зачетов с оценкой";
                ManyThree.Cells[2, 12].Value = "Зачетов";
                ManyThree.Cells[2, 13].Value = "Курсовых работ";
                ManyThree.Cells[2, 14].Value = "Курсовых проектов";
                ManyThree.Cells[2, 15].Value = "Отл (5)";
                ManyThree.Cells[2, 16].Value = "Хор (4)";
                ManyThree.Cells[2, 17].Value = "Удовл (3)";
                ManyThree.Cells[2, 18].Value = "Зачтено";
                ManyThree.Cells[2, 19].Value = "Неуд (2)";
                ManyThree.Cells[2, 20].Value = "Незачет";

                ManyThree.Cells[2, 21].Value = "Абсолютная";
                ManyThree.Cells[2, 22].Value = "Качественная";

                ManyThree.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                int counter = 0;
                foreach (string[] s in list)
                {

                    if (list[counter][23] == "Успевающий с удовлетворительными оценками" && list[counter][16] != "1")
                    {
                        int column = 1;
                        foreach (string s2 in s)
                        {
                            if (double.TryParse(s2, out double numericValue))
                            {

                                ManyThree.Cells[row, column].Value = numericValue;

                            }
                            else
                            {

                                ManyThree.Cells[row, column].Value = s2;
                            }

                            column++;
                        }
                        row++;
                    }
                    counter++;
                }
                ManyThree.Cells[2, 1, 2, 36].AutoFilter = true;
                ManyThree.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                ManyThree.Cells[1, 1, row, 36].Style.Font.Size = 10;
                ManyThree.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    ManyThree.Column(i).Width = 9;
                }
                ManyThree.Cells[1, 36, row, 36].AutoFitColumns();
                ManyThree.Cells[1, 1, row, 8].AutoFitColumns();
                ManyThree.Column(9).Width = 30;
                ManyThree.Cells[1, 19, row, 19].AutoFitColumns();
                ManyThree.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ManyThree.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ManyThree.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ManyThree.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ManyThree.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                ManyThree.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                ManyThree.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                ManyThree.Cells[1, 10, 1, 14].Merge = true;
                ManyThree.Cells[1, 10].Value = "Количество";
                ManyThree.Cells[1, 15, 1, 20].Merge = true;
                ManyThree.Cells[1, 15].Value = "Количество сданных на:";
                ManyThree.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                ManyThree.Cells[1, 1, 2, 8].Style.WrapText = true;
                ManyThree.Cells[1, 21, 2, 36].Style.WrapText = true;
                ManyThree.Cells[1, 21, 1, 22].Merge = true;
                ManyThree.Cells[1, 21].Value = "Успеваемость (в %)";
                ManyThree.Cells[1, 23, 2, 23].Merge = true;
                ManyThree.Cells[1, 24, 2, 24].Merge = true;
                ManyThree.Cells[1, 24].Value = "Статус успеваемости";
                ManyThree.Cells[1, 25, 2, 25].Merge = true;
                ManyThree.Cells[1, 25].Value = "Сумма баллов";
                ManyThree.Cells[1, 26, 2, 26].Merge = true;
                ManyThree.Cells[1, 26].Value = "Количество дисциплин";
                ManyThree.Cells[1, 27, 2, 27].Merge = true;
                ManyThree.Cells[1, 27].Value = "Средний балл";
                ManyThree.Cells[1, 28, 2, 28].Merge = true;
                ManyThree.Cells[1, 28].Value = "Сумма оценок";
                ManyThree.Cells[1, 29, 2, 29].Merge = true;
                ManyThree.Cells[1, 29].Value = "Количество оценок";
                ManyThree.Cells[1, 30, 2, 30].Merge = true;
                ManyThree.Cells[1, 30].Value = "Средняя оценка";
                ManyThree.Cells[1, 31, 2, 31].Merge = true;
                ManyThree.Cells[1, 31].Value = "Количество АЗ после сессии";
                ManyThree.Cells[1, 32, 2, 32].Merge = true;
                ManyThree.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                ManyThree.Cells[1, 33, 2, 33].Merge = true;
                ManyThree.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                ManyThree.Cells[1, 34, 2, 34].Merge = true;
                ManyThree.Cells[1, 34].Value = "Результат пересдач";
                ManyThree.Cells[1, 35, 2, 35].Merge = true;
                ManyThree.Cells[1, 35].Value = "Сессия продлена до";
                ManyThree.Cells[1, 36, 2, 36].Merge = true;
                ManyThree.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    ManyThree.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    ManyThree.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (ManyThree.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(ManyThree.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(ManyThree.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(ManyThree.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(ManyThree.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(ManyThree.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(ManyThree.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            ManyThree.Cells[row, 21].Value = absol;
                            ManyThree.Cells[row, 22].Value = kach;
                            ManyThree.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            ManyThree.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(ManyThree.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(ManyThree.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                ManyThree.Cells[row, 27].Value = sum / kol;
                                ManyThree.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                ManyThree.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(ManyThree.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(ManyThree.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                ManyThree.Cells[row, 30].Value = sumo / kolo;
                                ManyThree.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                ManyThree.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }

            // Одна 2
            var OneTwo = package.Workbook.Worksheets.Add("Одна АЗ");
            {
                for (int i = 1; i < 10; i++)
                {
                    OneTwo.Cells[1, i, 2, i].Merge = true;
                }
                OneTwo.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                OneTwo.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                OneTwo.Cells[1, 1].Value = "УчП";
                OneTwo.Cells[1, 2].Value = "Группа";
                OneTwo.Cells[1, 3].Value = "Курс";
                OneTwo.Cells[1, 4].Value = "Форма обучения";
                OneTwo.Cells[1, 5].Value = "Уровень образования";
                OneTwo.Cells[1, 6].Value = "ФИО студента";
                OneTwo.Cells[1, 7].Value = "Гражданство";
                OneTwo.Cells[1, 8].Value = "Финансирование";
                OneTwo.Cells[1, 9].Value = "Льготы";
                OneTwo.Cells[2, 10].Value = "Экзаменов";
                OneTwo.Cells[2, 11].Value = "Зачетов с оценкой";
                OneTwo.Cells[2, 12].Value = "Зачетов";
                OneTwo.Cells[2, 13].Value = "Курсовых работ";
                OneTwo.Cells[2, 14].Value = "Курсовых проектов";
                OneTwo.Cells[2, 15].Value = "Отл (5)";
                OneTwo.Cells[2, 16].Value = "Хор (4)";
                OneTwo.Cells[2, 17].Value = "Удовл (3)";
                OneTwo.Cells[2, 18].Value = "Зачтено";
                OneTwo.Cells[2, 19].Value = "Неуд (2)";
                OneTwo.Cells[2, 20].Value = "Незачет";

                OneTwo.Cells[2, 21].Value = "Абсолютная";
                OneTwo.Cells[2, 22].Value = "Качественная";

                OneTwo.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                int counter = 0;
                foreach (string[] s in list)
                {

                    if (list[counter][23] == "Неуспевающий" && (list[counter][18] + list[counter][19] == "10" || list[counter][18] + list[counter][19] == "01"))
                    {
                        int column = 1;
                        foreach (string s2 in s)
                        {
                            if (double.TryParse(s2, out double numericValue))
                            {

                                OneTwo.Cells[row, column].Value = numericValue;

                            }
                            else
                            {

                                OneTwo.Cells[row, column].Value = s2;
                            }

                            column++;
                        }
                        row++;
                    }
                    counter++;
                }
                OneTwo.Cells[2, 1, 2, 36].AutoFilter = true;
                OneTwo.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                OneTwo.Cells[1, 1, row, 36].Style.Font.Size = 10;
                OneTwo.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    OneTwo.Column(i).Width = 9;
                }
                OneTwo.Cells[1, 36, row, 36].AutoFitColumns();
                OneTwo.Cells[1, 1, row, 8].AutoFitColumns();
                OneTwo.Column(9).Width = 30;
                OneTwo.Cells[1, 19, row, 19].AutoFitColumns();
                OneTwo.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                OneTwo.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                OneTwo.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                OneTwo.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                OneTwo.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                OneTwo.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                OneTwo.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                OneTwo.Cells[1, 10, 1, 14].Merge = true;
                OneTwo.Cells[1, 10].Value = "Количество";
                OneTwo.Cells[1, 15, 1, 20].Merge = true;
                OneTwo.Cells[1, 15].Value = "Количество сданных на:";
                OneTwo.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                OneTwo.Cells[1, 1, 2, 8].Style.WrapText = true;
                OneTwo.Cells[1, 21, 2, 36].Style.WrapText = true;
                OneTwo.Cells[1, 21, 1, 22].Merge = true;
                OneTwo.Cells[1, 21].Value = "Успеваемость (в %)";
                OneTwo.Cells[1, 23, 2, 23].Merge = true;
                OneTwo.Cells[1, 24, 2, 24].Merge = true;
                OneTwo.Cells[1, 24].Value = "Статус успеваемости";
                OneTwo.Cells[1, 25, 2, 25].Merge = true;
                OneTwo.Cells[1, 25].Value = "Сумма баллов";
                OneTwo.Cells[1, 26, 2, 26].Merge = true;
                OneTwo.Cells[1, 26].Value = "Количество дисциплин";
                OneTwo.Cells[1, 27, 2, 27].Merge = true;
                OneTwo.Cells[1, 27].Value = "Средний балл";
                OneTwo.Cells[1, 28, 2, 28].Merge = true;
                OneTwo.Cells[1, 28].Value = "Сумма оценок";
                OneTwo.Cells[1, 29, 2, 29].Merge = true;
                OneTwo.Cells[1, 29].Value = "Количество оценок";
                OneTwo.Cells[1, 30, 2, 30].Merge = true;
                OneTwo.Cells[1, 30].Value = "Средняя оценка";
                OneTwo.Cells[1, 31, 2, 31].Merge = true;
                OneTwo.Cells[1, 31].Value = "Количество АЗ после сессии";
                OneTwo.Cells[1, 32, 2, 32].Merge = true;
                OneTwo.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                OneTwo.Cells[1, 33, 2, 33].Merge = true;
                OneTwo.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                OneTwo.Cells[1, 34, 2, 34].Merge = true;
                OneTwo.Cells[1, 34].Value = "Результат пересдач";
                OneTwo.Cells[1, 35, 2, 35].Merge = true;
                OneTwo.Cells[1, 35].Value = "Сессия продлена до";
                OneTwo.Cells[1, 36, 2, 36].Merge = true;
                OneTwo.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    OneTwo.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    OneTwo.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (OneTwo.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(OneTwo.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(OneTwo.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(OneTwo.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(OneTwo.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(OneTwo.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(OneTwo.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            OneTwo.Cells[row, 21].Value = absol;
                            OneTwo.Cells[row, 22].Value = kach;
                            OneTwo.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            OneTwo.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(OneTwo.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(OneTwo.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                OneTwo.Cells[row, 27].Value = sum / kol;
                                OneTwo.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                OneTwo.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(OneTwo.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(OneTwo.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                OneTwo.Cells[row, 30].Value = sumo / kolo;
                                OneTwo.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                OneTwo.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }

            // Неудов больше 1
            var ManyTwo = package.Workbook.Worksheets.Add("АЗ больше 1");
            {
                for (int i = 1; i < 10; i++)
                {
                    ManyTwo.Cells[1, i, 2, i].Merge = true;
                }
                ManyTwo.Cells[1, 1, 2, 34].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ManyTwo.Cells[1, 1, 2, 34].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ManyTwo.Cells[1, 1].Value = "УчП";
                ManyTwo.Cells[1, 2].Value = "Группа";
                ManyTwo.Cells[1, 3].Value = "Курс";
                ManyTwo.Cells[1, 4].Value = "Форма обучения";
                ManyTwo.Cells[1, 5].Value = "Уровень образования";
                ManyTwo.Cells[1, 6].Value = "ФИО студента";
                ManyTwo.Cells[1, 7].Value = "Гражданство";
                ManyTwo.Cells[1, 8].Value = "Финансирование";
                ManyTwo.Cells[1, 9].Value = "Льготы";
                ManyTwo.Cells[2, 10].Value = "Экзаменов";
                ManyTwo.Cells[2, 11].Value = "Зачетов с оценкой";
                ManyTwo.Cells[2, 12].Value = "Зачетов";
                ManyTwo.Cells[2, 13].Value = "Курсовых работ";
                ManyTwo.Cells[2, 14].Value = "Курсовых проектов";
                ManyTwo.Cells[2, 15].Value = "Отл (5)";
                ManyTwo.Cells[2, 16].Value = "Хор (4)";
                ManyTwo.Cells[2, 17].Value = "Удовл (3)";
                ManyTwo.Cells[2, 18].Value = "Зачтено";
                ManyTwo.Cells[2, 19].Value = "Неуд (2)";
                ManyTwo.Cells[2, 20].Value = "Незачет";

                ManyTwo.Cells[2, 21].Value = "Абсолютная";
                ManyTwo.Cells[2, 22].Value = "Качественная";

                ManyTwo.Cells[1, 23].Value = "Стипендия";
                int row = 3;
                int counter = 0;
                foreach (string[] s in list)
                {

                    if (list[counter][23] == "Неуспевающий" && (list[counter][18] + list[counter][19] != "10" && list[counter][18] + list[counter][19] != "01"))
                    {
                        int column = 1;
                        foreach (string s2 in s)
                        {
                            if (double.TryParse(s2, out double numericValue))
                            {

                                ManyTwo.Cells[row, column].Value = numericValue;

                            }
                            else
                            {

                                ManyTwo.Cells[row, column].Value = s2;
                            }

                            column++;
                        }
                        row++;
                    }
                    counter++;
                }
                ManyTwo.Cells[2, 1, 2, 36].AutoFilter = true;
                ManyTwo.Cells[1, 1, row, 36].Style.Font.Name = "Times New Roman";
                ManyTwo.Cells[1, 1, row, 36].Style.Font.Size = 10;
                ManyTwo.Cells[2, 9, 2, 36].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
                for (int i = 10; i < 23; i++)
                {
                    ManyTwo.Column(i).Width = 9;
                }
                ManyTwo.Cells[1, 36, row, 36].AutoFitColumns();
                ManyTwo.Cells[1, 1, row, 8].AutoFitColumns();
                ManyTwo.Column(9).Width = 30;
                ManyTwo.Cells[1, 19, row, 19].AutoFitColumns();
                ManyTwo.Cells[1, 1, row - 1, 36].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ManyTwo.Cells[1, 1, row - 1, 36].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ManyTwo.Cells[1, 1, row - 1, 36].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ManyTwo.Cells[1, 1, row - 1, 36].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                ManyTwo.Cells[1, 10, row - 1, 10].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                ManyTwo.Cells[1, 21, row - 1, 21].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                ManyTwo.Cells[1, 31, row - 1, 31].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                ManyTwo.Cells[1, 10, 1, 14].Merge = true;
                ManyTwo.Cells[1, 10].Value = "Количество";
                ManyTwo.Cells[1, 15, 1, 20].Merge = true;
                ManyTwo.Cells[1, 15].Value = "Количество сданных на:";
                ManyTwo.Cells[1, 1, 2, 36].Style.Font.Bold = true;
                ManyTwo.Cells[1, 1, 2, 8].Style.WrapText = true;
                ManyTwo.Cells[1, 21, 2, 36].Style.WrapText = true;
                ManyTwo.Cells[1, 21, 1, 22].Merge = true;
                ManyTwo.Cells[1, 21].Value = "Успеваемость (в %)";
                ManyTwo.Cells[1, 23, 2, 23].Merge = true;
                ManyTwo.Cells[1, 24, 2, 24].Merge = true;
                ManyTwo.Cells[1, 24].Value = "Статус успеваемости";
                ManyTwo.Cells[1, 25, 2, 25].Merge = true;
                ManyTwo.Cells[1, 25].Value = "Сумма баллов";
                ManyTwo.Cells[1, 26, 2, 26].Merge = true;
                ManyTwo.Cells[1, 26].Value = "Количество дисциплин";
                ManyTwo.Cells[1, 27, 2, 27].Merge = true;
                ManyTwo.Cells[1, 27].Value = "Средний балл";
                ManyTwo.Cells[1, 28, 2, 28].Merge = true;
                ManyTwo.Cells[1, 28].Value = "Сумма оценок";
                ManyTwo.Cells[1, 29, 2, 29].Merge = true;
                ManyTwo.Cells[1, 29].Value = "Количество оценок";
                ManyTwo.Cells[1, 30, 2, 30].Merge = true;
                ManyTwo.Cells[1, 30].Value = "Средняя оценка";
                ManyTwo.Cells[1, 31, 2, 31].Merge = true;
                ManyTwo.Cells[1, 31].Value = "Количество АЗ после сессии";
                ManyTwo.Cells[1, 32, 2, 32].Merge = true;
                ManyTwo.Cells[1, 32].Value = "Количество АЗ после пересдачи №1";
                ManyTwo.Cells[1, 33, 2, 33].Merge = true;
                ManyTwo.Cells[1, 33].Value = "Количество АЗ после пересдачи №2";
                ManyTwo.Cells[1, 34, 2, 34].Merge = true;
                ManyTwo.Cells[1, 34].Value = "Результат пересдач";
                ManyTwo.Cells[1, 35, 2, 35].Merge = true;
                ManyTwo.Cells[1, 35].Value = "Сессия продлена до";
                ManyTwo.Cells[1, 36, 2, 36].Merge = true;
                ManyTwo.Cells[1, 36].Value = "Индивидуальный график";
                for (int i = 21; i < 29; i++)
                {
                    ManyTwo.Column(i).Width = 12;
                }
                for (int i = 29; i < 35; i++)
                {
                    ManyTwo.Column(i).Width = 20;
                }

                //Math
                bool f = true;
                row = 3;
                while (f)
                {
                    if (ManyTwo.Cells[row, 21].Value != null)
                    {
                        if (double.TryParse(ManyTwo.Cells[row, 15].Value.ToString(), out double otl) &
                            double.TryParse(ManyTwo.Cells[row, 16].Value.ToString(), out double hor) &
                            double.TryParse(ManyTwo.Cells[row, 17].Value.ToString(), out double tri) &
                            double.TryParse(ManyTwo.Cells[row, 19].Value.ToString(), out double dva) &
                            double.TryParse(ManyTwo.Cells[row, 18].Value.ToString(), out double sach) &
                            double.TryParse(ManyTwo.Cells[row, 20].Value.ToString(), out double nesach))
                        {
                            double summ = otl + hor + tri + dva + sach + nesach;
                            double absol = (otl + hor + tri + sach) / summ;
                            double kach = (otl + hor + sach) / summ;
                            ManyTwo.Cells[row, 21].Value = absol;
                            ManyTwo.Cells[row, 22].Value = kach;
                            ManyTwo.Cells[row, 21].Style.Numberformat.Format = "0.00%";
                            ManyTwo.Cells[row, 22].Style.Numberformat.Format = "0.00%";
                        }
                        if (double.TryParse(ManyTwo.Cells[row, 25].Value.ToString(), out double sum) &
                            double.TryParse(ManyTwo.Cells[row, 26].Value.ToString(), out double kol))
                        {
                            if (kol != 0)
                            {

                                ManyTwo.Cells[row, 27].Value = sum / kol;
                                ManyTwo.Cells[row, 27].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                ManyTwo.Cells[row, 27].Value = 0;
                            }
                        }
                        if (double.TryParse(ManyTwo.Cells[row, 28].Value.ToString(), out double sumo) &
                    double.TryParse(ManyTwo.Cells[row, 29].Value.ToString(), out double kolo))
                        {
                            if (kolo != 0)
                            {
                                ManyTwo.Cells[row, 30].Value = sumo / kolo;
                                ManyTwo.Cells[row, 30].Style.Numberformat.Format = "0.00";
                            }
                            else
                            {

                                ManyTwo.Cells[row, 30].Value = 0;
                            }

                        }
                    }
                    else
                    {
                        f = false;
                    }
                    row++;
                }
            }


            list = req.getDisc(year, sem, uo, fo, curs);

            var discs = package.Workbook.Worksheets.Add("По дисциплинам");
            {
                discs.Cells[1, 1].Value = "№ ведомости";
                discs.Cells[1, 2].Value = "ФИО студента";
                discs.Cells[1, 3].Value = "УчП";
                discs.Cells[1, 4].Value = "Группа";
                discs.Cells[1, 5].Value = "Курс";
                discs.Cells[1, 6].Value = "Льготы";
                discs.Cells[1, 7].Value = "Дисциплина";
                discs.Cells[1, 8].Value = "УчП + дисциплина";
                discs.Cells[1, 9].Value = "УчП + преподаватель";
                discs.Cells[1, 10].Value = "Преподаватель";
                discs.Cells[1, 11].Value = "УГСН";
                discs.Cells[1, 12].Value = "Код НПС";
                discs.Cells[1, 13].Value = "Наименование НПС";
                discs.Cells[1, 14].Value = "Направление";
                discs.Cells[1, 15].Value = "Баллы";
                discs.Cells[1, 16].Value = "Оценка";
                discs.Cells[1, 17].Value = "Оценка по рейтингу";
                discs.Cells[1, 18].Value = "Тип ведомости";
                discs.Cells[1, 19].Value = "Дифференцированный зачет";
                discs.Cells[1, 20].Value = "Закрыта";
                discs.Cells[1, 21].Value = "Уровень";
                discs.Cells[1, 22].Value = "Форма обучения";
                discs.Cells[1, 23].Value = "Статус студента";
                discs.Cells[1, 24].Value = "Код + НПС";
                discs.Cells[1, 25].Value = "Курс+Уровень";

                int row = 2;
                foreach (string[] s in list)
                {
                    int column = 1;
                    foreach (string s2 in s)
                    {
                        if (double.TryParse(s2, out double numericValue))
                        {
                            discs.Cells[row, column].Value = numericValue;
                        }
                        else
                        {

                            discs.Cells[row, column].Value = s2;
                        }
                        if (discs.Cells[row, column].Value.ToString() == "False")
                        {
                            discs.Cells[row, column].Value = "Нет";
                        }
                        else if (discs.Cells[row, column].Value.ToString() == "True")
                        {
                            discs.Cells[row, column].Value = "Да";
                        }
                        column++;
                    }
                    row++;
                }

                discs.Cells[1, 1, 1, 25].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                discs.Cells[1, 1, 1, 25].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                discs.Cells[1, 1, row, 25].Style.Font.Name = "Times New Roman";
                discs.Cells[1, 1, row, 25].Style.Font.Size = 10;
                discs.Cells[1, 1, row - 1, 25].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                discs.Cells[1, 1, row - 1, 25].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                discs.Cells[1, 1, row - 1, 25].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                discs.Cells[1, 1, row - 1, 25].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                discs.Cells[1, 1, 1, 25].Style.Font.Bold = true;
                discs.Cells[1, 1, 1, 25].Style.WrapText = true;
                discs.Cells[1, 25, row, 25].AutoFitColumns();

                discs.Cells[1, 1, 1, 25].AutoFilter = true;

            }
         

            return await package.GetAsByteArrayAsync();
        }
    }

}
