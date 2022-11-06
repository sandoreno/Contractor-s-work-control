using ContractorsWorkAPI.Data;
using ContractorsWorkAPI.FunkMethod;
using ContractorsWorkAPI.Model;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text;

namespace ContractorsWorkAPI.Services.Impl
{
    public class StorageService : IStorageService
    {

        private readonly DataContext _context;

        public StorageService(DataContext context)
        {
            _context = context;
        }

        public const string Compound_Makros = "Составлен(а)";

        public const string Makros_1 = "Составлен(а) по";
        public const string Makros_2 = "с учетом Дополнения №:";
        public const string Makros_3 = "№ и период сборника коэффициентов (индексов) пересчета:";
        public const string Makros_4 = "за";
        public const string Makros_5 = "года";
        //
        public const string Makros_6 = "Составлен(а) в уровне текущих (прогнозных) цен на";
        public const string Makros_7 = "г.";

        public async Task<bool> SafeFiles(IFormFile file)
        {

            if (file.Length > 0)
            {
                // метод сохранения файла
                string filePath = Path.Combine(@"C:\\cwc\\ContractorsWorkAPI\\Files", file.FileName);
                using (Stream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    file.CopyTo(fileStream);
                }
                // сохраняем путь файла в базе
                var now = DateTime.Now;
                var files = new Files();
                files.Name = file.FileName;
                files.Path = filePath;
                files.CreateDate = now;
                _context.Files.Add(files);
                await _context.SaveChangesAsync();
                // метод пасинга
                PrserMethod(filePath);
                return true;
            };

            return false;
        }

        public void PrserMethod(string filePath)
        { 
                 var stm = new TSMDictinaryModel();
                 var tbm = new List<TSMBodyModel>();

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    var IndexNumberPP = 0;
                    var IndexCipher = 0;
                    var IndexWorkName = 0;
                    var IndexMeasurement = 0;
                    var IndexQuanity = 0;
                    var IndexUnitPrice = 0;
                    var InexCorrection = 0;
                    var IndexWinter = 0;
                    var IndexBasicPrice = 0;
                    var IndexCOnversion = 0;
                    var IndexTotalPrice = 0;
                    
                    //StreamWriter st = new StreamWriter($@"C:\cwc\ContractorsWorkAPI\ParsingFiles\cap.txt");
                    //StreamWriter sw = new StreamWriter($@"C:\cwc\ContractorsWorkAPI\ParsingFiles\table.txt");
                    //StreamWriter se = new StreamWriter($@"C:\cwc\ContractorsWorkAPI\ParsingFiles\final.txt");
                    // создаем новый exel файл
                    var file_name = Path.GetFileName(filePath);
                    var now = new DateTime();
                    var model_safe = new TCMBodyForechModel();
                    var body_model = new List<TCMBodyForechModel>();
                    var ending_model = new TSMEndingModel();
                    FileInfo excelFile = new FileInfo(@$"C:\cwc\ContractorsWorkAPI\ParsingFiles\{file_name}_{now.ToString("MM_dd_yyyy_HH_mm_ss")}.xlsx");
                    xlPackage.Workbook.Worksheets.Add("Worksheet1");

                    var sb = new StringBuilder(); //this is your data

                    for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        if (row.Any(x => x == "№пп") || row.Any(x => x == "Стоим. ед. с нач., руб."))
                        {
                            var str = sb.ToString();
                            str = NormalizeWhiteSpaceForLoop(str);
                            var index = str.ToLower().IndexOf(Compound_Makros.ToLower());
                            str = str.Remove(0, index);
                            //var mac_start = str.Remove(0, index);
                            var find_data = row.ToList();
                            IndexNumberPP = find_data.FindIndex(x => x.ToLower() == "№пп".ToLower());
                            IndexCipher = find_data.FindIndex(x => x.ToLower() == "Шифр расценки и коды ресурсов".ToLower());
                            if (IndexCipher == -1)
                            {
                                IndexCipher = find_data.FindIndex(x => x.ToLower() == "Шифр, номера нормативов и коды ресурсов".ToLower());
                                stm.Compilation_Date = stm.GetStringElement(Makros_6, Makros_7, ref str);
                                //st.WriteLine($"{stm.Compilation_Date}");
                            }
                            else
                            {
                                stm.Compound = stm.GetStringElement(Makros_1, Makros_2, ref str);
                                stm.Addition = stm.GetStringElement(Makros_2, Makros_3, ref str);
                                stm.Conversion_Index = stm.GetStringElement(Makros_3, Makros_4, ref str);
                                stm.Compilation_Date = stm.GetStringElement(Makros_4, Makros_5, ref str);
                               // st.WriteLine($"{stm.Compound} {stm.Addition} {stm.Conversion_Index} {stm.Compilation_Date}");
                            }
                            IndexWorkName = find_data.FindIndex(x => x.ToLower() == "Наименование работ и затрат".ToLower());
                            IndexMeasurement = find_data.FindIndex(x => x.ToLower() == "Ед. изм.".ToLower());
                            IndexQuanity = find_data.FindIndex(x => x.ToLower() == "Кол-во единиц".ToLower());
                            IndexUnitPrice = find_data.FindIndex(x => x.ToLower() == "Цена на ед. изм., руб.".ToLower());
                            if (IndexUnitPrice == -1)
                            {
                                IndexUnitPrice = find_data.FindIndex(x => x.ToLower() == "Цена на единицу измерения, руб.".ToLower());
                            }
                            InexCorrection = find_data.FindIndex(x => x.ToLower() == "Поправочные коэффициенты".ToLower());
                            IndexWinter = find_data.FindIndex(x => x.ToLower() == "Коэффициенты зимних удорожаний".ToLower());
                            IndexBasicPrice = find_data.FindIndex(x => x.ToLower() == "Всего затрат в базисном уровне цен, руб.".ToLower());
                            if (IndexBasicPrice == -1)
                            {
                                IndexBasicPrice = find_data.FindIndex(x => x.ToLower() == "Коэффициенты пересчета".ToLower());
                            }
                            IndexCOnversion = find_data.FindIndex(x => x.ToLower() == "Коэффициенты (индексы) пересчета, нормы НР и СП".ToLower());
                            if (IndexCOnversion == -1)
                            {
                                IndexCOnversion = find_data.FindIndex(x => x.ToLower() == "ВСЕГО затрат, руб.".ToLower());
                            }
                            IndexTotalPrice = find_data.FindIndex(x => x.ToLower() == "Всего затрат в текущем уровне цен, руб.".ToLower());
                            if (IndexTotalPrice == -1)
                            {
                                IndexTotalPrice = find_data.FindIndex(x => x.ToLower() == "Справ.".ToLower());
                                rowNum += 3;
                            }
                        }
                        //sb.AppendLine(string.Join(" ", row));
                        if (myWorksheet.Cells[rowNum, IndexCipher + 1].Value?.ToString() != "2" && IndexTotalPrice != 0 && myWorksheet.Cells[rowNum, IndexCipher + 1].Value?.ToString() != "Шифр расценки и коды ресурсов" && myWorksheet.Cells[rowNum, IndexCipher + 1].Value?.ToString() != "Шифр, номера нормативов и коды ресурсов" && IndexTotalPrice != -1)
                        {
                            if (myWorksheet.Cells[rowNum, IndexWorkName + 1].Value?.ToString() == "Итого по разделу")
                            {
                                for (int i = rowNum; i <= totalRows; i++)
                                {
                                    if ((myWorksheet.Cells[i, IndexTotalPrice + 1].Value?.ToString() != null))
                                    {
                                        ending_model.TotalPrice = ($"{myWorksheet.Cells[i, IndexTotalPrice + 1].Value?.ToString()}");
                                    }
                                    else if (myWorksheet.Cells[i, IndexCOnversion + 1].Value?.ToString() != null)
                                    {
                                        ending_model.IndexCOnversion = ($"{myWorksheet.Cells[i, IndexCOnversion + 1].Value?.ToString()}");
                                    }
                                }
                                break;
                            }
                            else
                            {
                                int[] indexd = { IndexUnitPrice + 1, InexCorrection + 1, IndexWinter + 1, IndexBasicPrice + 1, IndexCOnversion + 1, IndexTotalPrice + 1 };
                                if ((myWorksheet.Cells[rowNum, IndexNumberPP + 1].Value?.ToString() != null) || (myWorksheet.Cells[rowNum, IndexTotalPrice + 1].Value?.ToString() != null) || (myWorksheet.Cells[rowNum, InexCorrection + 1].Value?.ToString()) != null)
                                {

                                    if ((myWorksheet.Cells[rowNum, IndexNumberPP + 1].Value?.ToString() != null) && (myWorksheet.Cells[rowNum, IndexNumberPP + 1].Value?.ToString() != ""))
                                    {
                                        if (model_safe != null)
                                            body_model.Add(model_safe); 
                                        model_safe = new TCMBodyForechModel();
                                        model_safe.body = new TSMBodyModel();
                                        model_safe.dict = new Dictionary<string, List<string>>();
                                        model_safe.body.Number  = ($"{myWorksheet.Cells[rowNum, IndexNumberPP + 1].Value?.ToString()}");
                                        model_safe.body.Cipher = ($"{myWorksheet.Cells[rowNum, IndexCipher + 1].Value?.ToString()}");
                                        model_safe.body.WorkName = ($"{myWorksheet.Cells[rowNum, IndexWorkName + 1].Value?.ToString()}");
                                        model_safe.body.Measurement =  ($"{myWorksheet.Cells[rowNum, IndexMeasurement + 1].Value?.ToString()}");
                                        model_safe.body.Quanity =  ($"{myWorksheet.Cells[rowNum, IndexQuanity + 1].Value?.ToString()}");
                                        if ((myWorksheet.Cells[rowNum, IndexUnitPrice + 1].Value?.ToString() == null) || (myWorksheet.Cells[rowNum, IndexUnitPrice + 1].Value?.ToString() == ""))
                                        {
                                            continue;
                                        }
                                    }
                                    if ((myWorksheet.Cells[rowNum, IndexBasicPrice + 1].Value?.ToString() != null) || (myWorksheet.Cells[rowNum, InexCorrection + 1].Value?.ToString() != null) || (myWorksheet.Cells[rowNum, IndexTotalPrice + 1].Value?.ToString() != null))
                                    {
                                        //model_safe.dict = new Dictionary<string, List<string>>();
                                        //var dict = new Dictionary<string, List<string>>();
                                        
                                        string name_rezldel = "";
                                        var list = new  List<string>();
                                        if ((myWorksheet.Cells[rowNum, IndexWorkName + 1].Value?.ToString() != null) && (myWorksheet.Cells[rowNum, IndexWorkName + 1].Value?.ToString() != ""))
                                        {
                                            name_rezldel = ($"{myWorksheet.Cells[rowNum, IndexWorkName + 1].Value?.ToString()}");
                                        }
                                        for (int i = 0; i < 6; i++)
                                        {
                                            if (myWorksheet.Cells[rowNum, indexd[i]].Value?.ToString() == null || myWorksheet.Cells[rowNum, indexd[i]].Value?.ToString() == "")
                                            {
                                                list.Add($"-\t");
                                            }
                                            else
                                            {
                                                list.Add($"{myWorksheet.Cells[rowNum, indexd[i]].Value?.ToString()}\t");
                                            }
                                        }
                                        model_safe.dict.Add(name_rezldel, list);
                                    }

                                }
                            }
                        }
                    }
                    //sw.Close();
                    //st.Close();
                    //se.Close();
                    CreateXLSX(filePath, stm, body_model, ending_model);
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void CreateXLSX(string filePath, TSMDictinaryModel stm, List<TCMBodyForechModel> tbm, TSMEndingModel etm)
        {

            using (ExcelPackage excel = new ExcelPackage())
            {
                var file_name = Path.GetFileName(filePath);
                var now = DateTime.Now;
                FileInfo excelFile = new FileInfo(@$"C:\cwc\ContractorsWorkAPI\ParsingFiles\{file_name}_{now.ToString("MM_dd_yyyy_HH_mm_ss")}.xlsx");
                ExcelWorksheet worksheet =  excel.Workbook.Worksheets.Add("Worksheet1");
                var index = 2;

                worksheet.Cells[1, 1].Value = "N";
                worksheet.Cells[1, 2].Value = "N п/п";
                worksheet.Cells[1, 3].Value = "Шифр,номера нормативов и коды ресурсов";
                worksheet.Cells[1, 4].Value = "Тип сборника";
                worksheet.Cells[1, 5].Value = "Дополнение";
                worksheet.Cells[1, 6].Value = "Номер сборника коэффицентов";
                worksheet.Cells[1, 7].Value = "Период сборника коэффицентов";
                //worksheet.Cells[1, 8].Value = "Раздел";
                //worksheet.Cells[1, 9].Value = "Подраздел";
                worksheet.Cells[1, 10].Value = "Наименование работ и затрат";
                worksheet.Cells[1, 11].Value = "Ед. изм";
                worksheet.Cells[1, 12].Value = "Кол-во едениц";
                worksheet.Cells[1, 13].Value = "Статья затрат";
                worksheet.Cells[1, 14].Value = "Ценаз за еденицу измерения, руб";
                worksheet.Cells[1, 15].Value = "Поправочный коэффицент";
                worksheet.Cells[1, 16].Value = "Коэффиицент зимний удорожаний";
                worksheet.Cells[1, 17].Value = "Затраты в базисонм уровне цен, руб";
                worksheet.Cells[1, 18].Value = "Коэффиценты (индексы) пересчета, нормы НР и СП";
                worksheet.Cells[1, 19].Value = "Всего затрат текущем уровне цен, руб";
                worksheet.Cells[1, 20].Value = "Итого всего затрат по наименованию работ";
                worksheet.Cells[1, 21].Value = "Итого по подразделу, руб";
                worksheet.Cells[1, 22].Value = "Итого по разделу, руб";
                worksheet.Cells[1, 23].Value = "Итого по всем разделам, руб";
                worksheet.Cells[1, 24].Value = "НДС, руб";
                worksheet.Cells[1, 25].Value = "Всего, руб";
                worksheet.Cells[1, 26].Value = "Итоговая сумма с коэффицентам финансирования, руб";
                var count = 1;

                foreach (var item in tbm)
                {
                    //foreach (var val in item.dict)
                    //{
                        worksheet.Cells[1, 1].Value = count;
                        worksheet.Cells[index, 2].Value = item.body.Number;
                        worksheet.Cells[index, 3].Value = item.body.Cipher;
                        worksheet.Cells[index, 4].Value = stm.Compound;
                        worksheet.Cells[index, 5].Value = stm.Addition;
                        worksheet.Cells[index, 6].Value = stm.Conversion_Index;
                        worksheet.Cells[index, 7].Value = stm.Compilation_Date;
                        //worksheet.Cells[index, 8].Value = item.body.;
                        worksheet.Cells[index, 10].Value = item.body.WorkName;
                        worksheet.Cells[index, 11].Value = item.body.Quanity;
                        worksheet.Cells[index, 12].Value = item.body.Measurement;
                        worksheet.Cells[index, 13].Value = item.dict.Keys;
                        worksheet.Cells[index, 14].Value = item.dict.Values;

                        index++;
                        count++;
                    //}
                }
                excel.SaveAs(excelFile);
            }
        }

        public static string NormalizeWhiteSpaceForLoop(string input)
        {
            int len = input.Length,
                index = 0,
                i = 0;
            var src = input.ToCharArray();
            bool skip = false;
            char ch;
            for (; i < len; i++)
            {
                ch = src[i];
                switch (ch)
                {
                    case '\u0020':
                    case '\u00A0':
                    case '\u1680':
                    case '\u2000':
                    case '\u2001':
                    case '\u2002':
                    case '\u2003':
                    case '\u2004':
                    case '\u2005':
                    case '\u2006':
                    case '\u2007':
                    case '\u2008':
                    case '\u2009':
                    case '\u200A':
                    case '\u202F':
                    case '\u205F':
                    case '\u3000':
                    case '\u2028':
                    case '\u2029':
                    case '\u0009':
                    case '\u000A':
                    case '\u000B':
                    case '\u000C':
                    case '\u000D':
                    case '\u0085':
                        if (skip) continue;
                        src[index++] = ch;
                        skip = true;
                        continue;
                    default:
                        skip = false;
                        src[index++] = ch;
                        continue;
                }
            }

            return new string(src, 0, index);
        }

    }
}
