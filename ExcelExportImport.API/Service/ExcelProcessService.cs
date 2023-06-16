using ClosedXML.Excel;

namespace ExcelExportImport.API.Service
{
    public class ExcelProcessService
    {

        /// <summary>
        /// Verilen listeyi excel çıktısına dönüştüren mettotur.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        public static Stream ExportToExcel<T>(List<T> data)
        {
            // Excel çalışma kitabını oluştur
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Başlık satırını ekleyin ve stilleri uygular
                int columnIndex = 1;
                var headerRow = worksheet.Row(1);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
                foreach (var property in typeof(T).GetProperties())
                {
                    worksheet.Cell(1, columnIndex).Value = property.Name;
                    columnIndex++;
                }

                // Verileri satırlara ekler
                int rowIndex = 2;
                foreach (var item in data)
                {
                    columnIndex = 1;
                    foreach (var property in typeof(T).GetProperties())
                    {
                        worksheet.Cell(rowIndex, columnIndex).Value = property.GetValue(item).ToString();
                        columnIndex++;
                    }
                    rowIndex++;
                }

                // Excel dosyasını bellek akışına kaydeder
                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    memoryStream.Position = 0;
                    return memoryStream;
                }
            }
        }
    }
}
