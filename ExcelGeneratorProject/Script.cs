using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

class Program
{
    static void Main()
    {
        try
        {
            string peridosFilePath = Path.Combine("..", "PEDIDOS.xlsx");
            string outputFilePath = "../Pedido Distro Leganes " + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx";

            if (!File.Exists(peridosFilePath))
            {
                Console.WriteLine($"Error: No se encontró el archivo {peridosFilePath}");
                return;
            }

            // Leer el archivo PEDIDOS
            using (FileStream fs = new FileStream(peridosFilePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(fs);

                // Obtener las páginas necesarias
                ISheet datosSheet = workbook.GetSheet("DATOS") ?? workbook.GetSheetAt(0);
                ISheet plantillaSheet = workbook.GetSheet("PLANTILLA");
                ISheet ecigSheet = workbook.GetSheet("ECIGLOGISTICA");
                ISheet vaperaliaSheet = workbook.GetSheet("VAPERALIA");
                ISheet oil4vapSheet = workbook.GetSheet("OIL4VAP");

                if (datosSheet == null)
                {
                    Console.WriteLine("Error: No se encontró la página DATOS");
                    return;
                }

                if (plantillaSheet == null)
                {
                    Console.WriteLine("Error: No se encontró la página PLANTILLA");
                    return;
                }

                // Obtener datos de DATOS
                var datosData = GetSheetData(datosSheet);

                // Obtener datos de distribuidoras
                var ecigData = GetSheetData(ecigSheet);
                var vaperaliaData = GetSheetData(vaperaliaSheet);
                var oil4vapData = GetSheetData(oil4vapSheet);

                // Crear nuevo workbook para salida
                XSSFWorkbook outputWorkbook = new XSSFWorkbook();

                // Crear página de coincidencias
                ISheet outputSheet = outputWorkbook.CreateSheet("Coincidencias");
                
                // Copiar explícitamente los encabezados de la PLANTILLA a la fila 0
                IRow plantillaHeaderRow = plantillaSheet.GetRow(0);
                if (plantillaHeaderRow != null)
                {
                    IRow outputHeaderRow = outputSheet.CreateRow(0);
                    for (int colIdx = 0; colIdx < plantillaHeaderRow.LastCellNum; colIdx++)
                    {
                        ICell srcCell = plantillaHeaderRow.GetCell(colIdx);
                        ICell tgtCell = outputHeaderRow.CreateCell(colIdx);
                        if (srcCell != null)
                        {
                            tgtCell.SetCellValue(srcCell.StringCellValue);
                            
                            // Clonar el estilo del workbook fuente al workbook de destino
                            ICellStyle srcStyle = srcCell.CellStyle;
                            if (srcStyle != null)
                            {
                                ICellStyle tgtStyle = outputWorkbook.CreateCellStyle();
                                
                                // Copiar propiedades del estilo
                                tgtStyle.Alignment = srcStyle.Alignment;
                                tgtStyle.VerticalAlignment = srcStyle.VerticalAlignment;
                                tgtStyle.DataFormat = srcStyle.DataFormat;
                                tgtStyle.FillPattern = srcStyle.FillPattern;
                                tgtStyle.FillForegroundColor = srcStyle.FillForegroundColor;
                                tgtStyle.FillBackgroundColor = srcStyle.FillBackgroundColor;
                                tgtStyle.BorderBottom = srcStyle.BorderBottom;
                                tgtStyle.BorderLeft = srcStyle.BorderLeft;
                                tgtStyle.BorderRight = srcStyle.BorderRight;
                                tgtStyle.BorderTop = srcStyle.BorderTop;
                                
                                // Copiar fuente
                                try
                                {
                                    short srcFontIdx = srcStyle.FontIndex;
                                    IFont srcFont = workbook.GetFontAt(srcFontIdx);
                                    IFont tgtFont = outputWorkbook.CreateFont();
                                    tgtFont.IsBold = srcFont.IsBold;
                                    tgtFont.FontHeightInPoints = srcFont.FontHeightInPoints;
                                    tgtFont.FontName = srcFont.FontName;
                                    tgtFont.Color = srcFont.Color;
                                    tgtFont.IsItalic = srcFont.IsItalic;
                                    tgtFont.IsStrikeout = srcFont.IsStrikeout;
                                    tgtFont.Underline = srcFont.Underline;
                                    tgtStyle.SetFont(tgtFont);
                                }
                                catch
                                {
                                    // Ignorar si no se puede copiar la fuente
                                }
                                
                                tgtCell.CellStyle = tgtStyle;
                            }
                        }
                    }
                    // Copiar anchos de columna
                    for (int i = 0; i < plantillaHeaderRow.LastCellNum; i++)
                    {
                        outputSheet.SetColumnWidth(i, plantillaSheet.GetColumnWidth(i));
                    }
                }

                // Procesar coincidencias
                List<int> processedIndices = new List<int>();
                int outputRow = 1; // Comienza en fila 1 (fila 0 contiene encabezados)
                string todayDate = DateTime.Now.ToString("dd/MM/yyyy");

                foreach (var datosItem in datosData)
                {
                    if (datosItem.ColumnA == null || datosItem.ColumnA.Trim() == "")
                        continue;

                    List<MatchResult> matches = new List<MatchResult>();

                    // Buscar en ECIGLOGISTICA
                    if (ecigData != null)
                    {
                        var ecigMatch = ecigData.FirstOrDefault(x => x.ColumnA != null && 
                            x.ColumnA.Trim().Equals(datosItem.ColumnA.Trim(), StringComparison.OrdinalIgnoreCase));
                        if (ecigMatch != null)
                        {
                            matches.Add(new MatchResult 
                            { 
                                Source = "ECIGLOGISTICA", 
                                Code = 3, 
                                Abbreviation = "ECIG", 
                                ColumnB = ecigMatch.ColumnB,
                                ColumnE = ecigMatch.ColumnE
                            });
                        }
                    }

                    // Buscar en VAPERALIA
                    if (vaperaliaData != null)
                    {
                        var vaperaliaMatch = vaperaliaData.FirstOrDefault(x => x.ColumnA != null && 
                            x.ColumnA.Trim().Equals(datosItem.ColumnA.Trim(), StringComparison.OrdinalIgnoreCase));
                        if (vaperaliaMatch != null)
                        {
                            matches.Add(new MatchResult 
                            { 
                                Source = "VAPERALIA", 
                                Code = 2, 
                                Abbreviation = "VAP", 
                                ColumnB = vaperaliaMatch.ColumnB,
                                ColumnE = vaperaliaMatch.ColumnE
                            });
                        }
                    }

                    // Buscar en OIL4VAP
                    if (oil4vapData != null)
                    {
                        var oil4vapMatch = oil4vapData.FirstOrDefault(x => x.ColumnA != null && 
                            x.ColumnA.Trim().Equals(datosItem.ColumnA.Trim(), StringComparison.OrdinalIgnoreCase));
                        if (oil4vapMatch != null)
                        {
                            matches.Add(new MatchResult 
                            { 
                                Source = "OIL4VAP", 
                                Code = 7, 
                                Abbreviation = "OIL", 
                                ColumnB = oil4vapMatch.ColumnB,
                                ColumnE = oil4vapMatch.ColumnE
                            });
                        }
                    }

                    // Si hay coincidencias, agregarlas
                    if (matches.Count > 0)
                    {
                        processedIndices.Add(datosData.IndexOf(datosItem));

                        // Columna A: Fecha sin hora
                        SetCellValue(outputSheet, outputRow, 0, todayDate);

                        // Columna B: Códigos separados por " - "
                        string codesStr = string.Join(" - ", matches.Select(m => m.Code.ToString()));
                        SetCellValue(outputSheet, outputRow, 1, codesStr);

                        // Columna E: Columna B de la distribuidora (formato numérico sin decimales)
                        string columnEValues = string.Join(" - ", matches.Select(m => m.ColumnB ?? ""));
                        SetCellValueNumeric(outputSheet, outputRow, 4, columnEValues);

                        // Columna F: Lógica especial según coincidencias
                        string columnFValue;
                        if (matches.Count == 1)
                        {
                            // Una sola coincidencia
                            var match = matches[0];
                            if (!string.IsNullOrEmpty(match.ColumnE) && match.ColumnE.Trim() != "")
                            {
                                // ColumnE no está vacía: multiplicar
                                if (double.TryParse(match.ColumnE, out double columnEValue) && 
                                    double.TryParse(datosItem.ColumnB, out double datosColumnBValue))
                                {
                                    columnFValue = (columnEValue * datosColumnBValue).ToString("0");
                                }
                                else
                                {
                                    columnFValue = datosItem.ColumnB ?? "";
                                }
                            }
                            else
                            {
                                // ColumnE está vacía: solo asignar ColumnB de DATOS
                                columnFValue = datosItem.ColumnB ?? "";
                            }
                        }
                        else
                        {
                            // Múltiples coincidencias: solo asignar ColumnB de DATOS
                            columnFValue = datosItem.ColumnB ?? "";
                        }
                        SetCellValue(outputSheet, outputRow, 5, columnFValue);

                        // Columna G: Abreviaturas separadas por " - "
                        string abbreviationsStr = string.Join(" - ", matches.Select(m => m.Abbreviation));
                        SetCellValue(outputSheet, outputRow, 6, abbreviationsStr);

                        // Columna H: "Leganés"
                        SetCellValue(outputSheet, outputRow, 7, "Leganés");

                        outputRow++;
                    }
                }

                // Crear segunda página para no coincidencias
                if (datosData.Count > processedIndices.Count)
                {
                    ISheet noMatchSheet = outputWorkbook.CreateSheet("NoCoincidencias");
                    
                    // Agregar encabezados
                    IRow headerRow = noMatchSheet.CreateRow(0);
                    headerRow.CreateCell(0).SetCellValue("Columna A");
                    headerRow.CreateCell(1).SetCellValue("Columna B");

                    int noMatchRow = 1;
                    for (int i = 0; i < datosData.Count; i++)
                    {
                        if (!processedIndices.Contains(i))
                        {
                            var item = datosData[i];
                            IRow row = noMatchSheet.CreateRow(noMatchRow);
                            row.CreateCell(0).SetCellValue(item.ColumnA ?? "");
                            row.CreateCell(1).SetCellValue(item.ColumnB ?? "");
                            noMatchRow++;
                        }
                    }
                }

                // Guardar el archivo de salida
                using (FileStream outputFs = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    outputWorkbook.Write(outputFs);
                }

                Console.WriteLine($"Proceso completado. Archivo guardado: {outputFilePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine($"StackTrace: {ex.StackTrace}");
        }
    }

    static List<SheetData> GetSheetData(ISheet sheet)
    {
        List<SheetData> data = new List<SheetData>();

        if (sheet == null)
            return data;

        for (int i = 0; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            if (row == null)
                continue;

            string columnA = GetCellStringValue(row, 0);
            string columnB = GetCellStringValue(row, 1);
            string columnE = GetCellStringValue(row, 4);

            if (!string.IsNullOrEmpty(columnA))
            {
                data.Add(new SheetData 
                { 
                    ColumnA = columnA.Trim(), 
                    ColumnB = columnB,
                    ColumnE = columnE
                });
            }
        }

        return data;
    }

    static string GetCellStringValue(IRow row, int columnIndex)
    {
        ICell cell = row.GetCell(columnIndex);
        if (cell == null)
            return "";

        switch (cell.CellType)
        {
            case CellType.String:
                return cell.StringCellValue;
            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(cell))
                    return string.Format("{0:dd/MM/yyyy}", cell.DateCellValue);
                else
                    return cell.NumericCellValue.ToString();
            case CellType.Boolean:
                return cell.BooleanCellValue.ToString();
            case CellType.Formula:
                return cell.CellFormula;
            default:
                return "";
        }
    }

    static void CopySheetStructure(ISheet sourceSheet, ISheet targetSheet)
    {
        // Copiar encabezados (primera fila)
        if (sourceSheet.LastRowNum >= 0)
        {
            IRow sourceHeaderRow = sourceSheet.GetRow(0);
            if (sourceHeaderRow != null)
            {
                IRow targetHeaderRow = targetSheet.CreateRow(0);
                for (int i = 0; i < sourceHeaderRow.LastCellNum; i++)
                {
                    ICell sourceCell = sourceHeaderRow.GetCell(i);
                    ICell targetCell = targetHeaderRow.CreateCell(i);
                    if (sourceCell != null)
                    {
                        targetCell.SetCellValue(sourceCell.StringCellValue);
                        // Clonar estilo básico al workbook de destino en lugar de asignar el estilo directamente
                        ICellStyle srcStyle = sourceCell.CellStyle;
                        if (srcStyle != null)
                        {
                            IWorkbook targetWb = targetSheet.Workbook;
                            IWorkbook sourceWb = sourceSheet.Workbook;
                            ICellStyle tgtStyle = targetWb.CreateCellStyle();

                            // Copiar propiedades comunes
                            tgtStyle.Alignment = srcStyle.Alignment;
                            tgtStyle.VerticalAlignment = srcStyle.VerticalAlignment;
                            tgtStyle.DataFormat = srcStyle.DataFormat;
                            tgtStyle.FillPattern = srcStyle.FillPattern;
                            tgtStyle.FillForegroundColor = srcStyle.FillForegroundColor;
                            tgtStyle.FillBackgroundColor = srcStyle.FillBackgroundColor;
                            tgtStyle.BorderBottom = srcStyle.BorderBottom;
                            tgtStyle.BorderLeft = srcStyle.BorderLeft;
                            tgtStyle.BorderRight = srcStyle.BorderRight;
                            tgtStyle.BorderTop = srcStyle.BorderTop;

                            // Copiar fuente
                            try
                            {
                                short srcFontIdx = srcStyle.FontIndex;
                                IFont srcFont = sourceWb.GetFontAt(srcFontIdx);
                                IFont tgtFont = targetWb.CreateFont();
                                try
                                {
                                    tgtFont.IsBold = srcFont.IsBold;
                                }
                                catch
                                {
                                    // Fallback por compatibilidad si IsBold no existe
                                    // Ignorar y continuar
                                }
                                tgtFont.FontHeightInPoints = srcFont.FontHeightInPoints;
                                tgtFont.FontName = srcFont.FontName;
                                tgtFont.Color = srcFont.Color;
                                tgtFont.IsItalic = srcFont.IsItalic;
                                tgtFont.IsStrikeout = srcFont.IsStrikeout;
                                tgtFont.Underline = srcFont.Underline;
                                tgtStyle.SetFont(tgtFont);
                            }
                            catch
                            {
                                // Ignorar si no se puede copiar la fuente; no es crítico
                            }

                            targetCell.CellStyle = tgtStyle;
                        }
                    }
                }
            }
        }

        // Copiar anchos de columna
        for (int i = 0; i < sourceSheet.GetRow(0)?.LastCellNum; i++)
        {
            targetSheet.SetColumnWidth(i, sourceSheet.GetColumnWidth(i));
        }
    }

    static void SetCellValue(ISheet sheet, int row, int column, string value)
    {
        IRow sheetRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
        ICell cell = sheetRow.CreateCell(column);
        cell.SetCellValue(value ?? "");
    }

    static void SetCellValueNumeric(ISheet sheet, int row, int column, string value)
    {
        IRow sheetRow = sheet.GetRow(row) ?? sheet.CreateRow(row);
        ICell cell = sheetRow.CreateCell(column);
        
        // Eliminar separadores " - " para formato numérico si es necesario
        string cleanValue = value.Replace(" - ", ";");
        cell.SetCellValue(cleanValue);
        
        // Establecer formato numérico sin decimales
        ICellStyle style = sheet.Workbook.CreateCellStyle();
        IDataFormat format = sheet.Workbook.CreateDataFormat();
        style.DataFormat = format.GetFormat("0");
        cell.CellStyle = style;
    }
}

class SheetData
{
    public string? ColumnA { get; set; }
    public string? ColumnB { get; set; }
    public string? ColumnE { get; set; }
}

class MatchResult
{
    public string? Source { get; set; }
    public int Code { get; set; }
    public string? Abbreviation { get; set; }
    public string? ColumnB { get; set; }
    public string? ColumnE { get; set; }
}
