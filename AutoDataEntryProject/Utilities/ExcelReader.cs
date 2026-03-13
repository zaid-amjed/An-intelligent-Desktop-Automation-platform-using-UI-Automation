using AutoDataEntryProject.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoDataEntryProject.Utilities
{
    public class ExcelReader
    {
        public List<Student> ReadStudentFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("مسار الملف غير صحيح", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("الملف غير موجود", filePath);

            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xlsm")
                throw new ArgumentException("نوع الملف غير مدعوم. يجب أن يكون ملف Excel (.xlsx أو .xlsm)", nameof(filePath));

            List<Student> students = new List<Student>();

            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
                {
                    if (doc.WorkbookPart == null)
                        throw new InvalidOperationException("ملف Excel تالف أو غير صحيح");

                    WorkbookPart workbookPart = doc.WorkbookPart;

                    WorksheetPart? worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    if (worksheetPart == null)
                        throw new InvalidOperationException("لا توجد أوراق عمل في ملف Excel");

                    SheetData? sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                    if (sheetData == null)
                        throw new InvalidOperationException("ورقة العمل فارغة");

                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count <= 1)
                        throw new InvalidOperationException("لا توجد بيانات في ملف Excel (يوجد فقط صف العناوين أو أقل)");

                    SharedStringTablePart? sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    int rowNumber = 1;
                    foreach (Row row in rows.Skip(1))
                    {
                        rowNumber++;
                        try
                        {
                            var cells = row.Elements<Cell>().ToList();

                            if (cells.Count < 3)
                            {
                                System.Diagnostics.Debug.WriteLine($"تحذير: الصف {rowNumber} يحتوي على أقل من 3 أعمدة، سيتم تجاهله");
                                continue;
                            }

                            string name = GetCellValue(doc, cells[0], sstPart);
                            string id = GetCellValue(doc, cells[1], sstPart);
                            string email = GetCellValue(doc, cells[2], sstPart);

                            if (string.IsNullOrWhiteSpace(name) && 
                                string.IsNullOrWhiteSpace(id) && 
                                string.IsNullOrWhiteSpace(email))
                            {
                                continue;
                            }

                            Student student = new Student(
                                name?.Trim() ?? string.Empty,
                                id?.Trim() ?? string.Empty,
                                email?.Trim() ?? string.Empty
                            );

                            students.Add(student);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"خطأ في معالجة الصف {rowNumber}: {ex.Message}");
                        }
                    }
                }

                if (students.Count == 0)
                    throw new InvalidOperationException("لم يتم العثور على بيانات صحيحة في ملف Excel");

                return students;
            }
            catch (IOException ex)
            {
                throw new IOException($"خطأ في قراءة الملف: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"خطأ في معالجة ملف Excel: {ex.Message}", ex);
            }
        }

        private string GetCellValue(SpreadsheetDocument doc, Cell cell, SharedStringTablePart? sstPart)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sstPart != null && sstPart.SharedStringTable != null)
                {
                    var stringTable = sstPart.SharedStringTable.ChildElements;
                    if (int.TryParse(value, out int index) && index < stringTable.Count)
                    {
                        return stringTable[index].InnerText;
                    }
                }
                return string.Empty;
            }

            return value;
        }

        public bool ValidateExcelStructure(string filePath)
        {
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = doc.WorkbookPart;
                    if (workbookPart == null) return false;

                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    if (worksheetPart == null) return false;

                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                    if (sheetData == null) return false;

                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count <= 1) return false;

                    var firstDataRow = rows.Skip(1).FirstOrDefault();
                    if (firstDataRow == null) return false;

                    var cells = firstDataRow.Elements<Cell>().ToList();
                    return cells.Count >= 3;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
