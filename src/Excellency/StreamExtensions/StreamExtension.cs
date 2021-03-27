using Excellency.Attributes;
using Excellency.DTOs;
using Excellency.Enums;
using Excellency.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;

namespace Excellency.StreamExtensions
{
    public static class StreamExtension
    {
        public static IEnumerable<T> DeserializeFromExcel<T>(this Stream excelFileStream, ExcelFileExtension excelFileExtension)
        {
            var sourceAttr = ((SheetSourceAttribute)typeof(T)
                .GetCustomAttribute(typeof(SheetSourceAttribute))).Source;

            return DeserializeFromExcel<T>(excelFileStream, excelFileExtension, new List<SheetMap> { new SheetMap(sourceAttr, typeof(T)) });
        }

        public static IEnumerable<T> DeserializeFromExcel<T>(this Stream excelFileStream, ExcelFileExtension excelFileExtension, IEnumerable<SheetMap> sheeMaps)
        {
            if (!sheeMaps.Any()) throw new PropertyNotMappedException("Failed: The map is not setted");

            using (excelFileStream)
            {
                excelFileStream.Position = 0;
                IWorkbook workBook = WorkbookFactory(excelFileExtension, excelFileStream);
                sheeMaps.ToList().ForEach(x =>
                {
                    x.Sheet = workBook.GetSheet(x.SheetName);
                });
                workBook.Close();
            }

            var objects = new List<object>();

            foreach (var sheet in sheeMaps)
            {
                var sheetProperties = sheet.Type.GetProperties().Where(prop => prop.GetCustomAttribute(typeof(CellSourceAttribute)) != null);

                var cellMaps = sheetProperties.Select(x => new CellMap
                {
                    Source = ((CellSourceAttribute)x.GetCustomAttribute(typeof(CellSourceAttribute))).Source,
                    Target = x.Name
                });

                var header = sheet.Sheet.GetRow(0).Cells.Where(x=> cellMaps.Any(y=>y.Source.ToLower() == x.StringCellValue.ToLower())).ToList();
                var indexes = header.Select(x => x.ColumnIndex);
                

                var assemblyName = new AssemblyName(sheet.Sheet.SheetName);
                var assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
                var moduleBuilder = assemblyBuilder.DefineDynamicModule($"{assemblyName.Name}.dll");
                var bulder = moduleBuilder.DefineType(sheet.Sheet.SheetName, TypeAttributes.Public);

                foreach (var cell in header)
                {
                    bulder.DefineField(cellMaps.First(c=>c.Source.ToLower() == cell.StringCellValue.ToLower()).Target, typeof(string), FieldAttributes.Public);
                }

                for (int row = 1; row <= sheet.Sheet.LastRowNum; row++)
                {
                    var currentRow = sheet.Sheet.GetRow(row);                  

                    if (currentRow != null)
                    {
                        if (currentRow.PhysicalNumberOfCells < header.Count)
                        {
                            sheet.Sheet.RemoveRow(currentRow);
                            continue;
                        }

                        var runTimeTargetObject = Activator.CreateInstance(sheet.Type);

                        foreach (var index in indexes)
                        {

                            var cellData = GetTypedValue(currentRow, index);
                            

                            if (cellData!=null)
                            {
                                try
                                {
                                    var target = runTimeTargetObject.GetType().GetProperty(runTimeTargetObject.GetType().GetProperty(GetCellMap(cellMaps, header.First(h=>h.ColumnIndex == index).StringCellValue).Target).Name);
                                    target.SetValue(runTimeTargetObject, Cast(target.PropertyType, cellData));
                                }
                                catch (Exception ex)
                                {
                                    throw new PropertyNotMappedException(ex.Message,ex);
                                }
                            }
                        }

                        objects.Add(runTimeTargetObject);
                    }
                }
            }
            return objects.Cast<T>().ToList();
        }

        private static IWorkbook WorkbookFactory(ExcelFileExtension excelFileExtension, Stream excelFileStream)
        {
            if (excelFileExtension == ExcelFileExtension.Xlsx)
            {
                return new XSSFWorkbook(excelFileStream);
            }
            else
            {
                return new HSSFWorkbook(excelFileStream);
            }

        }
        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1;
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }
        private static object GetTypedValue(IRow currentRow, int index)
        {
            var cell = currentRow.GetCell(index);
            var cellType = cell.CellType;

            if (cellType == CellType.Numeric && DateUtil.IsCellDateFormatted(currentRow.GetCell(index)))
            {
                return FromExcelSerialDate(Convert.ToInt32(currentRow.GetCell(index).NumericCellValue));
            }
            else if (cellType == CellType.Numeric)
            {
                return currentRow.GetCell(index).NumericCellValue;
            }
            else if (cellType == CellType.String)
            {
                return currentRow.GetCell(index).StringCellValue;
            }
            else if (cellType == CellType.Boolean)
            {
                return currentRow.GetCell(index).BooleanCellValue;
            }
            else if (cellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric)
            {
                return currentRow.GetCell(index).NumericCellValue;
            }
            else if (cellType == CellType.Blank)
            {
                return null;
            }
            else
            {
                return currentRow.GetCell(index).StringCellValue;
            }
        }
        private static CellMap GetCellMap(IEnumerable<CellMap> cellMaps, string source)
        {
            var map = cellMaps.First(cell => cell.Source.ToLower() == source.ToLower());
            return map;
        }
        private static object Cast(Type type, object transformedData)
        {
            if (type == typeof(DateTime) || type == typeof(DateTime?))
            {
                return Convert.ToDateTime(transformedData);
            }
            else if (type == typeof(int) || type == typeof(int?))
            {
                return Convert.ToInt32(transformedData);
            }
            else if (type == typeof(double) || type == typeof(double?))
            {
                return Convert.ToDouble(transformedData);
            }
            else if (type == typeof(decimal) || type == typeof(decimal?))
            {
                return Convert.ToDecimal(transformedData);
            }
            else if (type == typeof(string))
            {
                return Convert.ToString(transformedData);
            }
            else
            {
                return transformedData;
            }
        }
    }

}
