using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace GenerateModel.Helpers
{
    public class ExcelHelper
    {
        /// <summary>
        /// 创建Excel文件
        /// </summary>
        /// <param name="Datas">需要写入到Excel文件中的数据集合</param>
        /// <param name="FilePath">Excel文件的物理路径</param>
        public static void Export(DataSet Datas, string FilePath)
        {
            try
            {
                FileInfo TargetFile = new FileInfo(FilePath);
                if (TargetFile.Extension != ".xlsx")
                {
                    throw new Exception("导出的文件类型只能是xlsx格式！");
                }
                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                }
                // 依据指定的物理地址创建Excel文件
                using (ExcelPackage Excel_Package = new ExcelPackage(TargetFile))
                {
                    // 遍历DataSet，依次读取其中的DataTable
                    foreach (DataTable Table in Datas.Tables)
                    {
                        // 根据读取到的DataTable在创建的Excel文件中建立新的Sheet
                        using (ExcelWorksheet Excel_Sheet = Excel_Package.Workbook.Worksheets.Add(Table.TableName))
                        {
                            // 遍历DataTable的列来创建Sheet里面的标题
                            int i = 1;
                            foreach (DataColumn Column in Table.Columns)
                            {
                                Excel_Sheet.Column(i).Width = 25.71;
                                Excel_Sheet.Column(i).Style.Numberformat.Format = "@";
                                using (ExcelRange Excel_Range = Excel_Sheet.Cells[1, i])
                                {
                                    Excel_Range.Style.WrapText = true;
                                    Excel_Range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    Excel_Range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    Excel_Range.Value = Column.ColumnName;
                                }
                                i++;
                            }
                            // 遍历DataTable的行来创建Sheet里面的内容
                            int j = 2;
                            foreach (DataRow Row in Table.Rows)
                            {
                                for (int k = 0; k < Table.Columns.Count; k++)
                                {
                                    using (ExcelRange Excel_Range = Excel_Sheet.Cells[j, k + 1])
                                    {
                                        Excel_Range.Style.WrapText = true;
                                        Excel_Range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                        Excel_Range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        Excel_Range.Value = Row[k];
                                    }
                                }
                                j++;
                            }
                            // 冻结Excel的第一行标题栏，使其不受到滚动条的滚动而消失
                            Excel_Sheet.View.FreezePanes(2, 1);
                            // 当前读取到的DataTable里面的数据处理完毕后，保存Excel文件
                            Excel_Package.Save();
                        }
                    }
                }
            }
            catch (Exception EX)
            {
                Console.WriteLine(EX);
            }
        }

        /// <summary>
        /// 读取Excel文件的内容
        /// </summary>
        /// <param name="FilePath">Excel文件的物理路径</param>
        /// <param name="excludeSheets">需要排除的Sheet集合</param>
        /// <returns></returns>
        public static DataSet Import(string FilePath, List<string> excludeSheets)
        {
            FileInfo TargetFile = new FileInfo(FilePath);
            DataSet Datas = new DataSet();
            try
            {
                if (TargetFile.Extension != ".xlsx")
                {
                    throw new Exception("导入的文件类型只能是xlsx格式！");
                }
                // 依据指定的物理地址读取Excel文件
                using (ExcelPackage Excel_Package = new ExcelPackage(TargetFile))
                {
                    // 遍历Excel文件的Sheet，依次读取其中的内容
                    foreach (ExcelWorksheet Excel_Sheet in Excel_Package.Workbook.Worksheets)
                    {
                        // 判断Sheet是否不需要被导入处理
                        if (excludeSheets.Contains(Excel_Sheet.Name)) continue;
                        // 创建DataTable用来储存Sheet里面的数据
                        using (DataTable Table = new DataTable() { TableName = Excel_Sheet.Name })
                        {
                            // 读取Sheet里面数据的标题
                            for (int i = Excel_Sheet.Dimension.Start.Column; i <= Excel_Sheet.Dimension.End.Column; i++)
                            {
                                using (ExcelRange Excel_Range = Excel_Sheet.Cells[1, i])
                                {
                                    Table.Columns.Add(Excel_Range.LocalAddress);
                                }
                            }
                            // 读取Sheet里面数据的内容
                            for (int i = Excel_Sheet.Dimension.Start.Row; i <= Excel_Sheet.Dimension.End.Row; i++)
                            {
                                Table.Rows.Add();
                                for (int j = Excel_Sheet.Dimension.Start.Column; j <= Excel_Sheet.Dimension.End.Column; j++)
                                {
                                    using (ExcelRange Excel_Range = Excel_Sheet.Cells[i, j])
                                    {
                                        Table.Rows[i - 1][j - 1] = Excel_Range.Value;
                                    }
                                }
                            }
                            // 读取完毕后将DataTable添加到DataSet里面去
                            Datas.Tables.Add(Table);
                        }
                    }
                }
            }
            catch (Exception EX)
            {
                Console.WriteLine(EX);
            }
            return Datas;
        }
    }
}