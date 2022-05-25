using GenerateModel.Helpers;
using GenerateModel.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace GenerateModel
{
    internal class Program
    {
        /// <summary>
        /// 文件路径列表
        /// </summary>
        private static List<string> FilePathList { get; set; }

        /// <summary>
        /// 需要排除导入的Sheet集合
        /// </summary>
        private static List<string> ExcludeSheets { get; set; }

        /// <summary>
        /// 是否使用自定义文件路径
        /// </summary>
        private static bool UseCustomFilePath { get; set; }

        /// <summary>
        /// 生成的文件的存放路径
        /// </summary>
        private static string OutPutDirectory { get; set; }

        static void Main(string[] args)
        {
            GetFilePath();
            switch (GetFile())
            {
                case "A":
                case "a":
                    Console.WriteLine("您选择了使用Excel来处理数据！");
                    HandleExcelData(FilePathList[0]);
                    break;
                case "B":
                case "b":
                    Console.WriteLine("您选择了使用Html来处理数据！");
                    HandleHtmlData(FilePathList[1]);
                    break;
                default:
                    Console.WriteLine("无效的输入，请重新输入！");
                    break;
            }
            Console.Read();
        }

        /// <summary>
        /// 获取模板路径
        /// </summary>
        /// <returns></returns>
        private static void GetFilePath()
        {
            try
            {
                FilePathList = new List<string>();
                using (StreamReader SR = new StreamReader(Environment.CurrentDirectory + "\\FilePath.json", Encoding.Default))
                {
                    using (JsonTextReader JsonReader = new JsonTextReader(SR))
                    {
                        var JsonObject = JToken.ReadFrom(JsonReader) as JObject;
                        if (JsonObject != null)
                        {
                            foreach (var item in JsonObject.Children())
                            {
                                switch (item.Path)
                                {
                                    case "CustomExcelFilePath":
                                        FilePathList.Add(JsonObject[item.Path].ToString());
                                        break;
                                    case "CustomHtmlFilePath":
                                        FilePathList.Add(JsonObject[item.Path].ToString());
                                        break;
                                    case "UseCustomFilePath":
                                        UseCustomFilePath = ParseHelper.ConvertToBool(JsonObject[item.Path].ToString());
                                        break;
                                    case "OutPutDirectory":
                                        OutPutDirectory = JsonObject[item.Path].ToString();
                                        break;
                                    case "ExcludeSheetName":
                                        ExcludeSheets = new List<string>();
                                        foreach (var subItem in JsonObject[item.Path].Values())
                                        {
                                            ExcludeSheets.Add(subItem.ToString());
                                        }
                                        break;
                                    default:
                                        Console.WriteLine("Json文件中存在无效的Key值");
                                        break;
                                }
                            }
                        }
                        else
                        {
                            // Do Nonthing
                        }
                    }
                }
            }
            catch (Exception EX)
            {
                FilePathList.Clear();
                Console.WriteLine(EX);
            }
            // 判断自定义文件路径是否获取成功
            if (FilePathList.Count != 2 || !UseCustomFilePath)
            {
                // 自定义导入文件路径无效或者不使用自定义文件路径
                Console.WriteLine("“FilePath.json”中的自定义文件路径获取失败或者定义了不使用自定义文件路径，将尝试去使用默认文件路径");
                FilePathList.Clear();
                FilePathList.Add(Environment.CurrentDirectory + "\\Template\\excelTemplate.xlsx");
                FilePathList.Add(Environment.CurrentDirectory + "\\Template\\htmlTemplate.html");
            }
            else
            {
                // 使用自定义文件路径并且自定义文件路径有效
            }
            // 判断自定义导出目录是否存在
            if (!UseCustomFilePath || !Directory.Exists(OutPutDirectory))
            {
                // 不使用自定义路径，或者自定义导出目录不存在
                OutPutDirectory = Environment.CurrentDirectory + "\\OutPut\\";
                Directory.CreateDirectory(OutPutDirectory);

            }
            else
            {
                // 使用自定义文件路径并且自定义文件夹存在
            }
        }

        /// <summary>
        /// 获取用户的选择
        /// </summary>
        /// <returns></returns>
        private static string GetFile()
        {
            Console.WriteLine("请选择数据源格式：");
            Console.WriteLine("A：Excel");
            Console.WriteLine("B：Html");
            Console.Write("您的选择是：");
            return Console.ReadLine();
        }

        /// <summary>
        /// 从Excel中加载数据
        /// </summary>
        /// <param name="filePath"></param>
        private static void HandleExcelData(string filePath)
        {
            // 判断文件物理路径是否存在
            if (!File.Exists(filePath))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("此Excel路径“" + filePath + "”无效，将终止读取操作");
                return;
            }
            // 判断导出目录是否存在
            if (!Directory.Exists(OutPutDirectory))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("此导出路径“" + OutPutDirectory + "”无效，将终止读取操作");
                return;
            }
            // 从Excel中读取内容
            using (DataSet Datas = ExcelHelper.Import(filePath, ExcludeSheets))
            {
                // 判断Excel文件的内容是否为空
                if (Datas.Tables == null || Datas.Tables.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("导入的Excel中没有有效数据，请检查");
                    return;
                }
                Console.WriteLine("一共导入了“" + Datas.Tables.Count + "”个Sheet页内容");
                Console.WriteLine("开始处理所有导入的Excel的Sheet页的数据");
                // 便利导入的Excel的Sheet页数据
                foreach (DataTable dataTable in Datas.Tables)
                {
                    // 定义全量ExcelModel集合
                    var excelModels = new List<ExcelModel>();
                    try
                    {
                        // 转换DataTable到ExcelModel
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            // 创建Excel对象模型
                            excelModels.Add(new ExcelModel()
                            {
                                A1 = dataTable.Rows[i][0].ToString(),
                                B1 = dataTable.Rows[i][1].ToString(),
                                C1 = dataTable.Rows[i][2].ToString(),
                                D1 = dataTable.Rows[i][3].ToString(),
                                E1 = dataTable.Rows[i][4].ToString(),
                                F1 = dataTable.Rows[i][5].ToString(),
                                Index = i,
                                Children = new List<ExcelModel>(),
                            });
                        }
                        dataTable.Dispose();
                    }
                    catch (Exception EX)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("处理Excel的“" + dataTable.TableName + "”Sheet页数据时出现了异常，异常信息如下：");
                        Console.WriteLine(EX);
                        continue;
                    }
                    // 获取请求的URL行
                    var urlModels = excelModels.FindAll(x => string.IsNullOrEmpty(x.A1) && x.B1 == "请求URL" && !string.IsNullOrEmpty(x.C1) && string.IsNullOrEmpty(x.D1) && string.IsNullOrEmpty(x.E1) && string.IsNullOrEmpty(x.F1));
                    // 判断URL行数据是否有误
                    if (urlModels.Count != 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据中请求URL行数据内容有误，无法执行导出操作");
                        continue;
                    }
                    // 获取请求的备注行
                    var titleModels = excelModels.FindAll(x => string.IsNullOrEmpty(x.A1) && x.B1 == "接口名称" && !string.IsNullOrEmpty(x.C1) && string.IsNullOrEmpty(x.D1) && string.IsNullOrEmpty(x.E1) && string.IsNullOrEmpty(x.F1));
                    // 判断备注行数据是否有误
                    if (titleModels.Count != 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据中接口名称行数据内容有误，无法执行导出操作");
                        continue;
                    }
                    // 获取“请求参数列表”的标签行
                    var requestTagModels = excelModels.FindAll(x => x.A1 == "请求参数列表" && string.IsNullOrEmpty(x.B1) && string.IsNullOrEmpty(x.C1) && string.IsNullOrEmpty(x.D1) && string.IsNullOrEmpty(x.E1) && string.IsNullOrEmpty(x.F1));
                    // 判断标签行数据是否有误
                    if (requestTagModels.Count != 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据中“请求参数列表”标签行数据有误，无法执行导出操作");
                        continue;
                    }
                    // 获取“响应参数列表”的标签行
                    var responseTagModels = excelModels.FindAll(x => x.A1 == "响应参数列表" && string.IsNullOrEmpty(x.B1) && string.IsNullOrEmpty(x.C1) && string.IsNullOrEmpty(x.D1) && string.IsNullOrEmpty(x.E1) && string.IsNullOrEmpty(x.F1));
                    // 判断标签行数据是否有误
                    if (responseTagModels.Count != 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据中“响应参数列表”标签行数据有误，无法执行导出操作");
                        continue;
                    }
                    // 判断“请求参数列表”和“响应参数列表”的标签行的索引大小
                    if (requestTagModels[0].Index >= responseTagModels[0].Index)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        /* “请求参数列表”的标签行必定在“响应参数列表”的标签行前面，也就是requestTagModels的索引必定要小于responseTagModels的索引，否则就会有异常 */
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据的“请求参数列表”和“响应参数列表”顺序有颠倒，无法执行导出操作");
                        continue;
                    }
                    // 获取请求响应的表格表头
                    var headerModels = excelModels.FindAll(x => x.A1 == "序号" && x.B1 == "对应数据元" && x.C1 == "Json参数名" && x.D1 == "对应界面中文标识" && x.E1 == "参数类型" && x.F1 == "备注");
                    // 判断表头数量是否有误（一般为两个，默认第一个是“请求参数列表”的表头，第二个是“响应参数列表”的表头）
                    if (headerModels.Count != 2)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据中未能找到符合条件的请求区域和响应区域，无法执行导出操作");
                        continue;
                    }
                    // 判断“请求参数列表”和“响应参数列表”的标签行的位置和表头是否匹配
                    if (headerModels[0].Index != requestTagModels[0].Index + 1 || headerModels[1].Index != responseTagModels[0].Index + 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据的数据内容有误，无法执行导出操作");
                        continue;
                    }
                    // 获取请求响应的表格内容
                    var contentModels = excelModels.FindAll(x => !string.IsNullOrEmpty(x.B1) && NodeHelper.NodeList.Contains(x.B1));;
                    // 判断表格内容为空
                    if (contentModels.Count < 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据的“请求参数列表”和“响应参数列表”均为空，无需执行导出操作");
                        continue;
                    }
                    // 判断数据格式是否有误
                    if (contentModels[0].Index < headerModels[0].Index + 1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据的数据内容有误，无法执行导出操作");
                        continue;
                    }
                    // 筛选出“请求参数列表”和“响应参数列表”
                    var requestModels = new List<ExcelModel>();
                    var responseModels = new List<ExcelModel>();

                    #region 归类分别筛选出“请求参数列表”和“响应参数列表”

                    // 遍历表格内容，将表格数据分别汇总到“请求参数列表”和“响应参数列表”中
                    foreach (var model in contentModels)
                    {
                        // 判断Excel对象模型的层级
                        if (model.B1 == NodeHelper.FirstNode)
                        {
                            // Excel对象模型是第一层级，判断Excel对象模型的索引
                            if (model.Index <= headerModels[1].Index)
                            {
                                // 向“请求参数列表”中添加Excel对象模型
                                requestModels.Add(model);
                            }
                            else
                            {
                                // 向“响应参数列表”中添加Excel对象模型
                                responseModels.Add(model);
                            }
                        }
                        else if (model.B1 == NodeHelper.SecondNode)
                        {
                            // Excel对象模型是第二层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            model.Parent = firstParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            firstParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.ThirdNode)
                        {
                            // Excel对象模型是第三层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            model.Parent = secondParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            secondParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.FourthNode)
                        {
                            // Excel对象模型是第四层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            model.Parent = thirdParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            thirdParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.FifthNode)
                        {
                            // Excel对象模型是第五层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            // 在第三层级的子集中寻找其父对象（父对象必定存在，第三层级的子集中的最后一个对象就是其父对象）
                            var fourthParent = thirdParent.Children[thirdParent.Children.Count - 1];
                            model.Parent = fourthParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            fourthParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.SixthNode)
                        {
                            // Excel对象模型是第六层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            // 在第三层级的子集中寻找其父对象（父对象必定存在，第三层级的子集中的最后一个对象就是其父对象）
                            var fourthParent = thirdParent.Children[thirdParent.Children.Count - 1];
                            // 在第四层级的子集中寻找其父对象（父对象必定存在，第四层级的子集中的最后一个对象就是其父对象）
                            var fifthParent = fourthParent.Children[fourthParent.Children.Count - 1];
                            model.Parent = fifthParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            fifthParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.SeventhNode)
                        {
                            // Excel对象模型是第七层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            // 在第三层级的子集中寻找其父对象（父对象必定存在，第三层级的子集中的最后一个对象就是其父对象）
                            var fourthParent = thirdParent.Children[thirdParent.Children.Count - 1];
                            // 在第四层级的子集中寻找其父对象（父对象必定存在，第四层级的子集中的最后一个对象就是其父对象）
                            var fifthParent = fourthParent.Children[fourthParent.Children.Count - 1];
                            // 在第五层级的子集中寻找其父对象（父对象必定存在，第五层级的子集中的最后一个对象就是其父对象）
                            var sixthParent = fifthParent.Children[fifthParent.Children.Count - 1];
                            model.Parent = sixthParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            sixthParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.EighthNode)
                        {
                            // Excel对象模型是第八层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            // 在第三层级的子集中寻找其父对象（父对象必定存在，第三层级的子集中的最后一个对象就是其父对象）
                            var fourthParent = thirdParent.Children[thirdParent.Children.Count - 1];
                            // 在第四层级的子集中寻找其父对象（父对象必定存在，第四层级的子集中的最后一个对象就是其父对象）
                            var fifthParent = fourthParent.Children[fourthParent.Children.Count - 1];
                            // 在第五层级的子集中寻找其父对象（父对象必定存在，第五层级的子集中的最后一个对象就是其父对象）
                            var sixthParent = fifthParent.Children[fifthParent.Children.Count - 1];
                            // 在第六层级的子集中寻找其父对象（父对象必定存在，第六层级的子集中的最后一个对象就是其父对象）
                            var seventhParent = sixthParent.Children[sixthParent.Children.Count - 1];
                            model.Parent = seventhParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            seventhParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.NinthNode)
                        {
                            // Excel对象模型是第九层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            // 在第三层级的子集中寻找其父对象（父对象必定存在，第三层级的子集中的最后一个对象就是其父对象）
                            var fourthParent = thirdParent.Children[thirdParent.Children.Count - 1];
                            // 在第四层级的子集中寻找其父对象（父对象必定存在，第四层级的子集中的最后一个对象就是其父对象）
                            var fifthParent = fourthParent.Children[fourthParent.Children.Count - 1];
                            // 在第五层级的子集中寻找其父对象（父对象必定存在，第五层级的子集中的最后一个对象就是其父对象）
                            var sixthParent = fifthParent.Children[fifthParent.Children.Count - 1];
                            // 在第六层级的子集中寻找其父对象（父对象必定存在，第六层级的子集中的最后一个对象就是其父对象）
                            var seventhParent = sixthParent.Children[sixthParent.Children.Count - 1];
                            // 在第七层级的子集中寻找其父对象（父对象必定存在，第七层级的子集中的最后一个对象就是其父对象）
                            var eighthParent = seventhParent.Children[seventhParent.Children.Count - 1];
                            model.Parent = eighthParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            eighthParent.Children.Add(model);
                        }
                        else if (model.B1 == NodeHelper.TenthNode)
                        {
                            // Excel对象模型是第十层级
                            // 根据Excel对象模型的索引来决定从“请求参数列表”还是“响应参数列表”的集合中寻找其父对象(父对象必定存在，集合中的最后一个对象就是其父对象)
                            var firstParent = model.Index <= headerModels[1].Index ? requestModels[requestModels.Count - 1] : responseModels[responseModels.Count - 1];
                            // 在第一层级的子集中寻找其父对象（父对象必定存在，第一层级的子集中的最后一个对象就是其父对象）
                            var secondParent = firstParent.Children[firstParent.Children.Count - 1];
                            // 在第二层级的子集中寻找其父对象（父对象必定存在，第二层级的子集中的最后一个对象就是其父对象）
                            var thirdParent = secondParent.Children[secondParent.Children.Count - 1];
                            // 在第三层级的子集中寻找其父对象（父对象必定存在，第三层级的子集中的最后一个对象就是其父对象）
                            var fourthParent = thirdParent.Children[thirdParent.Children.Count - 1];
                            // 在第四层级的子集中寻找其父对象（父对象必定存在，第四层级的子集中的最后一个对象就是其父对象）
                            var fifthParent = fourthParent.Children[fourthParent.Children.Count - 1];
                            // 在第五层级的子集中寻找其父对象（父对象必定存在，第五层级的子集中的最后一个对象就是其父对象）
                            var sixthParent = fifthParent.Children[fifthParent.Children.Count - 1];
                            // 在第六层级的子集中寻找其父对象（父对象必定存在，第六层级的子集中的最后一个对象就是其父对象）
                            var seventhParent = sixthParent.Children[sixthParent.Children.Count - 1];
                            // 在第七层级的子集中寻找其父对象（父对象必定存在，第七层级的子集中的最后一个对象就是其父对象）
                            var eighthParent = seventhParent.Children[seventhParent.Children.Count - 1];
                            // 在第八层级的子集中寻找其父对象（父对象必定存在，第八层级的子集中的最后一个对象就是其父对象）
                            var ninthParent = eighthParent.Children[eighthParent.Children.Count - 1];
                            model.Parent = ninthParent;
                            // 将Excel对象模型添加到其父对象的子集中
                            ninthParent.Children.Add(model);
                        }
                        else
                        {
                            // 其余Excel对象模型层级暂未处理
                        }
                    }

                    #endregion

                    // 准备导出表格内容到ts文件中
                    // 获取文件名和url名称
                    var urlSplits = urlModels[0].C1.Split('/');
                    // 定义命名空间
                    var nameSpaces = urlSplits[urlSplits.Length - 1];
                    // 格式化命名空间(命名空间格式为所有的单词，首字母均要大写)
                    nameSpaces = string.Format("{0}{1}", nameSpaces.Substring(0, 1).ToUpper(), nameSpaces.Remove(0, 1));
                    // 创建内容构造流
                    StringBuilder stringBuilder = new StringBuilder();
                    // 向内容构造流中添加空白的import内容
                    stringBuilder.AppendLine(@"import { baseUrl } from '@env/enviroment';");
                    // 向内容构造流中追加空白行
                    stringBuilder.AppendLine();
                    // 向内容构造流中追加类注释内容
                    stringBuilder.AppendLine(@"/**");
                    stringBuilder.AppendLine(string.Format(@" * {0}", titleModels[0].C1));
                    stringBuilder.AppendLine(@" */");
                    // 向内容构造流中追加命名空间
                    stringBuilder.AppendLine(string.Format(@"export namespace {0}Model {1}", nameSpaces, "{"));
                    // 向内容构造流中追加常量定义
                    stringBuilder.AppendLine(string.Format(@"    export const URL_ADDRESS = baseUrl + '{0}';", urlModels[0].C1));
                    // 向内容构造流中追加空白行
                    stringBuilder.AppendLine();
                    // 向内容构造流中追加请求Model
                    stringBuilder.AppendLine(@"    // 入参类型");
                    stringBuilder.AppendLine(@"    export interface RequestModel {");
                    // 向内容构造流中追加入参数据
                    AppendExcelData(stringBuilder, requestModels, 1);
                    // 向内容构造流中追加闭环符号
                    stringBuilder.AppendLine(@"    }");
                    // 向内容构造流中追加空白行
                    stringBuilder.AppendLine();
                    // 向内容构造流中追加响应Model
                    stringBuilder.AppendLine(@"    // 出参类型");
                    stringBuilder.AppendLine(@"    export interface ResponseModel {");
                    // 向内容构造流中追加出参数据
                    AppendExcelData(stringBuilder, responseModels, 1);
                    // 向内容构造流中追加闭环符号
                    stringBuilder.AppendLine(@"    }");
                    // 向内容构造流中追加闭环符号
                    stringBuilder.AppendLine(@"}");
                    /* 到此处为止，所有需要写入到文件的文本内容已全部构造完毕 */
                    // 定义文件名(文件名是命名空间的小写化，每一个单词之间使用 "-" 隔开，然后每一个单词都要小写)
                    string fileName = string.Empty;
                    // 遍历命名空间创建文件名
                    foreach (var str in nameSpaces)
                    {
                        // 判断字母大小写
                        if (!char.IsUpper(str))
                        {
                            // 小写字母
                            fileName += str;
                        }
                        else
                        {
                            // 大写字母
                            fileName = fileName + "-" + str.ToString().ToLower();
                        }
                    }
                    // 判断文件名首字母是否是 "-"
                    if (fileName.Substring(0, 1) == "-")
                    {
                        // 移除掉文件名首字母的 "-"
                        fileName = fileName.Remove(0, 1);
                    }
                    else
                    {
                        // 文件名首字母不是 "-"，无需处理
                    }
                    // 文件名最佳后缀
                    fileName += ".model.ts";
                    // 将内容流写入到文件中(如果文件已存在，则会覆盖原始文件)
                    File.WriteAllText(OutPutDirectory + fileName, stringBuilder.ToString(), Encoding.UTF8);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Excel的“" + dataTable.TableName + "”Sheet页数据处理完毕，已成功导出为ts文件");
                }
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Excel的中的数据已全部处理完毕，按任意键退出");
            }
        }

        /// <summary>
        /// 递归追加参数数据
        /// </summary>
        /// <param name="stringBuilder"></param>
        /// <param name="models"></param>
        /// <param name="level"></param>
        private static void AppendExcelData(StringBuilder stringBuilder, List<ExcelModel> models, int level = 1)
        {
            // 定义补充占位空格数(4个空格)
            string space = "    ";
            // 遍历补充占位符(占位符是4的倍数)
            for (int i = 0; i < level * 4; i++)
            {
                space += " ";
            }
            // 遍历Excel对象模型集合
            foreach (var model in models)
            {
                // 向内容构造流中追加响应Model
                stringBuilder.AppendLine(string.Format(@"{0}// {1}", space, model.D1));
                // 判断此Excel对象模型是否存在子项集合
                if (model.Children.Count < 1)
                {
                    // Excel对象模型没有子集
                    stringBuilder.AppendLine(string.Format(@"{0}{1}: {2}", space, model.C1, model.E1));
                }
                else
                {
                    // Excel对象模型存在子集
                    stringBuilder.AppendLine(string.Format(@"{0}{1}: {2}", space, model.C1, "{"));
                    // 继续向内容构造流中追加参数数据
                    AppendExcelData(stringBuilder, model.Children, level + 1);
                    // 向内容构造流中追加闭环符号
                    stringBuilder.AppendLine(string.Format(@"{0}{1}", space, "}"));
                }
            }
        }

        /// <summary>
        /// 从Html中加载数据
        /// </summary>
        /// /// <param name="filePath"></param>
        private static void HandleHtmlData(string filePath)
        {
            // 判断文件物理路径是否存在
            if (!File.Exists(filePath))
            {
                Console.WriteLine("此Html路径“" + filePath + "”无效，将终止读取操作");
                return;
            }
            Console.WriteLine("暂未实现Html方法，请尝试使用Excel");
        }
    }
}