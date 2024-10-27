using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

class Program
{
    static void Main(string[] args)
    {
        // 设置ExcelPackage.LicenseContext为非商业使用
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // 获取当前程序的运行目录
        string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

        // 搜索当前目录下的所有 .yaml 文件
        string[] yamlFiles = Directory.GetFiles(currentDirectory, "*.yaml");

        Console.WriteLine($"当前程序目录: {currentDirectory}");

        // 如果没有找到 .yaml 文件，输出提示并结束程序
        if (yamlFiles.Length == 0)
        {
            Console.WriteLine("没有找到 .yaml 文件。");
            return;
        }

        // 定义输出路径的Excel文件夹
        string outputFolder = Path.Combine(currentDirectory, "生成的Excel文件");

        // 如果Excel文件夹不存在，则创建它
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
            Console.WriteLine($"创建生成的Excel文件夹: {outputFolder}");
        }

        // 定义Excel文件的完整路径（将所有工作表写入到同一个Excel文件中）
        string excelFilePath = Path.Combine(outputFolder, "合并的Yaml数据.xlsx");

        // 创建一个ExcelPackage对象
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            // 处理每一个 .yaml 文件，创建对应的工作表
            foreach (var yamlFilePath in yamlFiles)
            {
                // 获取文件名（不带扩展名），用于生成对应的工作表名
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(yamlFilePath);

                Console.WriteLine($"正在处理文件: {yamlFilePath}");

                // 读取YAML文件
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance) // 使用CamelCase命名
                    .Build();

                // 这里将类型修改为 Dictionary<string, object> 以处理多种情况
                Dictionary<string, object> yamlData;

                using (var reader = new StreamReader(yamlFilePath))
                {
                    yamlData = deserializer.Deserialize<Dictionary<string, object>>(reader);
                }

                // 创建新的工作表，命名为yaml文件的名称
                var worksheet = package.Workbook.Worksheets.Add(fileNameWithoutExtension);

                // 写入表头 ('Key', '文本')
                worksheet.Cells[1, 1].Value = "Key";
                worksheet.Cells[1, 2].Value = "文本";

                // 初始化行索引
                int rowIndex = 2; // 从第2行开始，第一行为表头

                // 递归处理每一个键值对
                void WriteYamlData(Dictionary<string, object> data, string parentKey)
                {
                    foreach (var entry in data)
                    {
                        string key = string.IsNullOrEmpty(parentKey) ? entry.Key : $"{parentKey}.{entry.Key}";
                        var value = entry.Value;

                        if (value is string stringValue)
                        {
                            // 如果值是字符串，直接写入
                            worksheet.Cells[rowIndex, 1].Value = key;
                            worksheet.Cells[rowIndex, 2].Value = stringValue;
                            rowIndex++;
                        }
                        else if (value is IList<object> listValue)
                        {
                            // 如果值是列表，将其格式化为多行字符串
                            string combinedValue = string.Join("\n- ", listValue.Select(v => v.ToString()));
                            combinedValue = $"- {combinedValue}";
                            worksheet.Cells[rowIndex, 1].Value = key;
                            worksheet.Cells[rowIndex, 2].Value = combinedValue;
                            rowIndex++;
                        }
                        else if (value is Dictionary<object, object> dictValue)
                        {
                            // 将嵌套字典转换为Dictionary<string, object>再递归处理
                            var stringKeyDict = dictValue.ToDictionary(k => k.Key.ToString(), k => k.Value);
                            WriteYamlData(stringKeyDict, key);
                        }
                    }
                    // 在每个顶级条目之间插入一行空行
                    rowIndex++;
                }

                // 写入每一个键值对，初始调用时没有父键
                WriteYamlData(yamlData, string.Empty);

                Console.WriteLine($"YAML 文件 {fileNameWithoutExtension}.yaml 已写入工作表");
            }

            // 保存整个Excel文件
            package.Save();
        }

        Console.WriteLine($"所有 .yaml 文件处理完毕，已生成 Excel 文件: {excelFilePath}");
    }
}
