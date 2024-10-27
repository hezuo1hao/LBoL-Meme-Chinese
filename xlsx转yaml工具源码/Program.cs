using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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

        // 搜索当前目录下的所有 .xlsx 文件
        string[] xlsxFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

        Console.WriteLine($"当前程序目录: {currentDirectory}");

        if (xlsxFiles.Length == 0)
        {
            Console.WriteLine("没有找到 .xlsx 文件。");
            return;
        }

        // 定义输出路径的“语言包”文件夹
        string outputFolder = Path.Combine(currentDirectory, "语言包");

        // 如果“语言包”文件夹不存在，则创建它
        if (!Directory.Exists(outputFolder))
        {
            Directory.CreateDirectory(outputFolder);
            Console.WriteLine($"创建语言包文件夹: {outputFolder}");
        }

        foreach (var excelFilePath in xlsxFiles)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelFilePath);

            Console.WriteLine($"正在处理文件: {excelFilePath}");

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    string sheetName = worksheet.Name;
                    string yamlFilePath = Path.Combine(outputFolder, $"{sheetName}.yaml");

                    // 存储最终生成的嵌套数据结构
                    var excelData = new Dictionary<string, object>();

                    int rows = worksheet.Dimension.Rows;
                    int cols = worksheet.Dimension.Columns;

                    // 检查是否至少有两列 (Key, 文本)
                    if (cols < 2)
                    {
                        Console.WriteLine($"工作表 {sheetName} 列数不足，跳过。");
                        continue;
                    }

                    // 遍历每一行，从第二行开始，第一行是表头
                    for (int row = 2; row <= rows; row++)
                    {
                        string key = worksheet.Cells[row, 1].Text;  // Key 列
                        string value = worksheet.Cells[row, 2].Text;  // 文本列

                        // 跳过空行或没有键值的行
                        if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }

                        // 将键 (Key) 按 "." 分隔开，生成嵌套的字典结构
                        AddNestedData(excelData, key.Split('.'), value);
                    }

                    // 使用默认的 YAML 序列化设置
                    var serializer = new SerializerBuilder()
                        .WithNamingConvention(CamelCaseNamingConvention.Instance)
                        .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull)
                        .Build();

                    var yaml = serializer.Serialize(excelData).Trim();

                    yaml = AddBlankLineBetweenTopLevelEntries(yaml);

                    File.WriteAllText(yamlFilePath, yaml);

                    Console.WriteLine($"工作表 {sheetName} 已成功输出到 YAML 文件 {yamlFilePath}");
                }
            }
        }

        Console.WriteLine("所有 .xlsx 文件处理完毕！");
    }

    // 在顶级条目之间添加空行
    static string AddBlankLineBetweenTopLevelEntries(string yamlContent)
    {
        yamlContent = yamlContent.Replace(">-", "|-");

        var lines = yamlContent.Split('\n');
        var result = new List<string>();

        for (int i = 0; i < lines.Length; i++)
        {
            // 跳过空行
            if (string.IsNullOrWhiteSpace(lines[i]) && i + 1 < lines.Length && !string.IsNullOrWhiteSpace(lines[i + 1]))
                continue;

            result.Add(lines[i]);
            // 下一个条目也是顶级条目，且非空白时，插入空行
            if (i < lines.Length - 1 && !string.IsNullOrWhiteSpace(lines[i + 1]) && !lines[i + 1].StartsWith("  "))
            {
                result.Add("");  // 插入空行
            }
        }

        return string.Join("\n", result);
    }

    // 递归地将键值对插入嵌套字典结构
    static void AddNestedData(Dictionary<string, object> parent, string[] keys, string value)
    {
        for (int i = 0; i < keys.Length; i++)
        {
            string currentKey = keys[i];

            if (i == keys.Length - 1)
            {
                // 检查值是否为列表格式
                if (value.StartsWith("- "))
                {
                    // 将值拆分成列表
                    var list = new List<string>();
                    foreach (var item in value.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (item.StartsWith("- "))
                        {
                            list.Add(item.Substring(2).Trim());
                        }
                    }
                    parent[currentKey] = list;
                }
                else
                {
                    parent[currentKey] = value;
                }
            }
            else
            {
                // 如果当前级别的键不存在，则创建一个新字典
                if (!parent.ContainsKey(currentKey))
                {
                    parent[currentKey] = new Dictionary<string, object>();
                }

                // 继续递归处理下一级
                parent = (Dictionary<string, object>)parent[currentKey];
            }
        }
    }
}
