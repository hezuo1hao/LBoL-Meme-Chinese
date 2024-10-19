using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

// 狗屁通是人类的好帮手
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
                    string yamlFilePath = Path.Combine(outputFolder, $"{fileNameWithoutExtension}_{sheetName}.yaml");

                    var excelData = new Dictionary<string, Dictionary<string, object>>();

                    int rows = worksheet.Dimension.Rows;
                    int cols = worksheet.Dimension.Columns;

                    var headers = new List<string>();
                    for (int col = 1; col <= cols; col++)
                    {
                        string header = worksheet.Cells[1, col].Text;
                        if (!string.IsNullOrWhiteSpace(header))
                        {
                            headers.Add(header);
                        }
                    }

                    for (int row = 2; row <= rows; row++)
                    {
                        var rowData = new Dictionary<string, object>();

                        for (int col = 1; col <= headers.Count; col++)
                        {
                            string header = headers[col - 1];
                            string cellValue = worksheet.Cells[row, col].Text;

                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                rowData[header] = cellValue;
                            }
                        }

                        string rowKey = worksheet.Cells[row, 1].Text;
                        if (!string.IsNullOrWhiteSpace(rowKey))
                        {
                            excelData[rowKey] = rowData;
                        }
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

    // 在顶级条目（如 ReimuAttackR 和 ReimuBlockW）之间添加空行
    static string AddBlankLineBetweenTopLevelEntries(string yamlContent)
    {
        yamlContent = yamlContent.Replace(">-", "|-");

        var lines = yamlContent.Split('\n');
        var result = new List<string>();

        for (int i = 0; i < lines.Length; i++)
        {
            // 跳过所有的空行
            if (string.IsNullOrWhiteSpace(lines[i]))
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
}
