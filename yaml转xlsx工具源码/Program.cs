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

        // 如果没有找到 .xlsx 文件，输出提示并结束程序
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

        // 处理每一个 .xlsx 文件
        foreach (var excelFilePath in xlsxFiles)
        {
            // 获取文件名（不带扩展名），用于生成对应的 YAML 文件名
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelFilePath);
            string yamlFilePath = Path.Combine(outputFolder, $"{fileNameWithoutExtension}.yaml");

            Console.WriteLine($"正在处理文件: {excelFilePath}");

            // 创建字典用于存储Excel中的数据
            var excelData = new Dictionary<string, object>();

            // 读取Excel文件
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                // 获取第一个工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 获取行列数
                int rows = worksheet.Dimension.Rows;
                int cols = worksheet.Dimension.Columns;

                // 假设第一行为表头，将数据存储到字典中
                for (int row = 2; row <= rows; row++)
                {
                    // 将第一列的值作为YAML的键
                    string rowKey = worksheet.Cells[row, 1].Text;

                    // 创建一个嵌套的字典来存储 Name, Description, FlavorText
                    var rowData = new Dictionary<string, string>
                    {
                        { "Name", worksheet.Cells[row, 2].Text },
                        { "Description", worksheet.Cells[row, 3].Text },
                        { "FlavorText", worksheet.Cells[row, 4].Text }
                    };

                    // 将该行数据添加到excelData字典中，键为第一列的值
                    excelData[rowKey] = rowData;
                }
            }

            // 序列化数据为YAML格式
            var serializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance) // YAML使用CamelCase命名
                .Build();

            var yaml = serializer.Serialize(excelData);

            // 将YAML数据写入文件
            File.WriteAllText(yamlFilePath, yaml);

            Console.WriteLine($"Excel 文件 {fileNameWithoutExtension}.xlsx 成功映射到 YAML 文件 {yamlFilePath}");
        }

        Console.WriteLine("所有 .xlsx 文件处理完毕！");
    }
}