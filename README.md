# UnitywithExcel

## Unity中对Excel表格进行读写
### 准备工作：
#### 1.从Nuget包中下载EPPLus包
#### 2.将EPPlus.dll，EPPlus.Interfaces.dll，Microsoft.IO.RecyclableMemoryStream.dll等依赖的dll放置Plugins文件夹下
#### 3.注意 EPPLus在5.0版本以后需要用许可证，若无则可以添加一条声明做为个人非商用许可说明
#### 4.补充一下很坑的地方 Unity打包出来后的应用 运行会遇到一个编码437的问题 ，需要到UnitrEditor\Data\MonoBleedingEdge\lib\mono\unityjit下找到I18N.dll和I18N.West.dll文件；将其放到打包出来后的执行文件夹中 
        public void ExportToExcel(List<DuplicateLine> duplicates, string filePath)
        {
            ExcelPackage.License.SetNonCommercialPersonal("唐力der");
            // 创建 Excel 文件
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Duplicates");
                
                // Debug.Log(1111);
                // 添加表头
                worksheet.Cells[1, 1].Value = "Compare Text";
                worksheet.Cells[1, 2].Value = "Similar Text";
                worksheet.Cells[1, 3].Value = "Page Numbers 1";
                worksheet.Cells[1, 4].Value = "Page Numbers 2";

                // 填充数据
                for (int i = 0; i < duplicates.Count; i++)
                {
                    var duplicate = duplicates[i];
                    // 设置 Compare 和 Similar Text 的颜色
                    if (duplicate.TypeNum == 1)
                    {
                        worksheet.Cells[i + 2, 1].Value = duplicate.CompareText;
                        worksheet.Cells[i + 2, 1].Style.Font.Color.SetColor(Color.Red);
                
                        worksheet.Cells[i + 2, 2].Value = duplicate.SimliarText;
                        worksheet.Cells[i + 2, 2].Style.Font.Color.SetColor(Color.Red);
                    }
                    else
                    {
                        worksheet.Cells[i + 2, 1].Value = duplicate.CompareText;
                        worksheet.Cells[i + 2, 1].Style.Font.Color.SetColor(Color.Green);
                
                        worksheet.Cells[i + 2, 2].Value = duplicate.SimliarText;
                        worksheet.Cells[i + 2, 2].Style.Font.Color.SetColor(Color.Green);
                    }

                    // 设置页码并设置黑色背景
                    worksheet.Cells[i + 2, 3].Value = string.Join(", ", duplicate.PageNumbers1);
                    worksheet.Cells[i + 2, 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    // worksheet.Cells[i + 2, 3].Style.Fill.BackgroundColor.SetColor(Color.Black);
                    worksheet.Cells[i + 2, 3].Style.Font.Color.SetColor(Color.Black); // 设置字体颜色为白色

                    worksheet.Cells[i + 2, 4].Value = string.Join(", ", duplicate.PageNumbers2);
                    worksheet.Cells[i + 2, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    // worksheet.Cells[i + 2, 4].Style.Fill.BackgroundColor.SetColor(Color.Black);
                    worksheet.Cells[i + 2, 4].Style.Font.Color.SetColor(Color.Black); // 设
                    // 启用换行
                    
                
                    worksheet.Cells[i + 2, 1, i + 2, 4].Style.WrapText = true;
                }
                
                // 设置列宽
                worksheet.Column(1).Width = 80; // 设置 Compare Text 列宽
                worksheet.Column(2).Width = 80; // 设置 Similar Text 列宽
                worksheet.Column(3).Width = 15; // 设置 Page Numbers 1 列宽
                worksheet.Column(4).Width = 15; // 设置 Page Numbers 2 列宽

                // 自动调整列宽
                // worksheet.Cells.AutoFitColumns();
                
                // 保存文件
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }


![image](https://github.com/user-attachments/assets/8df50b64-182d-4a4d-acdb-31e37c022a40)

![image](https://github.com/user-attachments/assets/231038ce-953e-4bd5-b721-34173022e779)
