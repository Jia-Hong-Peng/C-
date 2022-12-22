
using System;
using System.IO;
using OfficeOpenXml;

namespace FileContentSearch
{
    class Program
    {
        static void Main(string[] args)
        {
            string folderPath = "C:\\MyFolder\\"; // 資料夾路徑
            string keyword = "hello"; // 關鍵字

            // 檢查資料夾是否存在
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"資料夾 '{folderPath}' 不存在");
                return;
            }

            bool found = false; // 是否找到關鍵字

            // 取得資料夾內的所有文件
            string[] files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);

            foreach (string file in files)
            {
                // 判斷文件副檔名
                string extension = Path.GetExtension(file);

                if (extension == ".txt") // 若是 .txt 檔
                {
                    // 使用 StreamReader 讀取文件內容
                    using (StreamReader reader = new StreamReader(file))
                    {
                        string content = reader.ReadToEnd(); // 讀取整個文件

                        // 判斷是否包含關鍵字
                        if (content.Contains(keyword))
                        {
                            Console.WriteLine($"找到關鍵字 '{keyword}' 於 {file}");
                            found = true;
                        }
                    }
                }
                else if (extension == ".docx" || extension == ".xlsx" || extension == ".pptx") // 若是 Office 檔
                {
                    // 使用 EPPlus 套件讀取文件內容
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
                    {
                        // 取得第一個工作表
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                        // 取得工作表中的所有內容 (使用 All 表示所有列所有欄)
                        object[,] values = worksheet.Cells.Value as object[,];

                        // 迭代每個儲存格
                        for (int row = 1; row <= values.GetLength(0); row++)
                        {
                            for (int col = 1; col <= values.GetLength(1); col++)
                            {
                                // 判斷是否包含關鍵字
                                if (values[row, col]?.ToString().Contains(keyword) == true)
                                {
                                    Console.WriteLine($"找到關鍵字 '{keyword}' 於 {file} (第 {row} 行第 {col} 列)");
                                    found = true;
                                }
                            }
                        }

                        // 若找不到關鍵字，則輸出訊息
                        if (!found)
                        {
                            Console.WriteLine($"在資料夾 '{folderPath}' 內的文件中，找不到關鍵字 '{keyword}'");
                        }

                        // 等待使用者按下 Enter 鍵結束程式
                        Console.ReadLine();
                    }
                }
            }
        }
    }
}
    