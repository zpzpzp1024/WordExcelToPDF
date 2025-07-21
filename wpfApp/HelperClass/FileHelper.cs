using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace wpfApp.HelperClass
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class FileHelper
    {
        public List<string> wordFiles { get; set; } = new List<string>();
        public List<string> excelFiles { get; set; } = new List<string>();

        public void OpenFile()
        {
            Process.Start("explorer.exe", "C:\\");
        }

        // 接收文件数据的方法


        public void ReceiveFileAsBase64(string base64Data, string fileName)
        {
            byte[] fileData = Convert.FromBase64String(base64Data);
            string savePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);
            File.WriteAllBytes(savePath, fileData);
            string extension = System.IO.Path.GetExtension(fileName.ToLower());

            if (extension == ".xls" || extension == ".xlsx" || extension == ".xlsm")
            {
                excelFiles.Add(savePath);
            }
            else if (extension == ".doc" || extension == ".docx")
            {
                wordFiles.Add(savePath);
            }
        }

        public bool ConvertAndMergeToPDF()
        {
            var newFilePath = new List<string>();

            string tempPath = Path.GetTempPath();

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var finalName = "";
            try
            {


                if (excelFiles.Count > 0)
                {
                    finalName = Path.Combine(desktopPath,
                        Path.GetFileNameWithoutExtension(excelFiles[0]) + ".pdf");
                    foreach (var VARIABLE in excelFiles)
                    {
                        var sd = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(VARIABLE) + ".pdf");
                        WordExcel2PDF.ExcelToPdf(VARIABLE, sd);
                        newFilePath.Add(sd);
                    }
                }


                if (wordFiles.Count > 0)
                {
                    finalName = Path.Combine(desktopPath,
                        Path.GetFileNameWithoutExtension(wordFiles[0]) + ".pdf");
                    foreach (var VARIABLE in wordFiles)
                    {
                        var sd = Path.Combine(tempPath, Path.GetFileNameWithoutExtension(VARIABLE) + ".pdf");
                        WordExcel2PDF.WordToPdf(VARIABLE, sd);
                        newFilePath.Add(sd);
                    }
                }

                using (PdfDocument outputDoc = new PdfDocument())
                {
                    foreach (string file in newFilePath)
                    {
                        using (PdfDocument inputDoc = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < inputDoc.PageCount; i++)
                            {
                                outputDoc.AddPage(inputDoc.Pages[i]);
                            }
                        }
                    }

                    outputDoc.Save(finalName);
                    wordFiles = new List<string>();
                    excelFiles = new List<string>();
                    return true;
                    // 停止旋转动画
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("文件合并失败：" + e.Message);
                return false;
            }
            return true;
        }
    }

    public static class WordExcel2PDF
    {
        public static void WordToPdf(string sourcepath, string outpath)
        {
            // 创建Word应用程序对象
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            try
            {
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone; // 关闭所有警告提示框
                // 打开Word文档
                var doc = wordApp.Documents.Open(sourcepath, ReadOnly: true);

                // 将Word文档保存为PDF格式
                doc.ExportAsFixedFormat(outpath, WdExportFormat.wdExportFormatPDF); // 禁止覆盖提示

                // 关闭Word文档
                doc.Close(false);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                // 退出Word应用程序
                wordApp.Quit();

                // 获取所有Excel进程
                Process[] processes = Process.GetProcessesByName(Path.GetFileName(sourcepath) + " - Word");

                // 遍历所有Excel进程并杀死它们
                foreach (Process process in processes)
                {
                    process.Kill();
                    process.WaitForExit();
                    process.Dispose();
                }
            }
        }

        public static void ExcelToPdf(string sourcepath, string outpath)
        {
            // 创建Excel应用程序对象
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                excelApp.DisplayAlerts = false; // 关闭所有警告提示框
                // 打开Excel文档
                Workbook workbook = excelApp.Workbooks.Open(sourcepath, ReadOnly: true);
                Worksheet worksheet = workbook.ActiveSheet;
                // 将Excel活动工作表保存为PDF格式
                worksheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outpath);
                //workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outpath);
                // 关闭Excel文档
                workbook.Close(false);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                // 退出Excel应用程序
                excelApp.Quit();
                // 获取所有Excel进程
                Process[] processes = Process.GetProcessesByName(Path.GetFileName(sourcepath) + " - Excel");

                // 遍历所有Excel进程并杀死它们
                foreach (Process process in processes)
                {
                    process.Kill();
                    process.WaitForExit();
                    process.Dispose();
                }
            }
        }
    }
}