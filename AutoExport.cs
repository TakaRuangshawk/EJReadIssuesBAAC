using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using Renci.SshNet;

namespace EJReadIssuesBAAC
{
    internal static class AutoExport
    {
        /// <summary>
        /// ดึง EJ ทั้งหมดของ targetDate จากทุกตู้ -> จับ prob จาก mastercode_baac.csv -> สร้าง CSV/XLSX ไปยัง saveBaseDir
        /// </summary>
        public static void RunExport(string saveBaseDir, DateTime targetDate)
        {
            string dateStr = targetDate.ToString("yyyyMMdd");
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string projectEJPath = Path.Combine(baseDir, "EJ", dateStr);
            string masterCodePath = Path.Combine(baseDir, "config", "mastercode_baac.csv");

            if (Directory.Exists(projectEJPath))
                Directory.Delete(projectEJPath, true);
            Directory.CreateDirectory(projectEJPath);

            if (!File.Exists(masterCodePath))
                throw new FileNotFoundException("ไม่พบไฟล์ mastercode_baac.csv", masterCodePath);

            string serversPath = Path.Combine(baseDir, "config", "servers.txt");
            if (!File.Exists(serversPath))
                throw new FileNotFoundException("ไม่พบไฟล์ servers.txt", serversPath);

            var serverIps = File.ReadAllLines(serversPath)
                                .Select(ip => ip.Trim())
                                .Where(ip => !string.IsNullOrWhiteSpace(ip))
                                .ToList();

            bool isTodayMode = targetDate.Date == DateTime.Today;

            // 1) ดึงไฟล์ต้นทาง
            foreach (var serverIp in serverIps)
            {
                using var sftp = new SftpClient(serverIp, "root", "12qwaszx!@QWASZX");
                sftp.Connect();

                if (isTodayMode)
                {
                    DownloadCurrentTxtFiles(sftp, projectEJPath, targetDate);
                }
                else
                {
                    DownloadOperationZipFiles(sftp, projectEJPath, targetDate);
                }

                sftp.Disconnect();
            }

            // 2) โหลด master code
            var probMap = File.ReadAllLines(masterCodePath)
               .Skip(1)
               .Select(line =>
               {
                   var parts = line.Split(',');
                   string rawName = (parts.Length > 1 ? parts[1] : "").Trim();
                   string cleanName = rawName.Replace("\\\"", "\"").Trim('"');

                   return new
                   {
                       Code = (parts.Length > 0 ? parts[0] : "").Trim(),
                       Name = cleanName,
                       Remark = (parts.Length > 2 ? parts[2] : "").Trim('"')
                   };
               })
               .Where(x => !string.IsNullOrEmpty(x.Code) && !string.IsNullOrEmpty(x.Name))
               .ToList();

            // 3) สแกน txt -> สร้าง CSV rows
            var output = new List<string>
            {
                "terminalid,probcode,remark,dtenderrcode13,dterrcode13,trxdatetime,status,createdate,updatedate,resolveprob"
            };

            DateTime now = DateTime.Now;

            foreach (var txtFile in Directory.GetFiles(projectEJPath, "*.txt", SearchOption.AllDirectories))
            {
                string terminalId = GetTerminalIdFromPath(txtFile);
                if (string.IsNullOrWhiteSpace(terminalId))
                    continue;

                string fileNameOnly = Path.GetFileNameWithoutExtension(txtFile); // EJ20260402
                string datePartStr = fileNameOnly.Replace("EJ", "");

                DateTime datePart;
                if (!DateTime.TryParseExact(
                        datePartStr,
                        "yyyyMMdd",
                        null,
                        System.Globalization.DateTimeStyles.None,
                        out datePart))
                {
                    datePart = targetDate.Date;
                }

                var lines = File.ReadAllLines(txtFile);
                DateTime? lastKnownTime = null;

                foreach (var line in lines)
                {
                    var match = System.Text.RegularExpressions.Regex.Match(
                        line,
                        @"\d{2}:\d{2}:\d{2}(\.\d{1,3})?");

                    int contentStartIndex = 0;

                    if (match.Success)
                    {
                        string timeOnly = match.Value;

                        if (!DateTime.TryParseExact(
                                timeOnly,
                                "HH:mm:ss.fff",
                                null,
                                System.Globalization.DateTimeStyles.None,
                                out var timeParsed))
                        {
                            DateTime.TryParseExact(
                                timeOnly,
                                "HH:mm:ss",
                                null,
                                System.Globalization.DateTimeStyles.None,
                                out timeParsed);
                        }

                        lastKnownTime = datePart.Date + timeParsed.TimeOfDay;
                        contentStartIndex = match.Index + match.Length;
                    }

                    string lineContent = line.Substring(contentStartIndex).Trim();

                    foreach (var prob in probMap)
                    {
                        if (lineContent.Contains(prob.Name))
                        {
                            DateTime trxTime = lastKnownTime ?? datePart.Date;
                            string remark = lineContent.Replace("\"", "\"\"");

                            string newLine =
                                $"{terminalId},\"{prob.Code}\",\"{remark}\",,,{trxTime:yyyy-MM-dd HH:mm:ss},1,{now:yyyy-MM-dd HH:mm:ss},{now:yyyy-MM-dd HH:mm:ss},operation";

                            output.Add(newLine);
                            break;
                        }
                    }
                }
            }

            // 4) เขียน CSV + XLSX
            Directory.CreateDirectory(saveBaseDir);

            string csvPath = Path.Combine(saveBaseDir, $"reportcase_{dateStr}.csv");
            string xlsxPath = Path.Combine(saveBaseDir, $"reportcase_{dateStr}.xlsx");

            File.WriteAllLines(csvPath, output, Encoding.UTF8);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Report Case");
                var headers = output[0].Split(',');

                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = headers[i];
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }

                for (int row = 1; row < output.Count; row++)
                {
                    var fields = output[row].Split(',');
                    for (int col = 0; col < fields.Length; col++)
                    {
                        worksheet.Cell(row + 1, col + 1).Value = fields[col].Trim('"');
                    }
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(xlsxPath);
            }

            // 5) เก็บกวาด
            Directory.Delete(projectEJPath, true);
        }

        private static void DownloadOperationZipFiles(SftpClient sftp, string projectEJPath, DateTime targetDate)
        {
            string dateStr = targetDate.ToString("yyyyMMdd");
            string remoteDir = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/";

            if (!sftp.Exists(remoteDir))
                return;

            var zips = sftp.ListDirectory(remoteDir)
                           .Where(f => !f.IsDirectory && f.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                           .ToList();

            foreach (var zip in zips)
            {
                string localZip = Path.Combine(projectEJPath, zip.Name);

                using (var fileStream = File.Create(localZip))
                    sftp.DownloadFile(zip.FullName, fileStream);

                string tempExtractPath = Path.Combine(projectEJPath, Path.GetFileNameWithoutExtension(zip.Name));
                Directory.CreateDirectory(tempExtractPath);

                ZipFile.ExtractToDirectory(localZip, tempExtractPath, true);
                File.Delete(localZip);
            }
        }

        private static void DownloadCurrentTxtFiles(SftpClient sftp, string projectEJPath, DateTime targetDate)
        {
            string currentRoot = "/data1/fileserverBAAC/EJ/Current/";

            if (!sftp.Exists(currentRoot))
                return;

            var terminalDirs = sftp.ListDirectory(currentRoot)
                                   .Where(f => f.IsDirectory && f.Name != "." && f.Name != "..")
                                   .ToList();

            foreach (var dir in terminalDirs)
            {
                string terminalId = dir.Name;
                string remoteTxt =
                    $"/data1/fileserverBAAC/EJ/Current/{terminalId}/{targetDate:yyyy}/{targetDate:MM}/EJ{targetDate:yyyyMMdd}.txt";

                if (!sftp.Exists(remoteTxt))
                    continue;

                string localDir = Path.Combine(projectEJPath, terminalId);
                Directory.CreateDirectory(localDir);

                string localTxt = Path.Combine(localDir, $"EJ{targetDate:yyyyMMdd}.txt");

                using (var fileStream = File.Create(localTxt))
                    sftp.DownloadFile(remoteTxt, fileStream);
            }
        }

        private static string GetTerminalIdFromPath(string txtFile)
        {
            string parentDir = new DirectoryInfo(Path.GetDirectoryName(txtFile)!).Name;

            // today mode: EJ\yyyyMMdd\T023...\EJyyyyMMdd.txt
            if (parentDir.StartsWith("T", StringComparison.OrdinalIgnoreCase)
                && !parentDir.Contains("_EJ", StringComparison.OrdinalIgnoreCase))
            {
                return parentDir;
            }

            // operation mode: EJ\yyyyMMdd\T023..._EJyyyyMMdd\xxx.txt
            if (parentDir.Contains("_EJ", StringComparison.OrdinalIgnoreCase))
            {
                return parentDir.Split(new[] { "_EJ" }, StringSplitOptions.None)[0];
            }

            return string.Empty;
        }
    }
}