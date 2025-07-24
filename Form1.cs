using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using Renci.SshNet;

namespace EJReadIssuesBAAC
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            // เซ็ตค่า default ลง TextBox
            string lastFolder = Properties.Settings.Default.LastFolderPath;
            if (string.IsNullOrWhiteSpace(lastFolder) || !Directory.Exists(lastFolder))
            {
                txtFolderPath.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
            {
                txtFolderPath.Text = lastFolder;
            }
            datePicker.Value = DateTime.Today;
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            string selected = cmbTerminalId.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(selected) || !selected.Contains("|"))
            {
                lblStatus.Text = "กรุณาเลือกไฟล์ ZIP";
                return;
            }

            var parts = selected.Split('|');
            string serverIp = parts[0].Trim();
            string fileName = parts[1].Trim();

            string dateStr = datePicker.Value.ToString("yyyyMMdd");
            string remoteFile = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/{fileName}";

            string baseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
            ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            : txtFolderPath.Text;
            string localPath = Path.Combine(baseDir, fileName);


            try
            {
                using (var sftp = new SftpClient(serverIp, "root", "12qwaszx!@QWASZX"))
                {
                    sftp.Connect();
                    using (var fileStream = File.Create(localPath))
                    {
                        sftp.DownloadFile(remoteFile, fileStream);
                    }
                    sftp.Disconnect();
                }

                lblStatus.ForeColor = Color.Green;
                lblStatus.Text = $"✅ ดาวน์โหลดสำเร็จ: {fileName} จาก {serverIp}";
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = $"เกิดข้อผิดพลาดจาก {serverIp}:\n{ex.Message}";
            }
        }


        private void LoadTerminalIdsFromSftp()
        {
            string dateStr = datePicker.Value.ToString("yyyyMMdd");
            string remoteDirFormat = "/data1/fileserverBAAC/EJ/Operation/{0}/";
            string serverConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "servers.txt");

            if (!File.Exists(serverConfigPath))
            {
                MessageBox.Show("ไม่พบไฟล์ servers.txt", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var serverIps = File.ReadAllLines(serverConfigPath)
                                .Select(ip => ip.Trim())
                                .Where(ip => !string.IsNullOrWhiteSpace(ip))
                                .ToList();

            cmbTerminalId.Items.Clear();
            cmbTerminalId.Items.Add("");
            foreach (var serverIp in serverIps)
            {
                try
                {
                    using (var sftp = new SftpClient(serverIp, "root", "12qwaszx!@QWASZX"))
                    {
                        sftp.Connect();
                        string remoteDir = string.Format(remoteDirFormat, dateStr);

                        if (!sftp.Exists(remoteDir))
                        {
                            continue;
                        }

                        var zipFiles = sftp.ListDirectory(remoteDir)
                                           .Where(f => !f.IsDirectory && f.Name.EndsWith(".zip"))
                                           .Select(f => $"{serverIp}|{f.Name}");

                        cmbTerminalId.Items.AddRange(zipFiles.ToArray());

                        sftp.Disconnect();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"โหลดจาก {serverIp} ล้มเหลว:\n{ex.Message}", "SFTP ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            if (cmbTerminalId.Items.Count > 0)
                cmbTerminalId.SelectedIndex = 0;
            else
                MessageBox.Show("ไม่พบ ZIP จากทุกเครื่อง", "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }




        private void datePicker_ValueChanged(object sender, EventArgs e)
        {
            if (datePicker.Value.Date < DateTime.Today)
            {
                LoadTerminalIdsFromSftp();
            }
        }


        private void btnDownloadAll_Click(object sender, EventArgs e)
        {
            if (cmbTerminalId.Items.Count == 0)
            {
                MessageBox.Show("ไม่มีไฟล์ให้ดาวน์โหลด", "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string dateStr = datePicker.Value.ToString("yyyyMMdd");
            string remoteDir = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/";

            // ✅ ใช้ path จาก txtFolderPath หรือ fallback เป็น Desktop
            string baseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : txtFolderPath.Text;

            string localDir = Path.Combine(baseDir, $"EJ_{dateStr}");

            string serverConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "servers.txt");
            if (!File.Exists(serverConfigPath))
            {
                MessageBox.Show("ไม่พบไฟล์ config/servers.txt", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var serverIps = File.ReadAllLines(serverConfigPath)
                                .Select(ip => ip.Trim())
                                .Where(ip => !string.IsNullOrWhiteSpace(ip))
                                .ToList();

            if (serverIps.Count == 0)
            {
                MessageBox.Show("ไม่มีรายการ IP ในไฟล์ servers.txt", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Directory.CreateDirectory(localDir);

                foreach (var item in cmbTerminalId.Items)
                {
                    string itemStr = item.ToString();
                    string[] parts = itemStr.Split('|');
                    if (parts.Length != 2) continue;

                    string ip = parts[0];
                    string fileName = parts[1];

                    if (!serverIps.Contains(ip)) continue; // skip IP ที่ไม่อยู่ใน servers.txt

                    string remoteFile = remoteDir + fileName;
                    string safeIp = ip.Replace(".", "_"); // เพื่อใช้ตั้งชื่อไฟล์
                    string localZipPath = Path.Combine(localDir, $"{safeIp}_{fileName}");

                    using (var sftp = new SftpClient(ip, "root", "12qwaszx!@QWASZX"))
                    {
                        sftp.Connect();

                        // ✅ ดาวน์โหลด .zip
                        using (var fileStream = File.Create(localZipPath))
                        {
                            sftp.DownloadFile(remoteFile, fileStream);
                        }

                        try
                        {
                            // ✅ แตกไฟล์ .txt และตั้งชื่อใหม่ตาม .zip
                            using (var archive = System.IO.Compression.ZipFile.OpenRead(localZipPath))
                            {
                                foreach (var entry in archive.Entries)
                                {
                                    if (!entry.FullName.EndsWith(".txt")) continue;

                                    string newTxtName = Path.GetFileNameWithoutExtension(fileName) + ".txt";
                                    string targetPath = Path.Combine(localDir, newTxtName);

                                    entry.ExtractToFile(targetPath, overwrite: true);
                                }
                            }

                            // 🧹 ลบ zip หลังแตกเสร็จ
                            File.Delete(localZipPath);
                        }
                        catch (Exception innerEx)
                        {
                            MessageBox.Show($"❌ แตก zip ผิดพลาด: {fileName}\n{innerEx.Message}", "Extract Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                        sftp.Disconnect();
                    }
                }

                MessageBox.Show($"✅ ดาวน์โหลดและแตกไฟล์ทั้งหมดสำเร็จ\n\nที่เก็บ: {localDir}", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"เกิดข้อผิดพลาดระหว่างดาวน์โหลดทั้งหมด:\n\n{ex.Message}", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportEJAndCsv_Click(object sender, EventArgs e)
        {
            string dateStr = datePicker.Value.ToString("yyyyMMdd");
            string projectEJPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EJ", dateStr);
            string masterCodePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "mastercode_baac.csv");
            if (Directory.Exists(projectEJPath))
            {
                Directory.Delete(projectEJPath, true); // ลบทั้งหมดทั้งไฟล์และ subfolder
            }
            Directory.CreateDirectory(projectEJPath); // สร้างใหม่

            if (!File.Exists(masterCodePath))
            {
                MessageBox.Show("ไม่พบไฟล์ mastercode_baac.csv ในโฟลเดอร์ config", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Directory.CreateDirectory(projectEJPath); // สร้างโฟลเดอร์ EJ/yyyyMMdd

            try
            {
                string serversPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config", "servers.txt");
                if (!File.Exists(serversPath))
                {
                    MessageBox.Show("ไม่พบไฟล์ servers.txt ในโฟลเดอร์ config", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var serverIps = File.ReadAllLines(serversPath)
                                    .Select(ip => ip.Trim())
                                    .Where(ip => !string.IsNullOrWhiteSpace(ip))
                                    .ToList();
                // โหลด ZIP จาก SFTP
                foreach (var serverIp in serverIps)
                {
                    using (var sftp = new SftpClient(serverIp, "root", "12qwaszx!@QWASZX"))
                    {
                        sftp.Connect();

                        string remoteDir = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/";
                        var zips = sftp.ListDirectory(remoteDir)
                                       .Where(f => !f.IsDirectory && f.Name.EndsWith(".zip"))
                                       .ToList();

                        foreach (var zip in zips)
                        {
                            string localZip = Path.Combine(projectEJPath, zip.Name);
                            using (var fileStream = File.Create(localZip))
                            {
                                sftp.DownloadFile(zip.FullName, fileStream);
                            }

                            // แตก ZIP
                            string tempExtractPath = Path.Combine(projectEJPath, Path.GetFileNameWithoutExtension(zip.Name));
                            Directory.CreateDirectory(tempExtractPath);
                            System.IO.Compression.ZipFile.ExtractToDirectory(localZip, tempExtractPath);

                            // ลบ ZIP หลังแตก
                            File.Delete(localZip);
                        }

                        sftp.Disconnect();
                    }
                }
                // อ่าน probname จาก csv
                var probMap = File.ReadAllLines(masterCodePath)
                   .Skip(1)
                   .Select(line =>
                   {
                       var parts = line.Split(',');
                
                       string rawName = parts[1].Trim();
                       string cleanName = rawName.Replace("\\\"", "\"").Trim('"');  // แก้กรณี \"
                
                       return new
                       {
                           Code = parts[0].Trim(),       // DEVICE05
                           Name = cleanName,             // TIME OUT, RETAIN CARD.
                           Remark = parts[2].Trim('"')   // DEVICE (ใช้ probtype)
                       };
                   })
                   .Where(x => !string.IsNullOrEmpty(x.Code) && !string.IsNullOrEmpty(x.Name))
                   .ToList();


                // สร้าง CSV output
                var output = new List<string>();
                output.Add("terminalid,probcode,remark,dtenderrcode13,dterrcode13,trxdatetime,status,createdate,updatedate,resolveprob");
                DateTime now = DateTime.Now;

                foreach (var dir in Directory.GetDirectories(projectEJPath))
                {
                    //string terminalId = Path.GetFileName(dir);
                    foreach (var txtFile in Directory.GetFiles(dir, "*.txt"))
                    {
                        string folderName = new DirectoryInfo(Path.GetDirectoryName(txtFile)).Name;
                        string terminalId = folderName.Split("_EJ")[0];
                        string fileNameOnly = Path.GetFileNameWithoutExtension(txtFile); // EJ20250709
                        string datePartStr = fileNameOnly.Replace("EJ", "");

                        DateTime datePart;
                        if (!DateTime.TryParseExact(datePartStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out datePart))
                        {
                            datePart = now.Date; // fallback ถ้า parse ไม่ได้
                        }

                        var lines = File.ReadAllLines(txtFile);
                        DateTime? lastKnownTime = null;

                        foreach (var line in lines)
                        {
                            // ✅ พยายามหาเวลาในทุกบรรทัด แล้วเก็บไว้
                            var match = System.Text.RegularExpressions.Regex.Match(line, @"\d{2}:\d{2}:\d{2}(\.\d{1,3})?");
                            if (match.Success)
                            {
                                string timeOnly = match.Value;
                                if (!DateTime.TryParseExact(timeOnly, "HH:mm:ss.fff", null, System.Globalization.DateTimeStyles.None, out var timeParsed))
                                {
                                    DateTime.TryParseExact(timeOnly, "HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out timeParsed);
                                }

                                lastKnownTime = datePart.Date + timeParsed.TimeOfDay;
                            }

                            // ✅ แล้วค่อยไปเช็ค probname
                            foreach (var prob in probMap)
                            {
                                if (prob.Name == "DISPENSE NOTE FAILED / ERRCODE:")
                                {
                                    // เช็คว่าบรรทัดขึ้นต้นด้วย pattern นี้
                                    int index = line.IndexOf(prob.Name);
                                    if (index >= 0)
                                    {
                                        DateTime trxTime = lastKnownTime ?? datePart.Date;
                                        string fullRemark = line.Substring(index).Trim(); // เก็บทั้งข้อความที่เหลือจาก index ไป
                                        string escapedRemark = fullRemark.Replace("\"", "\"\"");

                                        string newLine = $"{terminalId},\"{prob.Code}\",\"{escapedRemark}\",,,{trxTime:yyyy-MM-dd HH:mm:ss},1,{now:yyyy-MM-dd HH:mm:ss},{now:yyyy-MM-dd HH:mm:ss},operation";
                                        output.Add(newLine);
                                        break;
                                    }
                                }
                                else if (line.Contains(prob.Name))
                                {
                                    DateTime trxTime = lastKnownTime ?? datePart.Date;
                                    string remark = prob.Name.Replace("\"", "\"\"");
                                    string newLine = $"{terminalId},\"{prob.Code}\",\"{remark}\",,,{trxTime:yyyy-MM-dd HH:mm:ss},1,{now:yyyy-MM-dd HH:mm:ss},{now:yyyy-MM-dd HH:mm:ss},operation";
                                    output.Add(newLine);
                                    break;
                                }
                            }
                        }


                    }

                }

                string saveBaseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : txtFolderPath.Text;

                string csvPath = Path.Combine(saveBaseDir, $"reportcase_{dateStr}.csv");
                string xlsxPath = Path.Combine(saveBaseDir, $"reportcase_{dateStr}.xlsx");
                File.WriteAllLines(csvPath, output, Encoding.UTF8);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Report Case");

                    // Header
                    var headers = output[0].Split(',');
                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    }

                    // Data rows
                    for (int row = 1; row < output.Count; row++)
                    {
                        var fields = output[row].Split(',');
                        for (int col = 0; col < fields.Length; col++)
                        {
                            worksheet.Cell(row + 1, col + 1).Value = fields[col].Trim('"');
                        }
                    }

                    worksheet.Columns().AdjustToContents(); // ปรับขนาดคอลัมน์อัตโนมัติ
                    workbook.SaveAs(xlsxPath);
                }

                // ลบโฟลเดอร์ EJ/yyyyMMdd
                Directory.Delete(projectEJPath, true);

                MessageBox.Show($"✅ ส่งออก CSV แล้ว และลบไฟล์ EJ แล้ว:\n{csvPath}", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"เกิดข้อผิดพลาด:\n{ex.Message}", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDlg = new FolderBrowserDialog())
            {
                folderDlg.Description = "เลือกโฟลเดอร์ปลายทางสำหรับบันทึกไฟล์";
                folderDlg.SelectedPath = Properties.Settings.Default.LastFolderPath; // โหลด path ล่าสุด

                if (folderDlg.ShowDialog() == DialogResult.OK)
                {
                    txtFolderPath.Text = folderDlg.SelectedPath;
                    Properties.Settings.Default.LastFolderPath = folderDlg.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
        }
        private void cmbTerminalId_TextChanged(object sender, EventArgs e)
        {
            string input = cmbTerminalId.Text.Trim();

            // ✅ ให้เลือกเฉพาะตอนพิมพ์เกิน 2 ตัวอักษร
            if (input.Length < 11)
            {
                cmbTerminalId.SelectedItem = null;
                return;
            }

            foreach (string item in cmbTerminalId.Items)
            {
                string[] parts = item.Split('|');
                if (parts.Length != 2) continue;

                string fileName = parts[1]; // T641Bxxx_EJxxxxxx.zip
                string terminalId = fileName.Split('_')[0]; // T641Bxxx

                if (terminalId.StartsWith(input, StringComparison.OrdinalIgnoreCase))
                {
                    cmbTerminalId.SelectedItem = item;
                    cmbTerminalId.SelectionStart = cmbTerminalId.Text.Length;
                    cmbTerminalId.SelectionLength = 0;
                    break;
                }
            }
        }






    }
}
