using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;                               // ✅ เพิ่ม (ใช้ .Select/.Where)
using System.Threading;                         // ✅ เพิ่ม (background job)
using System.Threading.Tasks;                   // ✅ เพิ่ม (background job)
using System.Drawing;                           // ✅ เพิ่ม (Color)
using System.Windows.Forms;
using System.Configuration;                     // ✅ เพิ่ม (อ่าน App.config)
using ClosedXML.Excel;
using Renci.SshNet;

namespace EJReadIssuesBAAC
{
    public partial class Form1 : Form
    {
        // ✅ เพิ่ม: token สำหรับยกเลิก background job (auto run ตามเวลา)
        private CancellationTokenSource _autoJobCts;

        public Form1()
        {
            InitializeComponent();
            // เซ็ตค่า default ลง TextBox (LastFolderPath มาจาก Settings.settings)
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

            // ✅ เพิ่ม: เริ่ม background job สำหรับ auto-run ตามเวลาใน App.config
            _autoJobCts = new CancellationTokenSource();
            StartAutoJob(_autoJobCts.Token);
        }

        // ✅ เพิ่ม: ยกเลิก background job เมื่อปิดฟอร์ม
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _autoJobCts.Cancel();
            base.OnFormClosed(e);
        }

        // ✅ เพิ่ม: อ่านเวลา AutoRunTime จาก App.config -> รอจนถึง "เวลาถัดไป" -> ทำงาน -> วน
        private async void StartAutoJob(CancellationToken token)
        {
            // อ่านค่าเวลาจาก App.config (เช่น "04:00" หรือ "23:30")
            string runAtStr = ConfigurationManager.AppSettings["AutoRunTime"];  // จาก App.config
            if (!TimeSpan.TryParse(runAtStr, out TimeSpan runTime))
                runTime = new TimeSpan(4, 0, 0); // fallback = 04:00

            while (!token.IsCancellationRequested)
            {
                DateTime now = DateTime.Now;
                DateTime nextRun = NextRunTime(now, runTime); // เวลาถัดไปของวันนี้ ถ้าผ่านแล้วค่อยขยับไปวันถัดไป
                TimeSpan wait = nextRun - now;

                try
                {
                    await Task.Delay(wait, token);
                }
                catch (TaskCanceledException) { break; }

                if (token.IsCancellationRequested) break;

                try
                {
                    // ใช้ path ที่ฟอร์มตั้งไว้ (หรือ Desktop ถ้ายังไม่ตั้ง/ถูกลบ)
                    string saveBaseDir = Properties.Settings.Default.LastFolderPath;
                    if (string.IsNullOrWhiteSpace(saveBaseDir) || !Directory.Exists(saveBaseDir))
                        saveBaseDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    // ทำ "เมื่อวาน"
                    var targetDate = DateTime.Today.AddDays(-1);
                    AutoExport.RunExport(saveBaseDir, targetDate);

                    // แจ้งสถานะบนฟอร์ม
                    this.Invoke(() =>
                    {
                        lblStatus.ForeColor = Color.Green;
                        lblStatus.Text = $"[AUTO {runTime:hh\\:mm}] ✅ ส่งออก {targetDate:yyyyMMdd} สำเร็จ";
                    });
                }
                catch (Exception ex)
                {
                    this.Invoke(() =>
                    {
                        lblStatus.ForeColor = Color.Red;
                        lblStatus.Text = "[AUTO] ❌ " + ex.Message;
                    });
                }
            }
        }

        // ✅ เพิ่ม: helper หาค่า "เวลาถัดไป" ตาม runTime (เช่น 04:00)
        private static DateTime NextRunTime(DateTime now, TimeSpan targetTime)
        {
            var todayTarget = now.Date + targetTime;
            return (now < todayTarget) ? todayTarget : todayTarget.AddDays(1);
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

            // ใช้ path จาก txtFolderPath หรือ fallback เป็น Desktop
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
                    string itemStr = item?.ToString() ?? "";
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

                        // ดาวน์โหลด .zip
                        using (var fileStream = File.Create(localZipPath))
                        {
                            sftp.DownloadFile(remoteFile, fileStream);
                        }

                        try
                        {
                            // แตกไฟล์ .txt และตั้งชื่อใหม่ตาม .zip
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

                            // ลบ zip หลังแตกเสร็จ
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

        // ✅ เปลี่ยน: ปุ่ม manual สั้นลง เรียก AutoExport.RunExport(...) โดยใช้วันที่จาก datePicker
        private void btnExportEJAndCsv_Click(object sender, EventArgs e)
        {
            string saveBaseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : txtFolderPath.Text;

            try
            {
                var targetDate = datePicker.Value;   // manual ใช้วันที่จาก datePicker
                AutoExport.RunExport(saveBaseDir, targetDate);

                MessageBox.Show($"✅ ส่งออก CSV/XLSX แล้ว:\n{Path.Combine(saveBaseDir, $"reportcase_{targetDate:yyyyMMdd}.csv")}",
                    "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            // เดิมโชว์ว่า "เกิน 2 ตัวอักษร" แต่โค้ดใช้ 11 ตัว — คงตามโค้ดของคุณไว้
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
