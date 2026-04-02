using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using ClosedXML.Excel;
using Renci.SshNet;
using System.Collections.Generic;

namespace EJReadIssuesBAAC
{
    public partial class Form1 : Form
    {
        private CancellationTokenSource _autoJobCts;

        public Form1()
        {
            InitializeComponent();

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

            _autoJobCts = new CancellationTokenSource();
            StartAutoJob(_autoJobCts.Token);

            // โหลดรายการทันที
            LoadTerminalIdsFromSftp();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _autoJobCts.Cancel();
            base.OnFormClosed(e);
        }

        private async void StartAutoJob(CancellationToken token)
        {
            string runAtStr = ConfigurationManager.AppSettings["AutoRunTime"];
            if (!TimeSpan.TryParse(runAtStr, out TimeSpan runTime))
                runTime = new TimeSpan(4, 0, 0);

            while (!token.IsCancellationRequested)
            {
                DateTime now = DateTime.Now;
                DateTime nextRun = NextRunTime(now, runTime);
                TimeSpan wait = nextRun - now;

                try
                {
                    await Task.Delay(wait, token);
                }
                catch (TaskCanceledException) { break; }

                if (token.IsCancellationRequested) break;

                try
                {
                    string saveBaseDir = Properties.Settings.Default.LastFolderPath;
                    if (string.IsNullOrWhiteSpace(saveBaseDir) || !Directory.Exists(saveBaseDir))
                        saveBaseDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    var targetDate = DateTime.Today.AddDays(-1);
                    AutoExport.RunExport(saveBaseDir, targetDate);

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

        private static DateTime NextRunTime(DateTime now, TimeSpan targetTime)
        {
            var todayTarget = now.Date + targetTime;
            return (now < todayTarget) ? todayTarget : todayTarget.AddDays(1);
        }

        private bool IsTodayMode()
        {
            return datePicker.Value.Date == DateTime.Today;
        }

        private string BuildTodayCurrentTxtPath(string terminalId, DateTime targetDate)
        {
            return $"/data1/fileserverBAAC/EJ/Current/{terminalId}/{targetDate:yyyy}/{targetDate:MM}/EJ{targetDate:yyyyMMdd}.txt";
        }

        private void btnDownload_Click(object sender, EventArgs e)
        {
            string selected = cmbTerminalId.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(selected) || !selected.Contains("|"))
            {
                lblStatus.Text = "กรุณาเลือกรายการไฟล์";
                return;
            }

            var parts = selected.Split('|');
            string serverIp = parts[0].Trim();
            string itemName = parts[1].Trim(); // วันนี้ = TerminalId, ย้อนหลัง = zip file name

            string baseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : txtFolderPath.Text;

            try
            {
                using (var sftp = new SftpClient(serverIp, "root", "12qwaszx!@QWASZX"))
                {
                    sftp.Connect();

                    if (IsTodayMode())
                    {
                        string terminalId = itemName;
                        string remoteFile = BuildTodayCurrentTxtPath(terminalId, datePicker.Value.Date);
                        string localPath = Path.Combine(baseDir, $"EJ{datePicker.Value:yyyyMMdd}_{terminalId}.txt");

                        using (var fileStream = File.Create(localPath))
                        {
                            sftp.DownloadFile(remoteFile, fileStream);
                        }

                        lblStatus.ForeColor = Color.Green;
                        lblStatus.Text = $"✅ ดาวน์โหลด txt สำเร็จ: {terminalId} จาก {serverIp}";
                    }
                    else
                    {
                        string fileName = itemName;
                        string dateStr = datePicker.Value.ToString("yyyyMMdd");
                        string remoteFile = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/{fileName}";
                        string localPath = Path.Combine(baseDir, fileName);

                        using (var fileStream = File.Create(localPath))
                        {
                            sftp.DownloadFile(remoteFile, fileStream);
                        }

                        lblStatus.ForeColor = Color.Green;
                        lblStatus.Text = $"✅ ดาวน์โหลดสำเร็จ: {fileName} จาก {serverIp}";
                    }

                    sftp.Disconnect();
                }
            }
            catch (Exception ex)
            {
                lblStatus.ForeColor = Color.Red;
                lblStatus.Text = $"เกิดข้อผิดพลาดจาก {serverIp}:\n{ex.Message}";
            }
        }

        private void LoadTerminalIdsFromSftp()
        {
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

                        if (IsTodayMode())
                        {
                            // อ่าน Current/{TerminalId}/yyyy/MM/EJyyyyMMdd.txt
                            string currentRoot = "/data1/fileserverBAAC/EJ/Current/";
                            if (!sftp.Exists(currentRoot))
                            {
                                sftp.Disconnect();
                                continue;
                            }

                            var terminalDirs = sftp.ListDirectory(currentRoot)
                                                  .Where(f => f.IsDirectory && f.Name != "." && f.Name != "..")
                                                  .ToList();

                            var items = new List<string>();

                            foreach (var dir in terminalDirs)
                            {
                                string terminalId = dir.Name;
                                string txtPath = BuildTodayCurrentTxtPath(terminalId, datePicker.Value.Date);

                                if (sftp.Exists(txtPath))
                                {
                                    items.Add($"{serverIp}|{terminalId}");
                                }
                            }

                            cmbTerminalId.Items.AddRange(items.ToArray());
                        }
                        else
                        {
                            string dateStr = datePicker.Value.ToString("yyyyMMdd");
                            string remoteDir = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/";

                            if (!sftp.Exists(remoteDir))
                            {
                                sftp.Disconnect();
                                continue;
                            }

                            var zipFiles = sftp.ListDirectory(remoteDir)
                                               .Where(f => !f.IsDirectory && f.Name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                                               .Select(f => $"{serverIp}|{f.Name}");

                            cmbTerminalId.Items.AddRange(zipFiles.ToArray());
                        }

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
                MessageBox.Show("ไม่พบข้อมูลจากทุกเครื่อง", "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void datePicker_ValueChanged(object sender, EventArgs e)
        {
            LoadTerminalIdsFromSftp();
        }

        private void btnDownloadAll_Click(object sender, EventArgs e)
        {
            if (cmbTerminalId.Items.Count == 0)
            {
                MessageBox.Show("ไม่มีไฟล์ให้ดาวน์โหลด", "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string baseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : txtFolderPath.Text;

            string dateStr = datePicker.Value.ToString("yyyyMMdd");
            string localDir = Path.Combine(baseDir, $"EJ_{dateStr}");

            try
            {
                Directory.CreateDirectory(localDir);

                foreach (var item in cmbTerminalId.Items)
                {
                    string itemStr = item?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(itemStr) || !itemStr.Contains("|")) continue;

                    string[] parts = itemStr.Split('|');
                    if (parts.Length != 2) continue;

                    string ip = parts[0];
                    string value = parts[1];

                    using (var sftp = new SftpClient(ip, "root", "12qwaszx!@QWASZX"))
                    {
                        sftp.Connect();

                        if (IsTodayMode())
                        {
                            string terminalId = value;
                            string remoteFile = BuildTodayCurrentTxtPath(terminalId, datePicker.Value.Date);
                            string localTxtPath = Path.Combine(localDir, $"EJ{datePicker.Value:yyyyMMdd}_{terminalId}.txt");

                            using (var fileStream = File.Create(localTxtPath))
                            {
                                sftp.DownloadFile(remoteFile, fileStream);
                            }
                        }
                        else
                        {
                            string remoteDir = $"/data1/fileserverBAAC/EJ/Operation/{dateStr}/";
                            string fileName = value;
                            string remoteFile = remoteDir + fileName;
                            string safeIp = ip.Replace(".", "_");
                            string localZipPath = Path.Combine(localDir, $"{safeIp}_{fileName}");

                            using (var fileStream = File.Create(localZipPath))
                            {
                                sftp.DownloadFile(remoteFile, fileStream);
                            }

                            try
                            {
                                using (var archive = System.IO.Compression.ZipFile.OpenRead(localZipPath))
                                {
                                    foreach (var entry in archive.Entries)
                                    {
                                        if (!entry.FullName.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) continue;

                                        string newTxtName = Path.GetFileNameWithoutExtension(fileName) + ".txt";
                                        string targetPath = Path.Combine(localDir, newTxtName);

                                        entry.ExtractToFile(targetPath, overwrite: true);
                                    }
                                }

                                File.Delete(localZipPath);
                            }
                            catch (Exception innerEx)
                            {
                                MessageBox.Show($"❌ แตก zip ผิดพลาด: {fileName}\n{innerEx.Message}", "Extract Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                        sftp.Disconnect();
                    }
                }

                MessageBox.Show($"✅ ดาวน์โหลดทั้งหมดสำเร็จ\n\nที่เก็บ: {localDir}", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"เกิดข้อผิดพลาดระหว่างดาวน์โหลดทั้งหมด:\n\n{ex.Message}", "ผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportEJAndCsv_Click(object sender, EventArgs e)
        {
            string saveBaseDir = string.IsNullOrWhiteSpace(txtFolderPath.Text)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : txtFolderPath.Text;

            try
            {
                var targetDate = datePicker.Value;
                AutoExport.RunExport(saveBaseDir, targetDate);

                MessageBox.Show(
                    $"✅ ส่งออก CSV/XLSX แล้ว:\n{Path.Combine(saveBaseDir, $"reportcase_{targetDate:yyyyMMdd}.csv")}",
                    "สำเร็จ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
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
                folderDlg.SelectedPath = Properties.Settings.Default.LastFolderPath;

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

            if (input.Length < 3)
            {
                cmbTerminalId.SelectedItem = null;
                return;
            }

            foreach (string item in cmbTerminalId.Items)
            {
                if (string.IsNullOrWhiteSpace(item) || !item.Contains("|")) continue;

                string[] parts = item.Split('|');
                if (parts.Length != 2) continue;

                string key = parts[1]; // วันนี้ = terminalId, ย้อนหลัง = zip filename

                string compareValue = IsTodayMode()
                    ? key
                    : key.Split('_')[0];

                if (compareValue.StartsWith(input, StringComparison.OrdinalIgnoreCase))
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