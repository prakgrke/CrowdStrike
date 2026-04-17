using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CrowdStrikeManager
{
    public partial class MainForm : Form
    {
        private TextBox txtLog;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgress;
        private Button btnCSV;
        private Button btnCSFolder;
        private Button btnScript;
        private Button btnStart;
        private Button btnStop;
        private ComboBox cmbVersion;
        private Label lblCsvPath;
        private Label lblFolderPath;
        private Label lblScriptPath;
        private DataGridView dgvResults;
        private CheckBox chkShowAll;
        private TextBox txtDomain;
        private TextBox txtAG;
        private TextBox txtCID;

        private string csvPath = null;
        private string csFolder = null;
        private string psScriptPath = null;
        private bool isRunning = false;
        private bool stopRequested = false;

        private List<MachineInfo> results = new List<MachineInfo>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "CrowdStrike Manager";
            this.Size = new System.Drawing.Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(900, 700);

            Label lblTitle = new Label
            {
                Text = "CrowdStrike Manager - Version Control & Deployment",
                Location = new System.Drawing.Point(20, 15),
                Size = new System.Drawing.Size(500, 25),
                Font = new System.Drawing.Font("Arial", 14, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblTitle);

            Label lblDomain = new Label { Text = "Domain (optional):", Location = new System.Drawing.Point(20, 55), Size = new System.Drawing.Size(120, 20) };
            this.Controls.Add(lblDomain);

            txtDomain = new TextBox { Location = new System.Drawing.Point(145, 52), Size = new System.Drawing.Size(300, 22), PlaceholderText = "DOMAIN or leave empty" };
            this.Controls.Add(txtDomain);

            Label lblCsv = new Label { Text = "CSV File:", Location = new System.Drawing.Point(20, 90), Size = new System.Drawing.Size(100, 20) };
            this.Controls.Add(lblCsv);

            btnCSV = new Button { Text = "Select CSV", Location = new System.Drawing.Point(120, 87), Size = new System.Drawing.Size(100, 25) };
            btnCSV.Click += BtnCSV_Click;
            this.Controls.Add(btnCSV);

            lblCsvPath = new Label { Text = "", Location = new System.Drawing.Point(230, 90), Size = new System.Drawing.Size(600, 20), ForeColor = System.Drawing.Color.Gray };
            this.Controls.Add(lblCsvPath);

            Label lblFolder = new Label { Text = "CS Folder:", Location = new System.Drawing.Point(20, 125), Size = new System.Drawing.Size(100, 20) };
            this.Controls.Add(lblFolder);

            btnCSFolder = new Button { Text = "Select Folder", Location = new System.Drawing.Point(120, 122), Size = new System.Drawing.Size(100, 25) };
            btnCSFolder.Click += BtnCSFolder_Click;
            this.Controls.Add(btnCSFolder);

            lblFolderPath = new Label { Text = "", Location = new System.Drawing.Point(230, 125), Size = new System.Drawing.Size(600, 20), ForeColor = System.Drawing.Color.Gray };
            this.Controls.Add(lblFolderPath);

            Label lblScript = new Label { Text = "PS Script:", Location = new System.Drawing.Point(20, 160), Size = new System.Drawing.Size(100, 20) };
            this.Controls.Add(lblScript);

            btnScript = new Button { Text = "Select Script", Location = new System.Drawing.Point(120, 157), Size = new System.Drawing.Size(100, 25) };
            btnScript.Click += BtnScript_Click;
            this.Controls.Add(btnScript);

            lblScriptPath = new Label { Text = "", Location = new System.Drawing.Point(230, 160), Size = new System.Drawing.Size(600, 20), ForeColor = System.Drawing.Color.Gray };
            this.Controls.Add(lblScriptPath);

            Label lblVersion = new Label { Text = "PS Script:", Location = new System.Drawing.Point(20, 195), Size = new System.Drawing.Size(100, 20) };
            this.Controls.Add(lblVersion);

            cmbVersion = new ComboBox { Location = new System.Drawing.Point(120, 192), Size = new System.Drawing.Size(200, 22), DropDownStyle = ComboBoxStyle.DropDownList };
            this.Controls.Add(cmbVersion);

            Label lblAG = new Label { Text = "Agent Group:", Location = new System.Drawing.Point(350, 195), Size = new System.Drawing.Size(80, 20) };
            this.Controls.Add(lblAG);

            txtAG = new TextBox { Location = new System.Drawing.Point(435, 192), Size = new System.Drawing.Size(150, 22), PlaceholderText = "Agent Group Name" };
            this.Controls.Add(txtAG);

            Label lblCID = new Label { Text = "CID:", Location = new System.Drawing.Point(610, 195), Size = new System.Drawing.Size(30, 20) };
            this.Controls.Add(lblCID);

            txtCID = new TextBox { Location = new System.Drawing.Point(645, 192), Size = new System.Drawing.Size(215, 22), PlaceholderText = "CrowdStrike CID" };
            this.Controls.Add(txtCID);

            chkShowAll = new CheckBox { Text = "Show All Machines", Location = new System.Drawing.Point(120, 225), Size = new System.Drawing.Size(150, 20), Checked = true };
            this.Controls.Add(chkShowAll);

            btnStart = new Button
            {
                Text = "Start Scan & Deploy",
                Location = new System.Drawing.Point(300, 225),
                Size = new System.Drawing.Size(150, 35),
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White
            };
            btnStart.Click += BtnStart_Click;
            this.Controls.Add(btnStart);

            btnStop = new Button
            {
                Text = "Stop",
                Location = new System.Drawing.Point(460, 225),
                Size = new System.Drawing.Size(80, 35),
                BackColor = System.Drawing.Color.FromArgb(200, 50, 50),
                ForeColor = System.Drawing.Color.White,
                Enabled = false
            };
            btnStop.Click += BtnStop_Click;
            this.Controls.Add(btnStop);

            progressBar = new ProgressBar { Location = new System.Drawing.Point(20, 275), Size = new System.Drawing.Size(840, 20) };
            this.Controls.Add(progressBar);

            lblProgress = new Label { Text = "", Location = new System.Drawing.Point(20, 300), Size = new System.Drawing.Size(840, 20) };
            this.Controls.Add(lblProgress);

            dgvResults = new DataGridView
            {
                Location = new System.Drawing.Point(20, 325),
                Size = new System.Drawing.Size(840, 285),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ReadOnly = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                BackgroundColor = System.Drawing.Color.White
            };
            dgvResults.Columns.Add("IP", "IP Address");
            dgvResults.Columns.Add("Hostname", "Hostname");
            dgvResults.Columns.Add("CSVersion", "CS Version");
            dgvResults.Columns.Add("Status", "Status");
            dgvResults.Columns.Add("Action", "Action Taken");
            dgvResults.Columns.Add("Details", "Details");
            this.Controls.Add(dgvResults);

            lblStatus = new Label { Text = "Ready", Location = new System.Drawing.Point(20, 620), Size = new System.Drawing.Size(840, 20) };
            this.Controls.Add(lblStatus);

            txtLog = new TextBox
            {
                Location = new System.Drawing.Point(20, 645),
                Size = new System.Drawing.Size(840, 35),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Height = 35
            };
            this.Controls.Add(txtLog);
        }

        private void BtnCSV_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "CSV Files (*.csv)|*.csv";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    csvPath = dlg.FileName;
                    lblCsvPath.Text = csvPath;
                    Log("CSV file selected: " + Path.GetFileName(csvPath));
                }
            }
        }

        private void BtnCSFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    csFolder = dlg.SelectedPath;
                    lblFolderPath.Text = csFolder;
                    Log("CS folder selected: " + csFolder);

                    string[] certFiles = Directory.GetFiles(csFolder, "*.cer");
                    string[] pfxFiles = Directory.GetFiles(csFolder, "*.pfx");
                    string[] psFiles = Directory.GetFiles(csFolder, "*.ps1");

                    Log($"Found: {certFiles.Length} certificates, {pfxFiles.Length} PFX files, {psFiles.Length} scripts");

                    cmbVersion.Items.Clear();
                    foreach (var file in psFiles)
                    {
                        string name = Path.GetFileName(file);
                        cmbVersion.Items.Add(name);
                        if (cmbVersion.Items.Count == 1)
                        {
                            cmbVersion.SelectedIndex = 0;
                            psScriptPath = file;
                            lblScriptPath.Text = file;
                        }
                    }
                }
            }
        }

        private void BtnScript_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "PowerShell Scripts (*.ps1)|*.ps1";
                dlg.InitialDirectory = csFolder;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    psScriptPath = dlg.FileName;
                    lblScriptPath.Text = psScriptPath;
                    Log("PowerShell script selected: " + Path.GetFileName(psScriptPath));
                }
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(csvPath) || string.IsNullOrEmpty(csFolder) || string.IsNullOrEmpty(psScriptPath))
            {
                MessageBox.Show("Please select CSV file, CS folder, and PowerShell script", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbVersion.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a PowerShell script version", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbVersion.SelectedItem != null)
            {
                string selectedScript = cmbVersion.SelectedItem.ToString();
                psScriptPath = Path.Combine(csFolder, selectedScript);
                if (!File.Exists(psScriptPath))
                {
                    MessageBox.Show("Selected script not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            isRunning = true;
            stopRequested = false;
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            btnCSV.Enabled = false;
            btnCSFolder.Enabled = false;
            btnScript.Enabled = false;

            results.Clear();
            dgvResults.Rows.Clear();

            Log("Starting CrowdStrike Manager...");
            progressBar.Value = 0;

            try
            {
                var lines = File.ReadAllLines(csvPath);
                if (lines.Length < 2)
                {
                    MessageBox.Show("CSV file is empty or has no data rows", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                progressBar.Maximum = lines.Length - 1;
                Log($"Found {lines.Length - 1} machines to process");

                string outputDir = @"C:\CS_Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                Directory.CreateDirectory(outputDir);

                for (int i = 1; i < lines.Length; i++)
                {
                    if (stopRequested) break;

                    var parts = lines[i].Split(',');
                    if (parts.Length < 3) continue;

                    string ip = parts[0].Trim();
                    string user = parts[1].Trim();
                    string pass = parts[2].Trim();
                    string domain = txtDomain.Text.Trim();

                    if (!string.IsNullOrEmpty(domain) && !user.Contains("\\"))
                    {
                        user = domain + "\\" + user;
                    }

                    lblProgress.Text = $"Processing: {ip} ({i}/{lines.Length - 1})";
                    lblStatus.Text = $"Processing: {ip}";
                    progressBar.Value = i;
                    Refresh();

                    try
                    {
                        Log($"Connecting to {ip}...");
                        MachineInfo info = ProcessMachine(ip, user, pass);
                        results.Add(info);

                        Invoke(new Action(() =>
                        {
                            dgvResults.Rows.Add(info.IP, info.Hostname, info.CSVersion, info.Status, info.Action, info.Details);
                            if (!chkShowAll.Checked && info.Status == "Up to Date")
                            {
                                dgvResults.Rows[dgvResults.Rows.Count - 1].Visible = false;
                            }
                        }));
                    }
                    catch (Exception ex)
                    {
                        Log($"ERROR on {ip}: {ex.Message}");
                        MachineInfo errorInfo = new MachineInfo { IP = ip, Status = "Error", Details = ex.Message };
                        results.Add(errorInfo);
                        Invoke(new Action(() => dgvResults.Rows.Add(ip, "N/A", "N/A", "Error", "N/A", ex.Message)));
                    }
                }

                string report = GenerateReport(outputDir);
                lblStatus.Text = "Completed! Report saved to: " + report;
                Log("Deployment completed. Report: " + report);
                System.Diagnostics.Process.Start(outputDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("ERROR: " + ex.Message);
            }
            finally
            {
                isRunning = false;
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                btnCSV.Enabled = true;
                btnCSFolder.Enabled = true;
                btnScript.Enabled = true;
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            stopRequested = true;
            Log("Stop requested...");
        }

        private MachineInfo ProcessMachine(string ip, string user, string pass)
        {
            MachineInfo info = new MachineInfo { IP = ip };

            Log($"  Getting system info from {ip}...");
            string hostname = RunPowerShellRemote(ip, user, pass, "$env:COMPUTERNAME");
            info.Hostname = hostname.Trim();

            Log($"  Checking CrowdStrike version on {ip}...");
            string csVersion = RunPowerShellRemote(ip, user, pass, @"
                $cs = Get-ItemProperty 'HKLM:\SOFTWARE\CrowdStrike\{9b9b93b8-79e0-11ed-82be-b9b31b4e1f7b}' -ErrorAction SilentlyContinue
                if ($cs) { $cs.Version } else { 'Not Installed' }
            ");
            info.CSVersion = csVersion.Trim();

            Log($"  Checking installed software on {ip}...");
            string software = RunPowerShellRemote(ip, user, pass, @"
                Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue | 
                Select-Object DisplayName, DisplayVersion | ConvertTo-Json -Compress
            ");

            Log($"  Checking installed patches on {ip}...");
            string patches = RunPowerShellRemote(ip, user, pass, @"
                Get-HotFix | Select-Object HotFixID, InstalledOn | ConvertTo-Json -Compress
            ");

            if (info.CSVersion == "Not Installed")
            {
                Log($"  CrowdStrike NOT installed on {ip}");
                info.Status = "Not Installed";
                info.Action = "Installing";

                InstallCertificates(ip, user, pass);
                InstallCrowdStrike(ip, user, pass);
                info.Details = "Certificates + Falcon Sensor installed";
            }
            else
            {
                Log($"  Current CS Version: {info.CSVersion}");
                info.Status = "Up to Date";
                info.Action = "No Action Required";
                info.Details = "Same version already installed";
            }

            string outputDir = @"C:\CS_Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string machineReport = Path.Combine(outputDir, $"{ip}_report.txt");
            StringBuilder report = new StringBuilder();
            report.AppendLine($"=== Machine Report: {ip} ===");
            report.AppendLine($"Hostname: {info.Hostname}");
            report.AppendLine($"CrowdStrike Version: {info.CSVersion}");
            report.AppendLine($"Status: {info.Status}");
            report.AppendLine($"Action: {info.Action}");
            report.AppendLine($"Details: {info.Details}");
            report.AppendLine();
            report.AppendLine("=== Installed Software ===");
            report.AppendLine(software);
            report.AppendLine();
            report.AppendLine("=== Installed Patches ===");
            report.AppendLine(patches);
            File.WriteAllText(machineReport, report.ToString());

            return info;
        }

        private void InstallCertificates(string ip, string user, string pass)
        {
            Log($"  Installing certificates on {ip}...");

            string[] certFiles = Directory.GetFiles(csFolder, "*.cer");
            string[] pfxFiles = Directory.GetFiles(csFolder, "*.pfx");

            foreach (var cert in certFiles)
            {
                string certName = Path.GetFileName(cert);
                Log($"    Installing .cer: {certName}");
                string script = $@"
                    try {{
                        $bytes = [System.IO.File]::ReadAllBytes('{cert.Replace("\\", "\\\\")}')
                        $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('TrustedPublisher', 'LocalMachine')
                        $certStore.Open('ReadWrite')
                        $certStore.Add({{ 
                            New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($bytes) 
                        }})
                        $certStore.Close()
                        Write-Output 'Certificate {certName} installed successfully'
                    }} catch {{
                        Write-Output 'Error installing {certName}: $_'
                    }}
                ";
                string result = RunPowerShellRemote(ip, user, pass, script);
                Log($"    Result: {result.Trim()}");
            }

            foreach (var pfx in pfxFiles)
            {
                string pfxName = Path.GetFileName(pfx);
                Log($"    Installing .pfx: {pfxName}");
                string script = $@"
                    try {{
                        $bytes = [System.IO.File]::ReadAllBytes('{pfx.Replace("\\", "\\\\")}')
                        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($bytes)
                        $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('TrustedPublisher', 'LocalMachine')
                        $certStore.Open('ReadWrite')
                        $certStore.Add($cert)
                        $certStore.Close()
                        Write-Output 'PFX certificate {pfxName} installed successfully'
                    }} catch {{
                        Write-Output 'Error installing {pfxName}: $_'
                    }}
                ";
                string result = RunPowerShellRemote(ip, user, pass, script);
                Log($"    Result: {result.Trim()}");
            }
        }

        private void InstallCrowdStrike(string ip, string user, string pass)
        {
            Log($"  Installing Falcon Sensor on {ip}...");

            string ag = txtAG.Text.Trim();
            string cid = txtCID.Text.Trim();

            string[] exeFiles = Directory.GetFiles(csFolder, "*.exe");
            string falconExe = null;
            
            foreach (var exe in exeFiles)
            {
                string exeName = Path.GetFileName(exe).ToLower();
                if (exeName.Contains("sensor") || exeName.Contains("falcon") || exeName.Contains("windows"))
                {
                    falconExe = exe;
                    break;
                }
            }

            if (falconExe == null && exeFiles.Length > 0)
            {
                falconExe = exeFiles[0];
            }

            if (falconExe == null)
            {
                Log($"  ERROR: Falcon sensor EXE not found in CS folder!");
                return;
            }

            string falconExeName = Path.GetFileName(falconExe);
            Log($"  Found Falcon EXE: {falconExeName}");

            string script;
            
            if (!string.IsNullOrEmpty(psScriptPath) && File.Exists(psScriptPath))
            {
                string scriptContent = File.ReadAllText(psScriptPath);
                
                script = $@"
                    New-Item -ItemType Directory -Path C:\TempCS -Force | Out-Null
                    Set-Content -Path C:\TempCS\falcon.ps1 -Value @'
{scriptContent}
'@
                    Write-Output 'Executing Falcon installation...'
                    & C:\TempCS\falcon.ps1
                ";
            }
            else
            {
                string psParams = "";
                if (!string.IsNullOrEmpty(ag))
                    psParams = $"/install /quiet AG='{ag}'";
                else if (!string.IsNullOrEmpty(cid))
                    psParams = $"/install /quiet CID='{cid}'";
                else
                    psParams = "/install /quiet";

                script = $@"
                    New-Item -ItemType Directory -Path C:\TempCS -Force | Out-Null
                    Write-Output 'Installing Falcon sensor with params: {psParams}'
                    Start-Process 'C:\TempCS\{falconExeName}' -ArgumentList '{psParams}' -Wait
                    Write-Output 'Falcon installation completed'
                ";
            }
            
            string result = RunPowerShellRemote(ip, user, pass, script);
            Log($"  Result: {result.Trim()}");
        }

        private string RunPowerShellRemote(string ip, string user, string pass, string command)
        {
            string script = $@"
                $sec = ConvertTo-SecureString '{pass}' -AsPlainText -Force
                $cred = New-Object PSCredential('{user}', $sec)
                $result = Invoke-Command -ComputerName '{ip}' -Credential $cred -ScriptBlock {{
                    {command}
                }} -ErrorAction Stop 2>&1 | Out-String
                $result
            ";

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-ExecutionPolicy Bypass -NoProfile -Command \"{script.Replace("\"", "\\\"")}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            };

            using (var process = System.Diagnostics.Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit(60000);

                if (!string.IsNullOrWhiteSpace(error) && !error.Contains("WARNING"))
                {
                    throw new Exception(error);
                }

                return output;
            }
        }

        private string GenerateReport(string outputDir)
        {
            StringBuilder report = new StringBuilder();
            report.AppendLine("=== CrowdStrike Manager Report ===");
            report.AppendLine($"Generated: {DateTime.Now}");
            report.AppendLine($"CSV File: {csvPath}");
            report.AppendLine($"CS Folder: {csFolder}");
            report.AppendLine($"PowerShell Script: {psScriptPath}");
            report.AppendLine();
            report.AppendLine("=== Summary ===");

            int notInstalled = 0, updated = 0, upToDate = 0, errors = 0;

            foreach (var info in results)
            {
                switch (info.Status)
                {
                    case "Not Installed": notInstalled++; break;
                    case "Updated": updated++; break;
                    case "Up to Date": upToDate++; break;
                    case "Error": errors++; break;
                }
            }

            report.AppendLine($"Machines Without CrowdStrike: {notInstalled}");
            report.AppendLine($"Machines Updated: {updated}");
            report.AppendLine($"Machines Already Up to Date: {upToDate}");
            report.AppendLine($"Errors: {errors}");
            report.AppendLine();
            report.AppendLine("=== Detailed Results ===");
            report.AppendLine();

            foreach (var info in results)
            {
                report.AppendLine($"IP: {info.IP}");
                report.AppendLine($"  Hostname: {info.Hostname}");
                report.AppendLine($"  CS Version: {info.CSVersion}");
                report.AppendLine($"  Status: {info.Status}");
                report.AppendLine($"  Action: {info.Action}");
                report.AppendLine($"  Details: {info.Details}");
                report.AppendLine();
            }

            string reportPath = Path.Combine(outputDir, "summary.txt");
            File.WriteAllText(reportPath, report.ToString());
            return reportPath;
        }

        private void Log(string message)
        {
            Invoke(new Action(() =>
            {
                txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") + " - " + message + Environment.NewLine);
            }));
        }

        [STAThread]
        static void Main()
        {
            if (!IsAdministrator())
            {
                var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                var proc = new System.Diagnostics.Process
                {
                    StartInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = exePath,
                        UseShellExecute = true,
                        Verb = "runas"
                    }
                };
                proc.Start();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private static bool IsAdministrator()
        {
            using (var identity = System.Security.Principal.WindowsIdentity.GetCurrent())
            {
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
            }
        }
    }

    class MachineInfo
    {
        public string IP { get; set; }
        public string Hostname { get; set; }
        public string CSVersion { get; set; }
        public string Status { get; set; }
        public string Action { get; set; }
        public string Details { get; set; }
    }
}
