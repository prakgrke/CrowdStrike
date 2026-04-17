using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CrowdStrikeTool
{
    public partial class MainForm : Form
    {
        private TextBox txtLog;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Button btnCSV;
        private Button btnFolder;
        private Button btnStart;
        private Label lblCsvPath;
        private Label lblFolderPath;

        private string csvPath = null;
        private string csFolder = null;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "CrowdStrike Deployment Tool";
            this.Size = new System.Drawing.Size(600, 400);
            this.StartPosition = FormStartPosition.CenterScreen;

            Label lblTitle = new Label
            {
                Text = "CrowdStrike Deployment Tool",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(300, 20),
                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblTitle);

            Label lblCsv = new Label { Text = "CSV File:", Location = new System.Drawing.Point(20, 70), Size = new System.Drawing.Size(80, 20) };
            this.Controls.Add(lblCsv);

            btnCSV = new Button { Text = "Select CSV", Location = new System.Drawing.Point(110, 67), Size = new System.Drawing.Size(100, 25) };
            btnCSV.Click += BtnCSV_Click;
            this.Controls.Add(btnCSV);

            lblCsvPath = new Label { Text = "", Location = new System.Drawing.Point(220, 70), Size = new System.Drawing.Size(350, 20), ForeColor = System.Drawing.Color.Gray };
            this.Controls.Add(lblCsvPath);

            Label lblFolder = new Label { Text = "CS Folder:", Location = new System.Drawing.Point(20, 110), Size = new System.Drawing.Size(80, 20) };
            this.Controls.Add(lblFolder);

            btnFolder = new Button { Text = "Select Folder", Location = new System.Drawing.Point(110, 107), Size = new System.Drawing.Size(100, 25) };
            btnFolder.Click += BtnFolder_Click;
            this.Controls.Add(btnFolder);

            lblFolderPath = new Label { Text = "", Location = new System.Drawing.Point(220, 110), Size = new System.Drawing.Size(350, 20), ForeColor = System.Drawing.Color.Gray };
            this.Controls.Add(lblFolderPath);

            btnStart = new Button
            {
                Text = "Start Deployment",
                Location = new System.Drawing.Point(200, 150),
                Size = new System.Drawing.Size(150, 35),
                BackColor = System.Drawing.Color.FromArgb(0, 120, 215),
                ForeColor = System.Drawing.Color.White
            };
            btnStart.Click += BtnStart_Click;
            this.Controls.Add(btnStart);

            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 210),
                Size = new System.Drawing.Size(540, 25)
            };
            this.Controls.Add(progressBar);

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(20, 250),
                Size = new System.Drawing.Size(540, 20)
            };
            this.Controls.Add(lblStatus);

            txtLog = new TextBox
            {
                Location = new System.Drawing.Point(20, 280),
                Size = new System.Drawing.Size(540, 80),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
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

        private void BtnFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    csFolder = dlg.SelectedPath;
                    lblFolderPath.Text = csFolder;
                    Log("Folder selected: " + csFolder);

                    string sensorPath = Path.Combine(csFolder, "WindowsSensor.exe");
                    if (File.Exists(sensorPath))
                    {
                        Log("WindowsSensor.exe found");
                    }
                    else
                    {
                        Log("WARNING: WindowsSensor.exe not found in folder");
                    }
                }
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(csvPath) || string.IsNullOrEmpty(csFolder))
            {
                MessageBox.Show("Please select both CSV file and CS Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sensorPath = Path.Combine(csFolder, "WindowsSensor.exe");
            if (!File.Exists(sensorPath))
            {
                MessageBox.Show("WindowsSensor.exe not found in selected folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnStart.Enabled = false;
            progressBar.Value = 0;

            try
            {
                var lines = File.ReadAllLines(csvPath);
                if (lines.Length < 2)
                {
                    MessageBox.Show("CSV file is empty or has no data rows", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    btnStart.Enabled = true;
                    return;
                }

                progressBar.Maximum = lines.Length - 1;
                Log("Starting deployment to " + (lines.Length - 1) + " machines...");

                string outputDir = @"C:\CS_Report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                Directory.CreateDirectory(outputDir);

                for (int i = 1; i < lines.Length; i++)
                {
                    var parts = lines[i].Split(',');
                    if (parts.Length < 3) continue;

                    string ip = parts[0].Trim();
                    string user = parts[1].Trim();
                    string pass = parts[2].Trim();

                    lblStatus.Text = "Processing: " + ip;
                    progressBar.Value = i;
                    Refresh();

                    try
                    {
                        Log("Connecting to " + ip + "...");
                        if (DeployToMachine(ip, user, pass, sensorPath))
                        {
                            Log("SUCCESS: " + ip);
                            File.AppendAllText(Path.Combine(outputDir, "result.txt"), ip + " SUCCESS" + Environment.NewLine);
                        }
                        else
                        {
                            Log("FAILED: " + ip);
                            File.AppendAllText(Path.Combine(outputDir, "result.txt"), ip + " FAILED" + Environment.NewLine);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("ERROR on " + ip + ": " + ex.Message);
                        File.AppendAllText(Path.Combine(outputDir, "result.txt"), ip + " FAILED - " + ex.Message + Environment.NewLine);
                    }
                }

                lblStatus.Text = "Completed!";
                Log("Deployment completed. Report saved to: " + outputDir);
                Process.Start(outputDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("ERROR: " + ex.Message);
            }
            finally
            {
                btnStart.Enabled = true;
            }
        }

        private bool DeployToMachine(string ip, string user, string pass, string sensorPath)
        {
            try
            {
                Log("  Enabling WinRM...");
                RunPowerShellCommand($@"
                    $sec = ConvertTo-SecureString '{pass}' -AsPlainText -Force
                    $cred = New-Object PSCredential('{user}', $sec)
                    Invoke-Command -ComputerName '{ip}' -Credential $cred -ScriptBlock {{
                        Set-Service -Name WinRM -StartupType Automatic -ErrorAction SilentlyContinue
                        Start-Service -Name WinRM -ErrorAction SilentlyContinue
                    }} -ErrorAction Stop
                ");

                Log("  Creating folder...");
                RunPowerShellCommand($@"
                    $sec = ConvertTo-SecureString '{pass}' -AsPlainText -Force
                    $cred = New-Object PSCredential('{user}', $sec)
                    Invoke-Command -ComputerName '{ip}' -Credential $cred -ScriptBlock {{
                        New-Item -ItemType Directory -Path C:\TempCS -Force | Out-Null
                    }} -ErrorAction Stop
                ");

                Log("  Copying file...");
                RunPowerShellCommand($@"
                    $sec = ConvertTo-SecureString '{pass}' -AsPlainText -Force
                    $cred = New-Object PSCredential('{user}', $sec)
                    $session = New-PSSession -ComputerName '{ip}' -Credential $cred -ErrorAction Stop
                    Copy-Item -Path '{sensorPath}' -Destination 'C:\TempCS\WindowsSensor.exe' -ToSession $session -ErrorAction Stop
                    Remove-PSSession $session
                ");

                Log("  Installing sensor...");
                RunPowerShellCommand($@"
                    $sec = ConvertTo-SecureString '{pass}' -AsPlainText -Force
                    $cred = New-Object PSCredential('{user}', $sec)
                    Invoke-Command -ComputerName '{ip}' -Credential $cred -ScriptBlock {{
                        Start-Process 'C:\TempCS\WindowsSensor.exe' -ArgumentList '/install /quiet' -Wait
                    }} -ErrorAction Stop
                ");

                return true;
            }
            catch (Exception ex)
            {
                Log("  Error: " + ex.Message);
                return false;
            }
        }

        private string RunPowerShellCommand(string script)
        {
            var psi = new ProcessStartInfo
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

            using (var process = Process.Start(psi))
            {
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrWhiteSpace(error))
                {
                    throw new Exception(error);
                }

                return output;
            }
        }

        private void Log(string message)
        {
            txtLog.AppendText(DateTime.Now.ToString("HH:mm:ss") + " - " + message + Environment.NewLine);
        }

        [STAThread]
        static void Main()
        {
            if (!IsAdministrator())
            {
                var exePath = Process.GetCurrentProcess().MainModule.FileName;
                var proc = new Process
                {
                    StartInfo = new ProcessStartInfo
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
}
