using Encryption;
using ImpersonationsTools;
using Interoperability;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowHandle;
using WindowsUtilities;

namespace AutomateDownloader
{
    class NCMForm : Form
    {
        #region Form Objects
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.TextBox pathTextBox;
        private System.Windows.Forms.TextBox ipTextBox;
        private System.Windows.Forms.TextBox unTextBox;
        private System.Windows.Forms.TextBox passTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox rdpCheckBox;
        private System.Windows.Forms.GroupBox rdpBox1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.TextBox heightBox;
        private System.Windows.Forms.TextBox widthBox;
        private System.Windows.Forms.TextBox topBox;
        private System.Windows.Forms.TextBox leftBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private Button killButtton;
        private TextBox textBox1;
        private TextBox textBox2;
        private GroupBox groupBox1;
        private Label label1;
        private TextBox firstClientIndexBox;
        private TextBox numClTextBox;
        private Label label5;
        private Button button3;
        private Button button4;
        private Button button5;
        private Button button6;
        private Button button7;
        private Button button8;
        private TextBox destinationPathBox;
        private TextBox sourcePathBox;
        private Label label11;
        private Label label10;
        private Label label12;
        private TextBox parallelBox;
        private TextBox mcpPathBox;
        private Label label13;
        private Button button9;
        private Button button2;
        private Button button11;
        private Button button12;
        private CheckBox checkBox2;
        private TextBox textBox3;
        private TextBox vpnPassBox;
        private Button button13;
        private ListBox listBox1;
        private CheckBox includeNonClientsBox;
        private Button button14;
        private Button button10;
        #endregion

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.Run(new NCMForm());
            //Done by Muresan Radu-Adrian MURA02 20200716
        }

        public NCMForm()
        {
            CloseConflictingProcesses("NCM_Downloader");

            InitializeComponent();

            string configPath = string.Empty;
            if (File.Exists(Path.Combine(Application.StartupPath, "NCM_Downloader.ini")))
            {
                configPath = Path.Combine(Application.StartupPath, "NCM_Downloader.ini");
            }
            else if (File.Exists(Path.Combine(Application.StartupPath, "\\NCM_Downloader.exe.ini")))
            {
                configPath = Path.Combine(Application.StartupPath, "\\NCM_Downloader.exe.ini");
            }
            if (configPath != string.Empty)
            {
                GetInitialValues(configPath);
            }

            SetTooltips();

            RenewIpsOrInit(false);

            button1.Click += new EventHandler(Button1_Click);

            this.DoubleClick += new EventHandler(Form1_DoubleClick);

            this.Controls.Add(firstClientIndexBox);

            computerName = Environment.MachineName;

            selectiveFolders = new List<string>
                {
                    "Gracs",
                    "ScriptLib",
                    "ScriptAct",
                    "TEXTBIB",
                    "Packages"
                };

            statusLabel.MaximumSize = new System.Drawing.Size(280, 0);
            statusLabel.AutoSize = true;
            Console.WriteLine(statusLabel.Size.Height / statusLabel.Font.Height);
        }

        #region init
        private List<string> ipList = new List<string>();
        private List<string> sdList = new List<string>();
        private string computerName = string.Empty;
        private readonly List<string> selectiveFolders = new List<string>();
        private readonly List<string> processes = new List<string>
            {
                //"CCAgent",
                //"CCDBUtils",
                //"CCEClient_x64",
                //"CCEServer_x64",
                //"CCProjectMgr",
                //"CCRemoteService",
                //"CCPackageMgr",
                //"CCUCSurrogate",
                //"CCPerfMon",
                //"CCDeltaLoader",
                //"CCLicenseService",
                //"CCOnScreenKeyboard",
                //"CCNSInfo2Provider",
                //"script",
                //"SDiagRT",
                //"RedundancyState",
                //"PassDBRT"

                "apdiag",
                "asosheartbeatx",
                "AutoStartRT",
                "BildAnw",
                "CAArchiveManagerBackupCopyX",
                "CADeleteWinCCMessageQueueing",
                "CAReportMergerX",
                "CC_DPSRV",
                "CC_MShh",
                "ccaeprovider",
                "CCAgent.exe",
                "CCAlgIAlarmDataCollector",
                "CCAlgIAlarmDataProxyServer",
                "CcAlgRtServer",
                "CCArchiveConnMon",
                "CCArchiveManager",
                "CCAuthorInformation",
                "CCAuditProviderSrv.exe",
                "CCAuditTrailServer.exe",
                "CCAuditViewer",
                "CCAuditDCV",
                "CCClientAPIs",
                "CCCloudConnect",
                "CCConfigServerMemory",
                "CCConfigStudio",
                "CCConfigStudio_x64",
                "CCCSigRTServer",
                "CCDBUtils.exe",
                "CCDeltaLoader",
                "CCDiagnosisView",
                "CCDiagRTServer",
                "CCDMClientHelper",
                "CCDmIVarProxyServer",
                "CCDmRtChannelHost",
                "CCDmRuntimePersistence",
                "CCEClient_x64",
                "CCEClient",
                "CCEmergencyWatchRTServer",
                "CCEServer_x64",
                "CCEServer",
                "CCLBMRTServer",
                "CCLicenseService",
                "CCMcpAutServer",
                "CCMigrator",
                "CCNSInfo2Provider",
                "CCOnlCmp",
                "CCOPCConfigPerm",
                "CCPackageMgr",
                "CCPerfMon",
                "CcPJob",
                "CCProfileServer" ,
                "CCProgressDlg",
                "CCProjectDuplicator",
                "CCProjectMgr",
                "CCPtmRTServer",
                "CCRedCodi",
                "CCRedundancyAgent",
                "CCRemoteService.exe",
                "CCRtsLoader",
                "CCRtsLoader_x64",
                "CCRunDTSPackage",
                "CCRunRedCodiCS",
                "CCScriptEditor",
                "CCSplash",
                "CCSsmRTServer",
                "CCSyncAgent",
                "CCSystemDiagnosticsHost",
                "CCSysTray",
                "CCSysTrayProxy",
                "CCTextDistributor",
                "CCTextServer",
                "CCTlgArchiveClient",
                "CCTlgServer",
                "CCTMConfiguration",
                "CCTMTimeSync",
                "CCTMTimeSyncServer",
                "CCTMTimeSyncServerV5",
                "CCTMTimeSyncV5",
                "CCTTRTServer",
                "CCTxtProxy",
                "CCUAEditor",
                "CCUAIUABasicProxyServer",
                "CCUAIUABasicStubServer",
                "CCUASetupDCOM",
                "CCUCSurrogate",
                "CCUsrAcv",
                "CCWinCCMTEditor",
                "CCWinCCStart",
                "CCWriteArchiveServer",
                "CCXREF.Presentation.Editor",
                "ChannelWrapperCS",
                "DdeServ",
                "DiagServ",
                "Grafexe",
                "gsccs",
                "gscrt",
                "HMRT",
                "HornCS",
                "imserverx",
                "LBMS",
                "MessageCorrectorx",
                "MSCS",
                "OPCTags",
                "OpcUaServerWinCC.exe",
                "osltmhandlerx",
                "osstatemachinex",
                "PassCS",
                "PassDBRT",
                "PassReg",
                "PDLCSApiTest",
                "PdlRT",
                "PdlServ",
                "Print1",
                "PrintIt",
                "Projecteditor",
                "ProtCS",
                "PrtScr",
                "PTMCS",
                "RebootX",
                "Rd1CS",
                "RedundancyControl",
                "RedundancyState",
                "RptRT",
                "s7asysvx",
                "s7aversx",
                "s7dosTraceControlPanelx",
                "s7oiehsx" ,
                "s7otbxsx" ,
                "s7sninsx",
                "S7TraceServiceX.exe",
                "s7tgtopx",
                "s7ubTstx",
                "sfcrt",
                "SAMExportToolx",
                "script",
                "SCSDialogX",
                "SCSDistServiceX.exe",
                "SCSFsX",
                "SCSMonitor",
                "SCSMX.exe",
                "SDiagRT",
                "setupdcom",
                "Siemens.Automation.FrameApplicationLoader",
                "Siemens.Automation.ObjectFrame.FileStorage.Server",
                "Siemens.Automation.Portal",
                "sn_cp1612",
                "ssc_visual",
                "SSMSettings",
                "TagAnw",
                "textbib",
                "TextLibrary",
                "TlgCS",
                "TouchInputPC",
                "TrendOnl",
                "trv",
                "TTEdit",
                "WebNavigatorRT",
                "WinCCChnDiag",
                "WinCCExplorer",
                "WinCCWebLicense",
                "LBMCS",
                "CCOnScreenKeyboard",
                "CCCAPHServer",
                "CalendarAccessProvider.exe",
                "WinCC_Calendar_Server",
                "WinCC_Calendar_Viewer",
                "PlantIntelligenceService.exe",
                "PMNSInfo2Provider",
                "CCSESRTSrv",
                "GfxRTS",
                "CCKeyboardHook",
                "CCProjectMgr.exe",
                //"SCSMX.exe",// servicename = "SCSMonitor" restartprio = "1" />
                //"SCSDistServiceX.exe",// servicename = "SCS Distribution Service" restartprio = "2" />
                //"CCAgent.exe",// servicename = "CCAgent" restartprio = "3" />     
                //"S7TraceServiceX.exe",// servicename = "S7TraceServiceX" restartprio = "4" />     
                //"CCDBUtils.exe",// servicename = "CCDBUtils" restartprio = "5" />     
                //"CCRemoteService.exe",// servicename = "CCRemoteService" restartprio = "6" />     
                //"OpcUaServerWinCC.exe",// servicename = "OpcUaServerWinCC" restartprio = "7" />
                //"CCAuditTrailServer.exe",// restartprio = "8" servicename = "CCAuditTrailServer" />
                //"CCAuditProviderSrv.exe",// restartprio = "9" servicename = "CCAuditProviderSrv" />
                //"CalendarAccessProvider.exe",// restartprio = "10" servicename = "CalendarAccessProvider" />     
                //"CCProjectMgr.exe",// restartprio = "11" servicename = "CCProjectMgr" />
                //"PlantIntelligenceService.exe",// restartprio = "12" servicename = "PerformanceMonitor Service" />
                //"sqlservr", //"MSSQL$WINCC",
            };
        private readonly List<string> rebootProcesses = new List<string>
            {
                //"sqlservr", //"MSSQL$WINCC",
                "SCSMX.exe",// servicename = "SCSMonitor" restartprio = "1" />
                "SCSDistServiceX.exe",// servicename = "SCS Distribution Service" restartprio = "2" />
                "CCAgent.exe",// servicename = "CCAgent" restartprio = "3" />     
                "S7TraceServiceX.exe",// servicename = "S7TraceServiceX" restartprio = "4" />     
                "CCDBUtils.exe",// servicename = "CCDBUtils" restartprio = "5" />     
                "CCRemoteService.exe",// servicename = "CCRemoteService" restartprio = "6" />     
                "OpcUaServerWinCC.exe",// servicename = "OpcUaServerWinCC" restartprio = "7" />
                "CCAuditTrailServer.exe",// restartprio = "8" servicename = "CCAuditTrailServer" />
                "CCAuditProviderSrv.exe",// restartprio = "9" servicename = "CCAuditProviderSrv" />
                "CalendarAccessProvider.exe",// restartprio = "10" servicename = "CalendarAccessProvider" />     
                "CCProjectMgr.exe",// restartprio = "11" servicename = "CCProjectMgr" />
                "PlantIntelligenceService.exe"// restartprio = "12" servicename = "PerformanceMonitor Service" />
            };
        private readonly string passPhrase = "Hush, little baby, don't say a word " +
            "And never mind that noise you heard " +
            "It's just the beasts under your bed " +
            "In your closet, in your head";

        private void GetInitialValues(string path)
        {
            var configFile = File.ReadLines(path);
            var fileLen = configFile.Count();
            var de = "";

            pathTextBox.Text = configFile.ElementAt(0);
            if (fileLen >= 2) firstClientIndexBox.Text = configFile.ElementAt(1);
            if (fileLen >= 3) numClTextBox.Text = configFile.ElementAt(2);
            if (fileLen >= 4) ipTextBox.Text = configFile.ElementAt(3);
            if (fileLen >= 5) unTextBox.Text = configFile.ElementAt(4);
            if (fileLen >= 6)
            {
                if (configFile.ElementAt(5) != "")
                {
                    de = StringCipher.Decrypt(configFile.ElementAt(5), passPhrase);
                    passTextBox.Text = de;
                }
            }
            if (fileLen >= 7) rdpCheckBox.Checked = Convert.ToBoolean(configFile.ElementAt(6));
            if (fileLen >= 8) widthBox.Text = configFile.ElementAt(7);
            if (fileLen >= 9) heightBox.Text = configFile.ElementAt(8);
            if (fileLen >= 10) leftBox.Text = configFile.ElementAt(9);
            if (fileLen >= 11) topBox.Text = configFile.ElementAt(10);
            if (fileLen >= 12) parallelBox.Text = configFile.ElementAt(11);
            if (fileLen >= 13) sourcePathBox.Text = configFile.ElementAt(12);
            if (fileLen >= 14) destinationPathBox.Text = configFile.ElementAt(13);
            if (fileLen >= 15) mcpPathBox.Text = configFile.ElementAt(14);

            if (fileLen >= 16)
            {
                if (configFile.ElementAt(15) != "")
                {

                    de = StringCipher.Decrypt(configFile.ElementAt(15), passPhrase);
                    vpnPassBox.Text = de;
                }
            }

            if (fileLen >= 17)
            {
                if (Boolean.TryParse(configFile.ElementAt(16), out bool result))
                    checkBox2.Checked = result;
            }

            if (fileLen >= 18)
            {
                if (Boolean.TryParse(configFile.ElementAt(17), out bool result))
                    includeNonClientsBox.Checked = result;
            }
        }

        private void RenewIpsOrInit(bool include)
        {
            GetIpsLmHosts(include);
            if (ipList.Count() > 0)
            {
                checkedListBox1.Items.Clear();
                foreach (var item in ipList)
                {
                    checkedListBox1.Items.Add(item.Split(Convert.ToChar("\t"))[1]);
                }
                numClTextBox.Text = Convert.ToString(ipList.Count());
                firstClientIndexBox.Text = Convert.ToString(sdList.Count() + 1);
                statusLabel.Text = "Ready";
            }
            else
            {
                statusLabel.Text = "Status: Did not find LmHosts file, please check path";
            }
            statusLabel.Refresh();
            checkedListBox1.Refresh();
        }

        private void GetIpsLmHosts(bool include)
        {
            ipList = new List<string>();
            sdList = new List<string>();
            string lmhostPath = ipTextBox.Text;
            if (File.Exists(lmhostPath))
            {
                var lmHosts = File.ReadLines(lmhostPath);
                foreach (var item in lmHosts)
                {
                    if (item.IndexOf("HMIC") > 0
                        || item.IndexOf("HmiC") > 0
                        || ((item.IndexOf("HmiE01") < 0 && item.IndexOf("HmiE0") > 0)
                        || (item.IndexOf("HMIE01") < 0 && item.IndexOf("HMIE0") > 0))
                        && item.StartsWith("#") == false && item != "")
                    {
                        ipList.Add(item);
                    }
                    if ((item.IndexOf("HMID01") > 0 ||
                        item.IndexOf("HmiD01") > 0 ||
                        item.IndexOf("HMIE01") > 0 ||
                        item.IndexOf("HmiE01") > 0 ||
                        item.IndexOf("HmiS") > 0 ||
                        item.IndexOf("HMIS") > 0) &&
                        item.StartsWith("#") == false && item != "")
                    {
                        if (include)
                        {
                            ipList.Add(item);
                        }
                        sdList.Add(item);
                    }
                }
            }
        }

        private void KeepConfig()
        {
            string configPath = "NCM_Downloader.ini";
            var configFile = File.CreateText(configPath);
            configFile.WriteLine(pathTextBox.Text);
            configFile.WriteLine(firstClientIndexBox.Text);
            configFile.WriteLine(numClTextBox.Text);
            configFile.WriteLine(ipTextBox.Text);
            configFile.WriteLine(unTextBox.Text);
            var en = StringCipher.Encrypt(passTextBox.Text, passPhrase).ToString();
            configFile.WriteLine(en);
            configFile.WriteLine(rdpCheckBox.Checked);
            configFile.WriteLine(widthBox.Text);
            configFile.WriteLine(heightBox.Text);
            configFile.WriteLine(leftBox.Text);
            configFile.WriteLine(topBox.Text);
            configFile.WriteLine(parallelBox.Text);
            configFile.WriteLine(sourcePathBox.Text);
            configFile.WriteLine(destinationPathBox.Text);
            configFile.WriteLine(mcpPathBox.Text);
            en = StringCipher.Encrypt(vpnPassBox.Text, passPhrase).ToString();
            configFile.WriteLine(en);
            configFile.WriteLine(checkBox2.Checked);
            configFile.WriteLine(includeNonClientsBox.Checked);
            configFile.WriteLine("");
            configFile.WriteLine("");
            configFile.WriteLine("Authored by Muresan Radu-Adrian (MURA02)");
            configFile.Close();
        }

        private void SetTooltips()
        {
            var toolTip1 = new System.Windows.Forms.ToolTip();
            toolTip1.SetToolTip(widthBox, "Set RDP Window Width");

            var toolTip2 = new System.Windows.Forms.ToolTip();
            toolTip2.SetToolTip(heightBox, "Set RDP Window Height");

            var toolTip3 = new System.Windows.Forms.ToolTip();
            toolTip3.SetToolTip(leftBox, "Set RDP Window x");

            var toolTip4 = new System.Windows.Forms.ToolTip();
            toolTip4.SetToolTip(topBox, "Set RDP Window y");

            var toolTip5 = new System.Windows.Forms.ToolTip();
            toolTip5.SetToolTip(ipTextBox, "Set path to your lmhosts file.");

            var toolTip6 = new System.Windows.Forms.ToolTip();
            toolTip6.SetToolTip(ipTextBox, "Set path to your lmhosts file.");

            var toolTip7 = new System.Windows.Forms.ToolTip();
            toolTip7.SetToolTip(pathTextBox, "Set path to your LOAD.LOG files as specified in readme.");

            var toolTip8 = new System.Windows.Forms.ToolTip();
            toolTip8.SetToolTip(parallelBox, "Set how many parallel download processes to run");

            new System.Windows.Forms.ToolTip().SetToolTip(killButtton, "Kill Task manager on selected machines (for test)");

            new System.Windows.Forms.ToolTip().SetToolTip(mcpPathBox, "Set the client mcp file path here please");
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NCMForm));
            this.button1 = new System.Windows.Forms.Button();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.pathTextBox = new System.Windows.Forms.TextBox();
            this.ipTextBox = new System.Windows.Forms.TextBox();
            this.unTextBox = new System.Windows.Forms.TextBox();
            this.passTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.rdpCheckBox = new System.Windows.Forms.CheckBox();
            this.rdpBox1 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.topBox = new System.Windows.Forms.TextBox();
            this.leftBox = new System.Windows.Forms.TextBox();
            this.heightBox = new System.Windows.Forms.TextBox();
            this.widthBox = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.statusLabel = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.killButtton = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.firstClientIndexBox = new System.Windows.Forms.TextBox();
            this.numClTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.button10 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.mcpPathBox = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.parallelBox = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.destinationPathBox = new System.Windows.Forms.TextBox();
            this.sourcePathBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.vpnPassBox = new System.Windows.Forms.TextBox();
            this.button13 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.includeNonClientsBox = new System.Windows.Forms.CheckBox();
            this.button14 = new System.Windows.Forms.Button();
            this.rdpBox1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.AutoSize = true;
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Location = new System.Drawing.Point(250, 214);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(73, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Download";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.BackColor = System.Drawing.SystemColors.Control;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 68);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.ScrollAlwaysVisible = true;
            this.checkedListBox1.Size = new System.Drawing.Size(229, 169);
            this.checkedListBox1.TabIndex = 8;
            this.checkedListBox1.SelectedIndexChanged += new System.EventHandler(this.checkedListBox1_SelectedIndexChanged);
            // 
            // pathTextBox
            // 
            this.pathTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.pathTextBox.Location = new System.Drawing.Point(68, 11);
            this.pathTextBox.Name = "pathTextBox";
            this.pathTextBox.Size = new System.Drawing.Size(258, 18);
            this.pathTextBox.TabIndex = 11;
            this.pathTextBox.Text = "D:\\Project\\SDIB_TCM\\wincproj\\SDIB_TCM_CLT_Ref";
            // 
            // ipTextBox
            // 
            this.ipTextBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.ipTextBox.Location = new System.Drawing.Point(68, 35);
            this.ipTextBox.Name = "ipTextBox";
            this.ipTextBox.Size = new System.Drawing.Size(258, 18);
            this.ipTextBox.TabIndex = 12;
            this.ipTextBox.Text = "C:\\Users\\MURA02\\source\\repos\\WindowsFormsApp1\\NCM_Downloader\\bin\\Debug\\lmhosts";
            this.ipTextBox.TextChanged += new System.EventHandler(this.ipTextBox_TextChanged);
            // 
            // unTextBox
            // 
            this.unTextBox.Location = new System.Drawing.Point(40, 19);
            this.unTextBox.Name = "unTextBox";
            this.unTextBox.Size = new System.Drawing.Size(52, 20);
            this.unTextBox.TabIndex = 13;
            this.unTextBox.Text = "SDI";
            // 
            // passTextBox
            // 
            this.passTextBox.Location = new System.Drawing.Point(40, 45);
            this.passTextBox.Name = "passTextBox";
            this.passTextBox.PasswordChar = '*';
            this.passTextBox.Size = new System.Drawing.Size(52, 20);
            this.passTextBox.TabIndex = 14;
            this.passTextBox.Text = "A02460";
            this.passTextBox.UseSystemPasswordChar = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label2.Location = new System.Drawing.Point(10, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "LmHosts";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label3.Location = new System.Drawing.Point(6, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 13);
            this.label3.TabIndex = 16;
            this.label3.Text = "User";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label4.Location = new System.Drawing.Point(6, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 17;
            this.label4.Text = "Pass";
            // 
            // rdpCheckBox
            // 
            this.rdpCheckBox.AutoSize = true;
            this.rdpCheckBox.Checked = true;
            this.rdpCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.rdpCheckBox.Location = new System.Drawing.Point(247, 113);
            this.rdpCheckBox.Name = "rdpCheckBox";
            this.rdpCheckBox.Size = new System.Drawing.Size(76, 17);
            this.rdpCheckBox.TabIndex = 18;
            this.rdpCheckBox.Text = "AutoRDPs";
            this.rdpCheckBox.UseVisualStyleBackColor = true;
            // 
            // rdpBox1
            // 
            this.rdpBox1.Controls.Add(this.label8);
            this.rdpBox1.Controls.Add(this.label7);
            this.rdpBox1.Controls.Add(this.topBox);
            this.rdpBox1.Controls.Add(this.leftBox);
            this.rdpBox1.Controls.Add(this.heightBox);
            this.rdpBox1.Controls.Add(this.widthBox);
            this.rdpBox1.Controls.Add(this.label4);
            this.rdpBox1.Controls.Add(this.label3);
            this.rdpBox1.Controls.Add(this.unTextBox);
            this.rdpBox1.Controls.Add(this.passTextBox);
            this.rdpBox1.Location = new System.Drawing.Point(15, 356);
            this.rdpBox1.Name = "rdpBox1";
            this.rdpBox1.Size = new System.Drawing.Size(225, 73);
            this.rdpBox1.TabIndex = 19;
            this.rdpBox1.TabStop = false;
            this.rdpBox1.Text = "Remote Desktop";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label8.Location = new System.Drawing.Point(114, 23);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(28, 13);
            this.label8.TabIndex = 23;
            this.label8.Text = "Pos.";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label7.Location = new System.Drawing.Point(115, 49);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(28, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "Res.";
            // 
            // topBox
            // 
            this.topBox.Location = new System.Drawing.Point(185, 19);
            this.topBox.Name = "topBox";
            this.topBox.Size = new System.Drawing.Size(30, 20);
            this.topBox.TabIndex = 21;
            this.topBox.Text = "300";
            // 
            // leftBox
            // 
            this.leftBox.Location = new System.Drawing.Point(149, 19);
            this.leftBox.Name = "leftBox";
            this.leftBox.Size = new System.Drawing.Size(30, 20);
            this.leftBox.TabIndex = 20;
            this.leftBox.Text = "700";
            // 
            // heightBox
            // 
            this.heightBox.Location = new System.Drawing.Point(185, 45);
            this.heightBox.Name = "heightBox";
            this.heightBox.Size = new System.Drawing.Size(30, 20);
            this.heightBox.TabIndex = 19;
            this.heightBox.Text = "480";
            // 
            // widthBox
            // 
            this.widthBox.Location = new System.Drawing.Point(149, 45);
            this.widthBox.Name = "widthBox";
            this.widthBox.Size = new System.Drawing.Size(30, 20);
            this.widthBox.TabIndex = 18;
            this.widthBox.Text = "640";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkBox1.Location = new System.Drawing.Point(247, 67);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkBox1.Size = new System.Drawing.Size(70, 17);
            this.checkBox1.TabIndex = 20;
            this.checkBox1.Text = "Select All";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label6.Location = new System.Drawing.Point(10, 14);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(52, 13);
            this.label6.TabIndex = 21;
            this.label6.Text = "NCMPath";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.BackColor = System.Drawing.Color.Transparent;
            this.statusLabel.Location = new System.Drawing.Point(9, 507);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(38, 13);
            this.statusLabel.TabIndex = 23;
            this.statusLabel.Text = "Ready";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(213, 285);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(0, 13);
            this.label9.TabIndex = 24;
            // 
            // killButtton
            // 
            this.killButtton.Location = new System.Drawing.Point(295, 356);
            this.killButtton.Name = "killButtton";
            this.killButtton.Size = new System.Drawing.Size(25, 11);
            this.killButtton.TabIndex = 25;
            this.killButtton.Text = "Kill";
            this.killButtton.UseVisualStyleBackColor = true;
            this.killButtton.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(6, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(102, 20);
            this.textBox1.TabIndex = 26;
            this.textBox1.Text = "TcmHmiC05";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(114, 19);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(135, 20);
            this.textBox2.TabIndex = 27;
            this.textBox2.Text = "Taskmgr";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Location = new System.Drawing.Point(364, 363);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(312, 47);
            this.groupBox1.TabIndex = 28;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Process Control Test";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(361, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Clients start index:";
            this.label1.Visible = false;
            // 
            // firstClientIndexBox
            // 
            this.firstClientIndexBox.Enabled = false;
            this.firstClientIndexBox.Location = new System.Drawing.Point(461, 65);
            this.firstClientIndexBox.Name = "firstClientIndexBox";
            this.firstClientIndexBox.Size = new System.Drawing.Size(28, 20);
            this.firstClientIndexBox.TabIndex = 1;
            this.firstClientIndexBox.Text = "3";
            this.firstClientIndexBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.firstClientIndexBox.Visible = false;
            // 
            // numClTextBox
            // 
            this.numClTextBox.Enabled = false;
            this.numClTextBox.Location = new System.Drawing.Point(461, 94);
            this.numClTextBox.Name = "numClTextBox";
            this.numClTextBox.Size = new System.Drawing.Size(28, 20);
            this.numClTextBox.TabIndex = 9;
            this.numClTextBox.Text = "4";
            this.numClTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numClTextBox.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(361, 97);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Number of clients:";
            this.label5.Visible = false;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(449, 130);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(40, 23);
            this.button3.TabIndex = 29;
            this.button3.Text = "Cln";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            this.button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(20, 327);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(40, 23);
            this.button4.TabIndex = 30;
            this.button4.Text = "Stop";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.Button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(228, 327);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(40, 23);
            this.button5.TabIndex = 31;
            this.button5.Text = "Start";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.Button5_Click);
            // 
            // button6
            // 
            this.button6.Enabled = false;
            this.button6.Location = new System.Drawing.Point(72, 327);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(40, 23);
            this.button6.TabIndex = 32;
            this.button6.Text = "Del";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.Button6_Click);
            // 
            // button7
            // 
            this.button7.Enabled = false;
            this.button7.Location = new System.Drawing.Point(124, 327);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(40, 23);
            this.button7.TabIndex = 33;
            this.button7.Text = "Copy";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.Button7_Click);
            // 
            // button8
            // 
            this.button8.Enabled = false;
            this.button8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button8.Location = new System.Drawing.Point(232, 247);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(91, 22);
            this.button8.TabIndex = 34;
            this.button8.Text = "Alt. Download";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.Button8_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(23, 304);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(84, 13);
            this.label13.TabIndex = 43;
            this.label13.Text = "Client MCP Path";
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(280, 327);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(40, 23);
            this.button10.TabIndex = 37;
            this.button10.Text = "RDx";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(176, 327);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(40, 23);
            this.button9.TabIndex = 36;
            this.button9.Text = "RDc";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // mcpPathBox
            // 
            this.mcpPathBox.Location = new System.Drawing.Point(113, 301);
            this.mcpPathBox.Name = "mcpPathBox";
            this.mcpPathBox.Size = new System.Drawing.Size(209, 20);
            this.mcpPathBox.TabIndex = 42;
            this.mcpPathBox.Text = "D:\\Project\\SDIB_CSPM_CLT\\SDIB_CSPM_CLT.MCP";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(246, 278);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(29, 13);
            this.label12.TabIndex = 41;
            this.label12.Text = "Multi";
            // 
            // parallelBox
            // 
            this.parallelBox.Location = new System.Drawing.Point(281, 275);
            this.parallelBox.Name = "parallelBox";
            this.parallelBox.Size = new System.Drawing.Size(41, 20);
            this.parallelBox.TabIndex = 40;
            this.parallelBox.Text = "2";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(20, 277);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(60, 13);
            this.label11.TabIndex = 39;
            this.label11.Text = "Destination";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(37, 251);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(41, 13);
            this.label10.TabIndex = 38;
            this.label10.Text = "Source";
            // 
            // destinationPathBox
            // 
            this.destinationPathBox.Location = new System.Drawing.Point(84, 274);
            this.destinationPathBox.Name = "destinationPathBox";
            this.destinationPathBox.Size = new System.Drawing.Size(142, 20);
            this.destinationPathBox.TabIndex = 37;
            this.destinationPathBox.Text = "\\\\vmware-host\\Shared Folders\\C\\Users\\MURA02\\source\\repos\\WindowsFormsApp1\\WinCC T" +
    "imer\\bin\\Debug\\results";
            // 
            // sourcePathBox
            // 
            this.sourcePathBox.Location = new System.Drawing.Point(84, 248);
            this.sourcePathBox.Name = "sourcePathBox";
            this.sourcePathBox.Size = new System.Drawing.Size(142, 20);
            this.sourcePathBox.TabIndex = 36;
            this.sourcePathBox.Text = "\\\\vmware-host\\Shared Folders\\C\\Users\\MURA02\\source\\repos\\WindowsFormsApp1\\WinCC T" +
    "imer\\bin\\Debug\\results";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(364, 329);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 36;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(364, 269);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(75, 23);
            this.button11.TabIndex = 37;
            this.button11.Text = "MCP test";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.Button11_Click);
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(364, 300);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(75, 23);
            this.button12.TabIndex = 38;
            this.button12.Text = "button12";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(254, 412);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(72, 17);
            this.checkBox2.TabIndex = 39;
            this.checkBox2.Text = "RD Mode";
            this.checkBox2.UseVisualStyleBackColor = true;
            this.checkBox2.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(364, 159);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(100, 20);
            this.textBox3.TabIndex = 44;
            this.textBox3.Text = "MURA02";
            // 
            // vpnPassBox
            // 
            this.vpnPassBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.vpnPassBox.Location = new System.Drawing.Point(169, 11);
            this.vpnPassBox.Name = "vpnPassBox";
            this.vpnPassBox.Size = new System.Drawing.Size(100, 18);
            this.vpnPassBox.TabIndex = 45;
            this.vpnPassBox.Text = "Andreea~";
            this.vpnPassBox.UseSystemPasswordChar = true;
            this.vpnPassBox.Visible = false;
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(275, 8);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(51, 23);
            this.button13.TabIndex = 47;
            this.button13.Text = "VPN";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Visible = false;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(12, 435);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(314, 69);
            this.listBox1.TabIndex = 48;
            // 
            // includeNonClientsBox
            // 
            this.includeNonClientsBox.AutoSize = true;
            this.includeNonClientsBox.Location = new System.Drawing.Point(247, 90);
            this.includeNonClientsBox.Name = "includeNonClientsBox";
            this.includeNonClientsBox.Size = new System.Drawing.Size(79, 17);
            this.includeNonClientsBox.TabIndex = 49;
            this.includeNonClientsBox.Text = "Non-clients";
            this.includeNonClientsBox.UseVisualStyleBackColor = true;
            this.includeNonClientsBox.CheckedChanged += new System.EventHandler(this.checkBox3_CheckedChanged);
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(250, 507);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(75, 23);
            this.button14.TabIndex = 50;
            this.button14.Text = "Explorer";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Click += new System.EventHandler(this.button14_Click);
            // 
            // NCMForm
            // 
            this.ClientSize = new System.Drawing.Size(339, 539);
            this.Controls.Add(this.button14);
            this.Controls.Add(this.includeNonClientsBox);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button13);
            this.Controls.Add(this.vpnPassBox);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.button10);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.button9);
            this.Controls.Add(this.button11);
            this.Controls.Add(this.mcpPathBox);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.parallelBox);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.killButtton);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.destinationPathBox);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.sourcePathBox);
            this.Controls.Add(this.rdpCheckBox);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.rdpBox1);
            this.Controls.Add(this.ipTextBox);
            this.Controls.Add(this.pathTextBox);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.numClTextBox);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.firstClientIndexBox);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "NCMForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "HMI Clients Updater";
            this.rdpBox1.ResumeLayout(false);
            this.rdpBox1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void CloseConflictingProcesses(string substr)
        {
            List<Process> myProcesses = new List<Process>();
            //int i = 0;
            foreach (var proc in Process.GetProcesses())
            {
                if (proc.ProcessName.StartsWith(substr))
                {
                    myProcesses.Add(proc);
                }
            }
            foreach (var proc in myProcesses)
            {
                if (proc.Id != Process.GetCurrentProcess().Id)
                {
                    proc.Kill();
                }
            }
        }
        #endregion

        #region Old NCM Download

        private void Button1_Click(object sender, EventArgs e)
        {
            statusLabel.Text = "";
            //
            // check inputs
            //
            if (pathTextBox.Text == null || pathTextBox.Text == "")
            {
                MessageBox.Show(new Form { TopMost = true }, " Please input the path in the following form: " + @"D:\Project\SDIB_TCM\wincproj\SDIB_TCM_CLT_Ref");
                return;
            }
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                //MessageBox.Show(new Form { TopMost = true }, " You have not checked any clients to download to!");
                statusLabel.Text = "Status: You have not checked any clients to download to!";
                statusLabel.Refresh();
                return;
            }
            KeepConfig();
            DownloadProcess();
        }

        public string GetControlText(IntPtr hWnd)
        {

            // Get the size of the string required to hold the window title (including trailing null.) 
            Int32 titleSize = PInvokeLibrary.SendMessage((int)hWnd, (int)WindowsMessages.WM_GETTEXTLENGTH, 0, 0).ToInt32();

            // If titleSize is 0, there is no title so return an empty string (or null)
            if (titleSize == 0)
                return String.Empty;

            StringBuilder title = new StringBuilder(titleSize + 1);

            PInvokeLibrary.SendMessage(hWnd, (int)WindowsMessages.WM_GETTEXT, title.Capacity, title);

            return title.ToString();
        }

        #region TextExtractArea
        // privileges
        const int PROCESS_CREATE_THREAD = 0x0002;
        const int PROCESS_QUERY_INFORMATION = 0x0400;
        const int PROCESS_VM_OPERATION = 0x0008;
        const int PROCESS_VM_WRITE = 0x0020;
        const int PROCESS_VM_READ = 0x0010;

        // used for memory allocation
        const uint MEM_COMMIT = 0x00001000;
        const int MEM_DECOMMIT = 0x4000;
        const uint MEM_RESERVE = 0x00002000;
        const uint PAGE_READWRITE = 4;

        ///<summary>Returns the tree node information from another process.</summary>
        ///<param name="hwndItem">Handle to a tree node item.</param>
        ///<param name="hwndTreeView">Handle to a tree view control.</param>
        ///<param name="process">Process hosting the tree view control.</param>
        ///
        private void Testing(IntPtr ncmHandle, string anyPopupClass)
        {
            //testing for treeview
            var parent3 = PInvokeLibrary.FindWindowEx(ncmHandle, IntPtr.Zero, "MDIClient", null);
            var parent2 = PInvokeLibrary.FindWindowEx(parent3, IntPtr.Zero, "Afx:400000:b:10003:6:104d09c9", null);
            var parent1 = PInvokeLibrary.FindWindowEx(parent2, IntPtr.Zero, "AfxFrameOrView42", null);
            var parent = PInvokeLibrary.FindWindowEx(parent1, IntPtr.Zero, anyPopupClass, "");
            var treeH = PInvokeLibrary.FindWindowEx(parent, IntPtr.Zero, "SysTreeView32", "Generic1");
            var treeItemHeight = (int)PInvokeLibrary.SendMessage(treeH, (int)WindowsMessages.TVM_GETITEMW, 5, IntPtr.Zero); //works

            //PInvokeLibrary.TVITEM item = new PInvokeLibrary.TVITEM();

            var t = AllocTest(Process.GetProcessById(7396), treeH, IntPtr.Zero);

            //instead build an xml

            var s = GetTreeItemText(treeH, IntPtr.Zero);

            Console.WriteLine("uhum");
        }

        private static NodeData AllocTest(Process process, IntPtr hwndTreeView, IntPtr hwndItem)
        {
            // code based on article posted here: http://www.codingvision.net/miscellaneous/c-inject-a-dll-into-a-process-w-createremotethread

            // handle of the process with the required privileges
            IntPtr procHandle = PInvokeLibrary.OpenProcess((PInvokeLibrary.ProcessAccessFlags)(PROCESS_CREATE_THREAD | PROCESS_QUERY_INFORMATION | PROCESS_VM_OPERATION | PROCESS_VM_WRITE | PROCESS_VM_READ), false, process.Id);

            // Write TVITEM to memory
            // Invoke TVM_GETITEM
            // Read TVITEM from memory

            var item = new TVITEMEX();
            item.hItem = hwndItem;
            item.mask = (int)(WindowsMessages.TVIF_HANDLE | /*WindowsMessages.TVIF_CHILDREN |*/ WindowsMessages.TVIF_TEXT);
            item.cchTextMax = 1024;
            item.pszText = PInvokeLibrary.VirtualAllocEx(procHandle, IntPtr.Zero, (uint)item.cchTextMax, (PInvokeLibrary.AllocationType)(MEM_COMMIT | MEM_RESERVE), (PInvokeLibrary.MemoryProtection)PAGE_READWRITE); // node text pointer

            byte[] data = GetBytes(item);

            int dwSize = (int)data.Length;
            IntPtr allocMemAddress = PInvokeLibrary.VirtualAllocEx(procHandle, IntPtr.Zero, (uint)dwSize, (PInvokeLibrary.AllocationType)(MEM_COMMIT | MEM_RESERVE), (PInvokeLibrary.MemoryProtection)PAGE_READWRITE); // TVITEM pointer

            int nSize = dwSize;
            IntPtr bytesWritten;
            bool successWrite = PInvokeLibrary.WriteProcessMemory(procHandle, allocMemAddress, data, nSize, out bytesWritten);

            var sm = PInvokeLibrary.SendMessage(hwndTreeView, (int)WindowsMessages.TVM_GETITEM, 0, allocMemAddress);

            IntPtr bytesRead;
            bool successRead = PInvokeLibrary.ReadProcessMemory(procHandle, allocMemAddress, data, nSize, out bytesRead);

            IntPtr bytesReadText;
            byte[] nodeText = new byte[item.cchTextMax];
            bool successReadText = PInvokeLibrary.ReadProcessMemory(procHandle, item.pszText, nodeText, (int)item.cchTextMax, out bytesReadText);

            bool success1 = PInvokeLibrary.VirtualFreeEx(procHandle, allocMemAddress, dwSize, (PInvokeLibrary.AllocationType)MEM_DECOMMIT);
            bool success2 = PInvokeLibrary.VirtualFreeEx(procHandle, item.pszText, (int)item.cchTextMax, (PInvokeLibrary.AllocationType)MEM_DECOMMIT);

            var item2 = FromBytes<TVITEMEX>(data);

            String name = Encoding.Unicode.GetString(nodeText);
            int x = name.IndexOf('\0');
            if (x >= 0)
                name = name.Substring(0, x);

            NodeData node = new NodeData
            {
                Text = name,
                HasChildren = (item2.cChildren == 1)
            };

            return node;
        }

        public class NodeData
        {
            public String Text { get; set; }
            public bool HasChildren { get; set; }
        }

        private static byte[] GetBytes(Object item)
        {
            int size = Marshal.SizeOf(item);
            byte[] arr = new byte[size];
            IntPtr ptr = Marshal.AllocHGlobal(size);

            Marshal.StructureToPtr(item, ptr, true);
            Marshal.Copy(ptr, arr, 0, size);
            Marshal.FreeHGlobal(ptr);

            return arr;
        }

        private static T FromBytes<T>(byte[] arr)
        {
            T item = default;
            int size = Marshal.SizeOf(item);
            IntPtr ptr = Marshal.AllocHGlobal(size);
            Marshal.Copy(arr, 0, ptr, size);
            item = (T)Marshal.PtrToStructure(ptr, typeof(T));
            Marshal.FreeHGlobal(ptr);
            return item;
        }

        // Note: different layouts required depending on OS versions.
        // https://msdn.microsoft.com/en-us/library/windows/desktop/bb773459%28v=vs.85%29.aspx
        [StructLayout(LayoutKind.Sequential)]
        public struct TVITEMEX
        {
            public uint mask;
            public IntPtr hItem;
            public uint state;
            public uint stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int cChildren;
            public IntPtr lParam;
            public int iIntegral;
            public uint uStateEx;
            public IntPtr hwnd;
            public int iExpandedImage;
            public int iReserved;
        }

        public static string GetTreeItemText(IntPtr treeViewHwnd, IntPtr hItem)
        {
            string itemText;

            PInvokeLibrary.GetWindowThreadProcessId(treeViewHwnd, out uint pid);

            IntPtr process = PInvokeLibrary.OpenProcess(PInvokeLibrary.ProcessAccessFlags.VirtualMemoryOperation | PInvokeLibrary.ProcessAccessFlags.VirtualMemoryRead | PInvokeLibrary.ProcessAccessFlags.VirtualMemoryWrite | PInvokeLibrary.ProcessAccessFlags.QueryInformation, false, (int)pid);
            if (process == IntPtr.Zero)
                throw new Exception("Could not open handle to owning process of TreeView", new Win32Exception());

            try
            {
                uint tviSize = (uint)Marshal.SizeOf(typeof(PInvokeLibrary.TVITEM));

                uint textSize = (uint)WindowsMessages.MY_MAXLVITEMTEXT;
                bool isUnicode = PInvokeLibrary.IsWindowUnicode(treeViewHwnd);
                if (isUnicode)
                    textSize *= 2;

                IntPtr tviPtr = IntPtr.Zero;
                try
                {
                    tviPtr = PInvokeLibrary.VirtualAllocEx(process, IntPtr.Zero, tviSize + textSize, PInvokeLibrary.AllocationType.Commit, PInvokeLibrary.MemoryProtection.ReadWrite);
                }
                catch (Exception)
                {
                    if (tviPtr == IntPtr.Zero)
                        throw new Exception("Could not allocate memory in owning process of TreeView", new Win32Exception());
                }

                try
                {
                    IntPtr textPtr = IntPtr.Add(tviPtr, (int)tviSize);

                    PInvokeLibrary.TVITEM tvi = new PInvokeLibrary.TVITEM();
                    tvi.mask = (uint)WindowsMessages.TVIF_TEXT;
                    tvi.hItem = hItem;
                    tvi.cchTextMax = (int)WindowsMessages.MY_MAXLVITEMTEXT;
                    tvi.pszText = textPtr;

                    IntPtr ptr = Marshal.AllocHGlobal((IntPtr)tviSize);
                    try
                    {
                        IntPtr myPtr = IntPtr.Zero;
                        Marshal.StructureToPtr(tvi, ptr, false);
                        if (!PInvokeLibrary.WriteProcessMemory(process, tviPtr, ptr, (int)tviSize, out myPtr))
                            throw new Exception("Could not write to memory in owning process of TreeView", new Win32Exception());
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }

                    if ((int)PInvokeLibrary.SendMessage(treeViewHwnd, (int)WindowsMessages.TVM_GETITEM, 0, tviPtr) != 1)
                        throw new Exception("Could not get item data from TreeView");

                    ptr = Marshal.AllocHGlobal((int)textSize);
                    try
                    {
                        int bytesRead = 0;
                        IntPtr bytesReadPtr = IntPtr.Zero;
                        if (!PInvokeLibrary.ReadProcessMemory(process, textPtr, ptr, (int)textSize, out bytesReadPtr))
                            throw new Exception("Could not read from memory in owning process of TreeView", new Win32Exception());

                        bytesRead = (int)bytesReadPtr;

                        if (isUnicode)
                            itemText = Marshal.PtrToStringUni(ptr, bytesRead / 2);
                        else
                            itemText = Marshal.PtrToStringAnsi(ptr, bytesRead);
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }
                }
                finally
                {
                    PInvokeLibrary.VirtualFreeEx(process, tviPtr, 0, (PInvokeLibrary.AllocationType)PInvokeLibrary.FreeType.Release);
                }
            }
            finally
            {
                PInvokeLibrary.CloseHandle(process);
            }

            //char[] arr = itemText.ToCharArray(); //<== use this array to look at the bytes in debug mode

            return itemText;
        }

        public List<string> ExtractWindowTextByHandle(IntPtr handle)
        {
            var extractedText = new List<string>();
            List<IntPtr> childObjects = new WindowHandleInfo(handle).GetAllChildHandles();
            for (int i = 0; i < childObjects.Count; i++)
            {
                extractedText.Add(GetControlText(childObjects[i]));
            }
            return extractedText;
        }

        public List<string> ExtractTextByProcessName(string handle)
        {
            List<System.IntPtr> childObjects = new List<System.IntPtr>();
            var extractedText = new List<string>();
            Process[] anotherApps = Process.GetProcessesByName("s7tgtopx");
            if (anotherApps.Length > 0)
            {
                if (anotherApps[0] != null)
                {
                    childObjects = new WindowHandleInfo(anotherApps[0].MainWindowHandle).GetAllChildHandles();
                    for (int i = 0; i < childObjects.Count; i++)
                    {
                        extractedText.Add(GetControlText(childObjects[i]));
                    }
                }
            }
            return extractedText;
        }
        #endregion

        private void SleepUntilDownloadFeedback(int clientIndex)
        {
            bool finished = false;
            bool copied = false;
            DateTime copiedTime = DateTime.Now;
            DateTime comparTime = copiedTime;

            var filePath = pathTextBox.Text + "(" + clientIndex + ")\\winccom\\LOAD.LOG";

            System.Threading.Thread.Sleep(30000);

            while (finished == false)
            {
                if (File.Exists(filePath))
                {
                    FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    StreamReader sr = new StreamReader(fs);
                    List<String> lst = new List<string>();

                    while (!sr.EndOfStream)
                        lst.Add(sr.ReadLine());

                    var lastLine = lst.Last(); //File.ReadLines(filePath).Last();
                    finished = lastLine.Contains("The lock on the project was removed"); //now can click ok to conclude download and onto next

                    //safety wait for 40 seconds after the files have been copied, then check is finished...
                    if (lastLine.Contains("The files were copied successfully") && copiedTime == comparTime)
                    {
                        copied = true;
                        copiedTime = DateTime.Now;
                    }
                    if (lastLine.Contains("The computer name was changed in the project") && copied == true)
                    {
                        int diffInSeconds = (DateTime.Now - copiedTime).Seconds;
                        if (finished == false && diffInSeconds > 40)
                        {
                            finished = true;
                        }
                    }
                }
                System.Threading.Thread.Sleep(500);
            }
        }

        private static List<IntPtr> FindWindowByType(IntPtr hWndParent, string type) //something's off
        {
            var list = new List<IntPtr>();
            int ct = 0;
            IntPtr result = IntPtr.Zero;

            do
            {
                result = PInvokeLibrary.FindWindowEx(hWndParent, IntPtr.Zero, type, null);
                if (result != IntPtr.Zero)
                {
                    list.Add(result);
                    ++ct;
                }
            } while (result != IntPtr.Zero);

            return list;
        }

        private void DownloadProcess()
        {
            //int clientProg = progressBar1.Maximum / checkedListBox1.CheckedIndices.Count;
            // init logFile
            //
            LogToFile("Started actions ******************************************");
            //
            // initialize handles for windows
            //
            var ncmWndClass = "s7tgtopx"; //ncm manager main window
            var anyPopupClass = "#32770"; //usually any popup

            IntPtr ncmHandle = PInvokeLibrary.FindWindow(ncmWndClass, null);
            IntPtr tgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, null);

            //Testing(ncmHandle, anyPopupClass);

            //richTextBox1.Text = Convert.ToString(childNode);

            if (ncmHandle == IntPtr.Zero)
            {
                MessageBox.Show(new Form { TopMost = true }, " Simatic NCM Manager is not running.");
                LogToFile("NCM Manager was not running.");

                return;
            }

            //handle the missing software package notification
            if (tgtWndHandle != IntPtr.Zero)
            {
                IntPtr btnHandle = PInvokeLibrary.FindWindowEx(tgtWndHandle, IntPtr.Zero, "Button", null);

                if (btnHandle != IntPtr.Zero)
                {
                    PInvokeLibrary.SetForegroundWindow(tgtWndHandle);
                    PInvokeLibrary.SendMessage(btnHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                }
            }
            else
            {
                LogToFile("The missing software package notification did not appear - that is ok");

                PInvokeLibrary.SetForegroundWindow(ncmHandle);
            }

            if (tgtWndHandle != IntPtr.Zero)
            {
                //now perform actions from now on, i.e. CTRL+L
                for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
                {
                    if (i > 0)
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    // sys tree view 32 already selected when focusing - navigate from here
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    System.Threading.Thread.Sleep(200);
                    int clientIndex = checkedListBox1.CheckedIndices[i];
                    string clientName = checkedListBox1.CheckedItems[i].ToString();
                    StopWinCCRuntime(clientName);
                    //
                    //get ip here
                    //
                    var myIp = ipList.Where(x => x.Contains(clientName)).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                    //
                    //download process starts here - first needs to navigate to correct index

                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    ResetExpansions(ncmHandle);
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    ReturnToFirstClient(ncmHandle);
                    PInvokeLibrary.SetForegroundWindow(ncmHandle);
                    DownloadToCurrentIndex(clientIndex, ncmHandle);

                    LogToFile("Attempting download to client " + clientName);
                    statusLabel.Text = "Downloading to " + clientName + "...";
                    statusLabel.Refresh();


                    var started = DateTime.Now;

                    //now new window with download os
                    System.Threading.Thread.Sleep(500);
                    IntPtr osDldTgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Download OS");
                    if (osDldTgtWndHandle == IntPtr.Zero)
                    {
                        osDldTgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Downloading to target system");
                        if (osDldTgtWndHandle == IntPtr.Zero)
                        {
                            osDldTgtWndHandle = PInvokeLibrary.FindWindow(anyPopupClass, null);
                        }
                    }

                    if (osDldTgtWndHandle != IntPtr.Zero)
                    {
                        PInvokeLibrary.SetForegroundWindow(osDldTgtWndHandle);
                        IntPtr DlButtonHandle = PInvokeLibrary.FindWindowEx(osDldTgtWndHandle, IntPtr.Zero, "Button", "OK");
                        if (DlButtonHandle != IntPtr.Zero)
                        {
                            PInvokeLibrary.SendMessage(DlButtonHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                            System.Threading.Thread.Sleep(500); //important to wait a bit

                            IntPtr deactivateRTPopup = PInvokeLibrary.FindWindow(anyPopupClass, "Target system");
                            if (deactivateRTPopup != IntPtr.Zero)
                            {
                                PInvokeLibrary.SetForegroundWindow(deactivateRTPopup);
                                System.Threading.Thread.Sleep(500); //important to wait a bit
                                IntPtr YesButton = PInvokeLibrary.FindWindowEx(deactivateRTPopup, IntPtr.Zero, "Button", "&Yes");
                                if (YesButton != IntPtr.Zero)
                                {
                                    PInvokeLibrary.SendMessage(YesButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                                }
                                else
                                {
                                    LogToFile("The Ok Button was not found in the deactivation window!");
                                    //MessageBox.Show(new Form { TopMost = true }, "The Ok Button was not found!"); //careful to focus on it
                                }
                            }
                            else
                            {
                                LogToFile("Did not find target system popup! (runtime was not active on this client)");
                                //MessageBox.Show(new Form { TopMost = true }, "Did not find target system popup!"); //careful to focus on it
                            }

                            #region TaskKill process if deactivating/closing

                            var user = unTextBox.Text;
                            var pass = passTextBox.Text;
                            var ip = myIp;
                            //IntPtr killGuideHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Downloading to target system");
                            //if (killGuideHandle != IntPtr.Zero)
                            //{
                            //    //Download to target system was completed successfully. do not send enter until this text is present in the window...
                            //    bool flagKilled = false;
                            //    do
                            //    {
                            //        var killGuideText = ExtractWindowTextByHandle(killGuideHandle);
                            //        if (killGuideText.Where(x => x.Contains("Closing project on the Runtime OS")).Count() > 0 || killGuideText.Where(x => x.Contains("Deactivating project on the Runtime OS")).Count() > 0) //check if closing project takes too long...
                            //        {
                            //            LogToFile("Attempting to kill " + processName + " at " + ip + " with username " + user + " and password " + pass + " on client " + clientName);
                            //            try
                            //            {
                            //                var processName = "CCOnScreenKeyboard";
                            //                KillProcessViaPowershellOnMachine(clientName, processName);
                            //            }
                            //            catch (Exception)
                            //            {
                            //                LogToFile("was not able to kill process cconscreenkeyboard");
                            //                continue;
                            //            }
                            //            flagKilled = true;
                            //        }
                            //        System.Threading.Thread.Sleep(50);
                            //    } while (flagKilled == false);
                            //} //this was the  freakin' useless wait


                            #endregion

                            //call remote desktop open here
                            if (rdpCheckBox.Checked == true)
                            {
                                //certificate needs to be generated here
                                System.Threading.Thread.Sleep(10000); //important to wait a bit
                                OpenRemoteSession(myIp, unTextBox.Text, passTextBox.Text);
                            }

                            //if Canceled by the user in LOAD.LOG , assume that RT station not obtainable //read load.log here to find canceled by user
                            _ = PInvokeLibrary.SetForegroundWindow(osDldTgtWndHandle);
                            //var downloadTargetWindowText = ExtractWindowTextByHandle(osDldTgtWndHandle);

                            #region reading the load.log file of ncm
                            //var filePath = pathTextBox.Text + "(" + checkedListBox1.CheckedIndices[i] + 1 + ")\\winccom\\LOAD.LOG";
                            //if (File.Exists(filePath))
                            //{
                            //    FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            //    StreamReader sr = new StreamReader(fs);
                            //    List<String> lst = new List<string>();

                            //    while (!sr.EndOfStream)
                            //        lst.Add(sr.ReadLine());
                            //}

                            //SleepUntilDownloadFeedback(clientIndex + 1);
                            #endregion

                            System.Threading.Thread.Sleep(30000);

                            IntPtr dldingTgtHandle = PInvokeLibrary.FindWindow(anyPopupClass, "Downloading to target system");
                            if (dldingTgtHandle != IntPtr.Zero)
                            {
                                //Download to target system was completed successfully. do not send enter until this text is present in the window...
                                var dldingTgtText = ExtractWindowTextByHandle(dldingTgtHandle);

                                DateTime foundTime = DateTime.Now;
                                bool stillClosing = false;
                                do
                                {
                                    dldingTgtText = ExtractWindowTextByHandle(dldingTgtHandle);
                                    if (dldingTgtText.Where(x => x.Contains("Closing project on the Runtime OS.")).Count() > 0) //check if closing project takes too long...
                                    {
                                        int diffInSeconds = (DateTime.Now - foundTime).Seconds;
                                        if (stillClosing == false && diffInSeconds > 60)
                                        {
                                            stillClosing = true; //not used yet

                                            MessageBox.Show(new Form { TopMost = true }, "Please check the client " + clientName + ", it seems something is open and prevents project close.");
                                        }
                                    }

                                    if (dldingTgtText.Where(x => x.Contains("Error")).Count() > 0)
                                    {
                                        System.Threading.Thread.Sleep(500);
                                        PInvokeLibrary.SetForegroundWindow(osDldTgtWndHandle);
                                        System.Threading.Thread.Sleep(500);
                                        IntPtr YesButton = PInvokeLibrary.FindWindowEx(osDldTgtWndHandle, IntPtr.Zero, "Button", "Ok");
                                        if (YesButton != IntPtr.Zero)
                                        {
                                            PInvokeLibrary.SendMessage(YesButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);
                                        }
                                        else
                                        {
                                            LogToFile("The Ok Button was not found in the downloading window!");

                                            MessageBox.Show(new Form { TopMost = true }, "The Ok Button was not found!"); //careful to focus on it
                                        }
                                        LogToFile("Error on download to client " + clientIndex + 1);

                                        continue;
                                    }

                                    if (dldingTgtText.Where(x => x.Contains("Canceled:")).Count() > 0)
                                    {
                                        LogToFile("Client " + clientIndex + " download canceled - RT station not obtainable");

                                        MessageBox.Show(new Form { TopMost = true }, " Client " + clientIndex + " download canceled - RT station not obtainable - will continue to next client download");
                                        continue;
                                    }

                                    System.Threading.Thread.Sleep(1000);

                                    foreach (var c in dldingTgtText)
                                    {
                                        LogToFile(c);
                                    }

                                } while (dldingTgtText.Where(x => x.Contains("Download to target system was completed successfully")).Count() == 0);

                                ClickButtonUsingMessage(dldingTgtHandle, "OK", "Download to target system was completed successfully");
                            }
                        }

                        #region check here
                        StopWinCCRuntime(clientName);
                        //System.Threading.Thread.Sleep(15000); //sleep before closing remote desktop
                        StartWinCCRuntime(clientName);
                        //System.Threading.Thread.Sleep(15000); //sleep before closing remote desktop
                        #endregion
                        if (rdpCheckBox.Checked == true)
                        {
                            CloseRemoteSession(myIp);
                        }

                        //progressBar1.Value = clientProg * (i + 1);
                        //statusLabel.Text = "Progress: " + progressBar1.Value.ToString() + "%";
                        //statusLabel.Text = "Progress: " + clientName;
                        var ended = DateTime.Now;
                        var secElapsed = Math.Round((ended - started).TotalSeconds, 2);
                        LogToFile(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() +
                            " : Finished download process for machine " + checkedListBox1.CheckedItems[i].ToString() +
                            " in " + secElapsed.ToString() + " seconds");

                        //int currentElapsed = Convert.ToInt32(label9.Text.Split(Convert.ToChar("~"))[1]);
                        //label9.Text = "Remaining [s]: ~" + (currentElapsed - 100).ToString();
                        label9.Refresh();
                        //progressBar1.Refresh();
                        statusLabel.Refresh();
                    }
                    else
                    {
                        MessageBox.Show(new Form { TopMost = true }, "Could not focus on download popup!"); //careful to focus on it
                    }
                }
                LogToFile("Closing logfile");

                MessageBox.Show(new Form { TopMost = true }, "The NCM download process has been finished!"); //careful to focus on it
            }
        }

        private void ClickButtonUsingMessage(IntPtr windowHandle, string buttonText, string windowText)
        {
            bool success = true;
            do
            {
                try
                {
                    PInvokeLibrary.SetForegroundWindow(windowHandle);
                    IntPtr OkButton = PInvokeLibrary.FindWindowEx(windowHandle, IntPtr.Zero, "Button", buttonText);
                    if (OkButton != IntPtr.Zero)
                    {
                        PInvokeLibrary.SendMessage(OkButton, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, OkButton);
                        var checkText = ExtractWindowTextByHandle(windowHandle);
                        if (checkText.Where(x => x.Contains(windowText)).Count() == 0)
                        {
                            success = true;
                        }
                    }
                    else
                    {
                        LogToFile("The " + buttonText + " button was not found in the downloading window to confirm finish!");

                        //MessageBox.Show(new Form { TopMost = true }, "The OK Button was not found in the downloading to target system window!"); //careful to focus on it
                        success = false;
                    }
                }
                catch (Exception exc)
                {
                    LogToFile(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC :" + exc.Message);

                    //MessageBox.Show(new Form { TopMost = true }, exc.Message);
                    success = false;
                }
            } while (success == false);
        }

        private void SendKeyHandled(IntPtr windowHandle, string key)
        {
            bool success;
            do
            {
                try
                {
                    PInvokeLibrary.SetForegroundWindow(windowHandle);
                    SendKeys.Send(key);
                    success = true;
                }
                catch (Exception e)
                {
                    LogToFile(DateTime.UtcNow.ToLongDateString() + " " + DateTime.UtcNow.ToLongTimeString() + " UTC :" + e.Message);

                    success = false;
                    System.Threading.Thread.Sleep(10);
                }
            } while (success == false);
        }

        private void NavigateToIndex(int index, IntPtr ncm)
        {
            for (int i = 0; i < index; i++)
            {
                SendKeyHandled(ncm, "{DOWN}");
            }
        }

        private void ResetExpansions(IntPtr ncm)
        {
            int num = int.Parse(numClTextBox.Text) * 3;
            for (int i = 0; i < num; i++) //go up
            {
                SendKeyHandled(ncm, "{UP}");
            }
            for (int i = 0; i < num; i++) //expand all
            {
                SendKeyHandled(ncm, "{DOWN}");
                for (int j = 0; j < 6; j++)
                {
                    SendKeyHandled(ncm, "{RIGHT}");
                }
            }
            for (int i = 0; i < num; i++) //back to compact
            {
                for (int j = 0; j < 4; j++)
                {
                    SendKeyHandled(ncm, "{LEFT}");
                }
                SendKeyHandled(ncm, "{UP}");
            }
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{DOWN}"); //go to first dev station or whatever in the list
        }

        private void DownloadToCurrentIndex(int index, IntPtr ncm)
        {
            NavigateToIndex(index, ncm);
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "{RIGHT}");
            SendKeyHandled(ncm, "^(l)");
            System.Threading.Thread.Sleep(500);
        }

        private void ReturnToFirstClient(IntPtr ncm)
        {
            for (int i = 0; i < 10; i++)
            {
                SendKeyHandled(ncm, "{LEFT}");
            }
            SendKeyHandled(ncm, "{RIGHT}");

            for (int i = 0; i < int.Parse(firstClientIndexBox.Text); i++)
            {
                SendKeyHandled(ncm, "{DOWN}");
            }
        }
        #endregion

        #region Changed Events
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Checked);
                }
            }
            else
            {
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
                }
            }
            checkedListBox1_SelectedIndexChanged(checkBox1, new System.EventArgs());
        }

        private void ipTextBox_TextChanged(object sender, EventArgs e)
        {
            RenewIpsOrInit(includeNonClientsBox.Checked);
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //label9.Text = "Remaining [s]: ~" + (checkedListBox1.CheckedIndices.Count * 100).ToString();
            label9.Refresh();
        }
        #endregion

        private static void LogToFile(string content)
        {
            using (var fileWriter = new StreamWriter(Path.Combine(Application.StartupPath, "NCM_Downloader.logger"), true))
            {
                DateTime date = DateTime.UtcNow;
                fileWriter.WriteLine(date.Year + "/" + date.Month + "/" + date.Day + " " + date.Hour + ":" + date.Minute + ":" + date.Second + ":" + date.Millisecond + " UTC: " + content);
                fileWriter.Close();
            }
        }

        private void KillProcessViaPowershellOnMachine(string machine, string process)
        {
            try
            {
                AddToTrustedHosts();
            }
            catch (Exception e)
            {
                LogToFile(e.Message);
            }
            finally
            {
                //Invoke-command -computername "TcmHmiC05" {Get-Process | ? {$_.name -match 'CCOnScreenKeyboard'} | Stop-Process -Force}
                //kill these:
                try
                {
                    Runspace runSpace = RunspaceFactory.CreateRunspace();
                    runSpace.Open();
                    Pipeline pipeline = runSpace.CreatePipeline();

                    Command invokeScript = new Command("Invoke-Command");
                    RunspaceInvoke invoke = new RunspaceInvoke();
                    //Invoke-Command -scriptBlock
                    ScriptBlock sb = invoke.Invoke("{Get-Process | ? {$_.name -match '" + process + "'} | Stop-Process -Force}")[0].BaseObject as ScriptBlock;
                    invokeScript.Parameters.Add("computername", machine);
                    invokeScript.Parameters.Add("scriptBlock", sb);

                    pipeline.Commands.Add(invokeScript);
                    Collection<PSObject> output = pipeline.Invoke();
                    foreach (PSObject obj in output)
                    {
                        LogToFile(obj.ToString());
                    }
                }
                catch (Exception e)
                {
                    LogToFile(e.Message + " at kill process");
                }
                finally
                {
                    LogToFile("could not kill process");
                }
            }
        }

        private void AddToTrustedHosts()
        {
            //run ps command here
            //or run ps file
            System.Diagnostics.Process.Start(Path.Combine(Application.StartupPath, "StartTrustedHosts.ps1"));
        }

        private void StartDownloads(int paralellismDeg)
        {
            if (mcpPathBox.Text == "")
            {
                MessageBox.Show(new Form { TopMost = true }, "MCP Path textbox is null");
                return;
            }

            //parallel method
            List<string> checkedItems = new List<string>();
            foreach (object c in checkedListBox1.CheckedItems)
            {
                checkedItems.Add(c.ToString());
            }

            #region first stop rt and reset wincc, delete paralelly
            var msg = "";
            Parallel.ForEach(checkedItems,
                    new ParallelOptions { MaxDegreeOfParallelism = paralellismDeg },
                    (CheckedItem) =>
                    {
                        //do something
                        msg = "Started download process for " + CheckedItem;
                        statusLabel.Text = msg;
                        LogToFile(msg);

                        var ip = ipList.Where(x => x.Contains(CheckedItem.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];

                        var clientPath = @"\\" + CheckedItem + "\\" + destinationPathBox.Text;
                        var sourcePath = sourcePathBox.Text;
                        if (sourcePath.EndsWith(@"\"))
                            sourcePath = sourcePath.Substring(0, sourcePath.Length - 1);
                        if (clientPath.EndsWith(@"\"))
                            clientPath = clientPath.Substring(0, clientPath.Length - 1);

                        StopWinCCRuntime(CheckedItem);
                        //DeleteProjectFolder(CheckedItem);
                        Copy(sourcePathBox.Text, @"\\" + ip + @"\" + destinationPathBox.Text, CheckedItem);
                        OpenRemoteSession(ip, unTextBox.Text, passTextBox.Text);
                        StartWinCCRuntime(CheckedItem);
                        CloseRemoteSession(ip);

                        msg = "Finished download process for " + CheckedItem;
                        statusLabel.Text = msg;
                        LogToFile(msg);

                        Console.WriteLine(CheckedItem);
                    });
            #endregion

        }

        #region Remote Commands
        private void OpenRemoteSession(string ip, string un, string pw)
        {
            var msg = "Opening remote session of " + ip + "...";
            statusLabel.Text = msg;
            LogToFile(msg);

            Process rdcProcess = new Process();
            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe");
            rdcProcess.StartInfo.Arguments = "/generic:TERMSRV/" + ip + " /user:" + un + " /pass:" + pw;
            rdcProcess.Start();
            rdcProcess.StartInfo.FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\mstsc.exe");
            rdcProcess.StartInfo.Arguments = "/v " + ip; // ip or name of computer to connect
            rdcProcess.Start();

            IntPtr myRdp;
            do
            {
                myRdp = PInvokeLibrary.FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");
                if (myRdp != IntPtr.Zero)
                {
                    // Set the window's position.
                    int width = int.Parse(widthBox.Text);
                    int height = int.Parse(heightBox.Text);
                    int x = int.Parse(leftBox.Text);
                    int y = int.Parse(topBox.Text);

                    // Prepare the WINDOWPLACEMENT structure.
                    var placement = new PInvokeLibrary.WINDOWPLACEMENT();
                    placement.Length = Marshal.SizeOf(placement);

                    // Get the window's current placement.
                    PInvokeLibrary.GetWindowPlacement(myRdp, ref placement);
                    if (placement.ShowCmd != PInvokeLibrary.ShowWindowCommands.Normal)
                    {
                        //alter the placement
                        placement.ShowCmd = PInvokeLibrary.ShowWindowCommands.Normal;
                        //set the changes
                        PInvokeLibrary.SetWindowPlacement(myRdp, ref placement);
                    }

                    PInvokeLibrary.SetWindowPos(myRdp, IntPtr.Zero, x, y, width, height, 0);

                }
                else
                {
                    System.Threading.Thread.Sleep(1000);
                }
            } while (myRdp == IntPtr.Zero);

            msg = "Opened remote session of " + ip;
            statusLabel.Text = msg;
            LogToFile(msg);
        }

        private void CloseRemoteSession(string ip)
        {
            var msg = "Closing remote session of " + ip + "...";
            statusLabel.Text = msg;
            LogToFile(msg);
            IntPtr myRdp = PInvokeLibrary.FindWindow("TscShellContainerClass", ip + " - Remote Desktop Connection");
            if (myRdp != IntPtr.Zero)
            {
                PInvokeLibrary.SendMessage(myRdp, (int)WindowsMessages.WM_CLOSE, (int)IntPtr.Zero, IntPtr.Zero);
                System.Threading.Thread.Sleep(1000);
            }

            msg = "Closed remote session of " + ip;
            statusLabel.Text = msg;
            LogToFile(msg);
        }

        private void ResetWinCCProcesses(string machine)
        {
            var msg = "Started cleaning processes on " + machine + "...";
            statusLabel.Text = msg;
            LogToFile(msg);

            foreach (var process in processes)
            {
                var action = "Stop-Process -Force";
                if (rebootProcesses.Contains(process))
                {
                    action = "Restart-Service";
                }
                try
                {

                    Runspace runSpace = RunspaceFactory.CreateRunspace();
                    runSpace.Open();
                    Pipeline pipeline = runSpace.CreatePipeline();

                    Command invokeScript = new Command("Invoke-Command");
                    RunspaceInvoke invoke = new RunspaceInvoke();
                    //Invoke-Command -scriptBlock
                    ScriptBlock sb = invoke.Invoke("{Get-Process | ? {$_.name -match '" + process + "'} | " + action + "}")[0].BaseObject as ScriptBlock;
                    invokeScript.Parameters.Add("computername", machine);
                    invokeScript.Parameters.Add("scriptBlock", sb);

                    pipeline.Commands.Add(invokeScript);
                    Collection<PSObject> output = pipeline.Invoke();
                    foreach (PSObject obj in output)
                    {
                        LogToFile(obj.ToString());
                    }
                    LogToFile(action + " " + process);
                }
                catch (Exception exc)
                {
                    LogToFile(exc.Message + " at " + action + " " + process);
                }
                finally
                {
                    LogToFile("could not " + action + " " + process);
                }
            }

            msg = "Prepared processes on " + machine;
            statusLabel.Text = msg;
            LogToFile(msg);
        }

        private void LaunchRemoteProcessWithSettingsFile(string machine, string exeFileName, string settingsFileName)
        {
            var exePath = Application.StartupPath + exeFileName;

            UserImpersonation impersonator = new UserImpersonation();
            impersonator.impersonateUser(unTextBox.Text, "", passTextBox.Text); //No Domain is required

            try
            {
                File.Copy(exePath, @"\\" + machine + @"\C$\Temp" + exeFileName);
                //if (settingsFileName != "")
                //    File.Copy(exePath, @"\\" + machine + @"\C$\Temp" + settingsFileName);
            }
            catch (Exception exc)
            {
                LogToFile(exc.Message);
            }

            Runspace runSpace = RunspaceFactory.CreateRunspace();
            runSpace.Open();
            Pipeline pipeline = runSpace.CreatePipeline();

            Command invokeScript = new Command("Invoke-Command");
            RunspaceInvoke invoke = new RunspaceInvoke();

            var s = new SecureString();
            foreach (var ch in passTextBox.Text)
            {
                s.AppendChar(ch);
            }
            var cred = new PSCredential(unTextBox.Text, s);

            //Invoke-Command -scriptBlock
            ScriptBlock sb = invoke.Invoke(@"{Invoke-Expression -Command:""cmd.exe /c 'C:\Temp" + exeFileName + @"'""}")[0].BaseObject as ScriptBlock;
            invokeScript.Parameters.Add("ComputerName", machine);
            invokeScript.Parameters.Add("Credential", cred);
            invokeScript.Parameters.Add("ScriptBlock", sb);

            pipeline.Commands.Add(invokeScript);
            Collection<PSObject> output = pipeline.Invoke();
            foreach (PSObject obj in output)
            {
                LogToFile(obj.ToString());
            }

            File.Delete(@"\\" + machine + @"\C$\Temp" + exeFileName);
            //if (settingsFileName != "")
            //    File.Delete(@"\\" + machine + @"\C$\Temp" + settingsFileName);
            impersonator.undoimpersonateUser();
        }

        private void StartWinCCRuntime(string machine)
        {
            if (mcpPathBox.Text == "")
            {
                MessageBox.Show(new Form { TopMost = true }, "MCP Path textbox is null");
                return;
            }
            var msg = "Starting runtime activation on " + machine + "...";
            statusLabel.Text = msg;
            LogToFile(msg);
            var pdlrt = @"C:\Program Files (x86)\Siemens\WinCC\bin\PdlRt.exe";
            var winccex = @"C:\Program Files (x86)\Siemens\WinCC\bin\WinCCExplorer.exe";


            var exe = "\\StartWinCCRuntime.exe";
            var set = ""; // "\\StartWinCCRuntimeSettings.txt"; remove this issue and automatically find the mcp in the start exe
            var settPath = Application.StartupPath + set;
            //using (var fileWriter = new StreamWriter(settPath, true))
            //{
            //    fileWriter.WriteLine(mcpPathBox.Text);
            //    fileWriter.WriteLine(winccex);
            //    fileWriter.WriteLine(pdlrt);
            //    fileWriter.Close();
            //}
            LaunchRemoteProcessWithSettingsFile(machine, exe, set);

            //KillProcessViaPowershellOnMachine(machine, "StartWinCCRuntime");
            // i can has feedback?
            msg = "Started runtime on " + machine;
            statusLabel.Text = msg;
            LogToFile(msg);
        }

        private void StopWinCCRuntime(string machine)
        {
            var msg = "Started deactivating runtime on " + machine + "...";
            statusLabel.Text = msg;
            LogToFile(msg);
            KeepConfig();

            var exe = "\\StopWinCCRuntime.exe";
            var set = "";
            LaunchRemoteProcessWithSettingsFile(machine, exe, set);

            msg = "Stopped runtime on " + machine;
            statusLabel.Text = msg;
            LogToFile(msg);
        }

        public static void Copy(string sourceDirectory, string targetDirectory, string machine)
        {
            DirectoryInfo source = new DirectoryInfo(sourceDirectory);
            DirectoryInfo target = new DirectoryInfo(targetDirectory);

            CopyAll(source, target, machine);
            //what follows is because the files in root folder were not being copied....
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                if (!fi.Extension.ToLower().Contains("mcp"))
                {
                    fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                }

                if (fi.Extension.ToLower().Contains("mcp"))
                {
                    var newfi = new FileInfo(Path.Combine(target.FullName, fi.Name));
                    var machineOld = Environment.MachineName;
                    var machineNew = machine;
                    //ReplaceTextInFile(machineOld, machineNew, newfi.FullName);
                }
            }
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target, string machine)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);

                if (!fi.Extension.ToLower().Contains("mcp"))
                {
                    fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                }

                if (fi.Extension.ToLower().Contains("mcp"))
                {
                    var newfi = new FileInfo(Path.Combine(target.FullName, fi.Name));
                    var machineOld = Environment.MachineName;
                    var machineNew = machine;
                    //ReplaceTextInFile(machineOld, machineNew, newfi.FullName);
                }
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                var name = diSourceSubDir.Name;
                if (name == Environment.MachineName)
                    name = machine;
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(name);
                CopyAll(diSourceSubDir, nextTargetSubDir, machine);
            }
        }

        private void DeleteProjectFolder(string machine)
        {
            var msg = "Started deleting on " + machine + "...";
            statusLabel.Text = msg;
            LogToFile(msg);
            KeepConfig();

            UserImpersonation impersonator = new UserImpersonation();
            impersonator.impersonateUser(unTextBox.Text, "", passTextBox.Text); //No Domain is required

            Directory.Delete(@"\\" + machine + "\\" + destinationPathBox.Text, true);

            msg = "Deleted project on " + machine;
            statusLabel.Text = msg;
            LogToFile(msg);

            impersonator.undoimpersonateUser();
        }
        #endregion

        #region Button Actions
        // Send a key to the button when the user double-clicks anywhere on the form.
        private void Form1_DoubleClick(object sender, EventArgs e)
        {
            // Send the enter key to the button, which raises the click
            // event for the button. This works because the tab stop of
            // the button is 0.
            //SendKeys.Send("{ENTER}");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var id = textBox1.Text;
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            foreach (var c in list)
                KillProcessViaPowershellOnMachine(c, textBox2.Text);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            foreach (var c in list)
                StopWinCCRuntime(c);

            //if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
            //    return;
            //}
            //Parallel.ForEach(list,
            //        new ParallelOptions { MaxDegreeOfParallelism = maxPar },
            //        (c) =>
            //        {
            //            //do something
            //            ResetWinCCProcesses(c);
            //        });
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            foreach (var c in list)
                StopWinCCRuntime(c);

            //if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            //{
            //    MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
            //    return;
            //}
            //Parallel.ForEach(list,
            //        new ParallelOptions { MaxDegreeOfParallelism = maxPar },
            //        (c) =>
            //        {
            //            //do something
            //            StopWinCCRuntime(c);
            //        });
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                //return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            //foreach (var c in list)
            //    StartWinCCRuntime(c);

            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }
            Parallel.ForEach(list,
                    new ParallelOptions { MaxDegreeOfParallelism = maxPar },
                    (c) =>
                    {
                        //do something
                        StartWinCCRuntime(c);
                    });
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            //foreach (var c in list)
            //{
            //    DeleteProjectFolderViaExe(c);
            //    //foreach (var subf in selectiveFolders)
            //    //DeleteOldProjectFolder(c, subf);
            //}
            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }
            Parallel.ForEach(list,
                    new ParallelOptions { MaxDegreeOfParallelism = maxPar },
                    (c) =>
                    {
                        //do something
                        //foreach (var subf in selectiveFolders)
                        //    DeleteOldProjectFolder(c, subf);
                        DeleteProjectFolder(c);
                    });
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            //foreach (var c in list)
            //{
            //    foreach (var subf in selectiveFolders) {
            //var machine = c;
            //var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
            //          CopyFiles(ip);
            //}}

            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }
            Parallel.ForEach(list,
                    new ParallelOptions { MaxDegreeOfParallelism = maxPar },
                    (c) =>
                    {
                        //do something
                        var machine = c;
                        var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                        Copy(sourcePathBox.Text, @"\\" + ip + @"\" + destinationPathBox.Text, c);
                    });
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var exePath = Path.Combine(Application.StartupPath, "StopWinCCRuntime.exe");
            if (!File.Exists(exePath))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please have StopWinCCRuntime.exe in this folder");
                return;
            }
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items to download to have been selected");
                return;
            }
            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }

            KeepConfig();

            StartDownloads(maxPar);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            //foreach (var c in list)
            //{
            //    var machine = c;
            //    var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
            //    OpenRemoteSession(ip, unTextBox.Text, passTextBox.Text);
            //}

            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }
            Parallel.ForEach(list,
                    new ParallelOptions { MaxDegreeOfParallelism = maxPar },
                    (c) =>
                    {
                        var machine = c;
                        var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                        OpenRemoteSession(ip, unTextBox.Text, passTextBox.Text);
                    });
        }

        private void button10_Click(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            //foreach (var c in list)
            //{
            //    var machine = c;
            //    var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
            //    CloseRemoteSession(ip);
            //}

            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }
            Parallel.ForEach(list,
                    new ParallelOptions { MaxDegreeOfParallelism = maxPar },
                    (c) =>
                    {
                        var machine = c;
                        var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                        CloseRemoteSession(ip);
                    });
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            KeepConfig();
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show(new Form { TopMost = true }, "No items selected");
                return;
            }
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            var source = sourcePathBox.Text;
            var target = destinationPathBox.Text;
            if (!Int32.TryParse(parallelBox.Text, out int maxPar))
            {
                MessageBox.Show(new Form { TopMost = true }, "Please write how many parallel downloads to run in the Multi textbox");
                return;
            }
            Parallel.ForEach(list,
                    new ParallelOptions { MaxDegreeOfParallelism = maxPar },
                    (c) =>
                    {
                        var machine = c;
                        var ip = ipList.Where(x => x.Contains(machine.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];
                        Copy(source, target, machine);
                    });
        }
        #endregion

        private void Button11_Click(object sender, EventArgs e)
        {
            var machine1 = "T\0C\0M\0H\0M\0I\0D\00\01";
            var machine2 = "T\0C\0M\0H\0M\0I\0C\00\01";
            var f = @"\\vmware-host\Shared Folders\C\Users\MURA02\source\repos\SDIB_TCM_CLT.mcp";

            var file = new FileInfo(f);
            var extension = file.Extension.ToLower().Contains("mcp");

            ReplaceTextInFile(machine1, machine2, f);

            Console.WriteLine("Done");
        }

        private static void ReplaceTextInFile(string currentMachineName, string clientName, string f)
        {
            String f1;
            Encoding encoding;
            using (var reader = new StreamReader(f))
            {
                f1 = reader.ReadToEnd();
                encoding = reader.CurrentEncoding;
            }

            var flag = false;
            if (f1.Contains(currentMachineName))
            {
                f1 = f1.Replace(currentMachineName, clientName);
                flag = true;
            }

            if (f1.Contains(".\0m\0d\0f\0:"))
            {
                var index = f1.IndexOf(".\0m\0d\0f\0:");
                //LogToFile(f1.Length - index);
                var sub = f1.Substring(index, f1.Length - index);
                LogToFile("sub was " + sub);
                var endOfDBName = sub.IndexOf("\u0003\0\0\0\u0004\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\0\u0012\0" + clientName);

                LogToFile("endOfDBName was " + endOfDBName);
                var wholeFileDBNameToReplace = sub.Substring(0, endOfDBName);

                var codedString = "r\0C\0Y\0N\0o\0*\0z\0*\0t\0E\0-\0D\04\0*\0I\0r\03\0f\0-\0y\0w\0k\0w\0F\07\0q\0Z\0c\03\0*\0L\01\0R\0S\0K\0g\0r\0v\0U\0Q\0G\01\0Y\0r\0a\0B\0O\0O\0o\0D\0X\0Y\03\0v\0x\0Q\0[\0[\0\0\0\0\0\0\0\0\0\0\0\0\0";
                var codedStringIndex = wholeFileDBNameToReplace.IndexOf(codedString);
                var cleanedDBName = wholeFileDBNameToReplace.Replace(codedString, "").Replace("\0", "");
                LogToFile("Found dbName " + cleanedDBName + " for mcp of " + clientName);

                //var machine = "E7H079961";
                var machine = clientName;
                //string dbName = ".mdf:" + GetDBName(machine); //".mdf:" + CC_ELVAL_HF_20_10_20_08_41_57
                var dbName = cleanedDBName;
                LogToFile("Replacing " + cleanedDBName + " with " + dbName);

                var rebuiltDBName = "";
                for (int i = 0; i < dbName.Length; i++)
                {
                    char c = dbName[i];
                    rebuiltDBName += c + "\0";
                }
                rebuiltDBName = rebuiltDBName.Insert(codedStringIndex, codedString);
                _ = wholeFileDBNameToReplace.Equals(rebuiltDBName);

                LogToFile("Replacing");
                LogToFile(wholeFileDBNameToReplace); LogToFile(rebuiltDBName);
                f1 = f1.Replace(wholeFileDBNameToReplace, rebuiltDBName);
                flag = true;
            }
            if (flag == true)
            {
                File.WriteAllText(f, f1, encoding);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            #region sql
            //sql connection to fill data 
            var machine = "E7H079961";
            GetDBName(machine);
            #endregion
        }

        public static string GetDBName(string machine)
        {
            string connectionString = "Data Source=" + machine + "\\WINCC;Initial Catalog=SMS_RTDesign;Integrated Security=SSPI";
            SqlConnection cnn = new SqlConnection(connectionString);
            List<string> result = new List<string>();
            cnn.Open();
            SqlCommand cmd = new SqlCommand("select * from sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb'); ", cnn);
            System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add(reader["name"].ToString());

            //MessageBox.Show(result[0]);
            cnn.Close();
            return result[0];
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                button1.Visible = false;
                //checkedListBox1.Visible = false;
                pathTextBox.Visible = false;
                //ipTextBox.Visible = false;
                //unTextBox.Visible = false;
                //passTextBox.Visible = false;
                //label2.Visible = false;
                //label3.Visible = false;
                //label4.Visible = false;
                rdpCheckBox.Visible = false;
                //rdpBox1.Visible = false;
                //label8.Visible = false;
                //label7.Visible = false;
                //topBox.Visible = false;
                //leftBox.Visible = false;
                //heightBox.Visible = false;
                //widthBox.Visible = false;
                checkBox1.Visible = false;
                label6.Visible = false;
                //statusLabel.Visible = false;
                label9.Visible = false;
                killButtton.Visible = false;
                textBox1.Visible = false;
                textBox2.Visible = false;
                groupBox1.Visible = false;
                label1.Visible = false;
                firstClientIndexBox.Visible = false;
                numClTextBox.Visible = false;
                label5.Visible = false;
                button3.Visible = false;
                button4.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button7.Visible = false;
                button8.Visible = false;
                //groupBox2.Visible = false;
                label13.Visible = false;
                //button10.Visible = false;
                //button9.Visible = false;
                mcpPathBox.Visible = false;
                label12.Visible = false;
                parallelBox.Visible = false;
                label11.Visible = false;
                label10.Visible = false;
                destinationPathBox.Visible = false;
                sourcePathBox.Visible = false;
                button2.Visible = false;
                button11.Visible = false;
                button12.Visible = false;
                vpnPassBox.Visible = true;
                button13.Visible = true;
                //checkBox2.Visible = false;
                checkedListBox1.Size = new System.Drawing.Size(314, 285);
                button9.Location = new System.Drawing.Point(246, 375);
                button10.Location = new System.Drawing.Point(286, 375);
                includeNonClientsBox.Location = new System.Drawing.Point(165, 510);
            }
            else
            {
                button1.Visible = true;
                //checkedListBox1.Visible = true;
                pathTextBox.Visible = true;
                //ipTextBox.Visible = true;
                unTextBox.Visible = true;
                passTextBox.Visible = true;
                //label2.Visible = true;
                label3.Visible = true;
                label4.Visible = true;
                rdpCheckBox.Visible = true;
                rdpBox1.Visible = true;
                label8.Visible = true;
                label7.Visible = true;
                topBox.Visible = true;
                leftBox.Visible = true;
                heightBox.Visible = true;
                widthBox.Visible = true;
                checkBox1.Visible = true;
                label6.Visible = true;
                statusLabel.Visible = true;
                label9.Visible = true;
                killButtton.Visible = true;
                textBox1.Visible = true;
                textBox2.Visible = true;
                groupBox1.Visible = true;
                label1.Visible = true;
                firstClientIndexBox.Visible = true;
                numClTextBox.Visible = true;
                label5.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
                button8.Visible = true;
                //groupBox2.Visible = true;
                label13.Visible = true;
                //button10.Visible = true;
                //button9.Visible = true;
                mcpPathBox.Visible = true;
                label12.Visible = true;
                parallelBox.Visible = true;
                label11.Visible = true;
                label10.Visible = true;
                destinationPathBox.Visible = true;
                sourcePathBox.Visible = true;
                button2.Visible = true;
                button11.Visible = true;
                button12.Visible = true;
                vpnPassBox.Visible = false;
                button13.Visible = false;
                //checkBox2.Visible = true;
                checkedListBox1.Size = new System.Drawing.Size(188, 124);
                button9.Location = new System.Drawing.Point(176, 327);
                button10.Location = new System.Drawing.Point(280, 327);
                includeNonClientsBox.Location = new System.Drawing.Point(247, 90);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            KeepConfig();
            var cisco = System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe");

            var anyPopupClass = "#32770"; //usually any popup

            IntPtr ciscoWindow;
            do
            {
                ciscoWindow = PInvokeLibrary.FindWindow(anyPopupClass, "Cisco AnyConnect Secure Mobility Client");
                System.Threading.Thread.Sleep(10);
            } while (ciscoWindow == IntPtr.Zero);

            PInvokeLibrary.SetForegroundWindow(ciscoWindow);

            var data = ExtractWindowTextByHandle(ciscoWindow);
            foreach (var c in data)
            {
                listBox1.Items.Add(c);
            }
            IntPtr parentOfButtons = PInvokeLibrary.FindWindowEx(ciscoWindow, IntPtr.Zero, anyPopupClass, null);
            if (data.Contains("Connect"))
            {
                IntPtr DlButtonHandle = PInvokeLibrary.FindWindowEx(parentOfButtons, IntPtr.Zero, "Button", "Connect");
                if (DlButtonHandle == IntPtr.Zero) return;
                PInvokeLibrary.SendMessage(DlButtonHandle, (int)WindowsMessages.BM_CLICK, (int)IntPtr.Zero, IntPtr.Zero);


                IntPtr logonWindow;
                do
                {
                    logonWindow = WndSearcher.SearchForWindow(anyPopupClass, "Cisco AnyConnect | ");
                    System.Threading.Thread.Sleep(10);
                } while (logonWindow == IntPtr.Zero);

                foreach (var c in vpnPassBox.Text)
                {
                    if (c.ToString() != "~")
                    {
                        SendKeyHandled(logonWindow, c.ToString());
                    }
                    else
                    {
                        SendKeyHandled(logonWindow, "{~}");
                    }
                }
                SendKeyHandled(logonWindow, "{ENTER}");

                IntPtr confirmation;
                do
                {
                    confirmation = PInvokeLibrary.FindWindow(anyPopupClass, "Cisco AnyConnect");
                    System.Threading.Thread.Sleep(10);
                } while (confirmation == IntPtr.Zero);

                SendKeyHandled(confirmation, "{ENTER}");
            }
            if (data.Contains("Disconnect"))
            {
                IntPtr DlButtonHandle = PInvokeLibrary.FindWindowEx(parentOfButtons, IntPtr.Zero, "Button", "Disconnect");
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            RenewIpsOrInit(includeNonClientsBox.Checked);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            KeepConfig();
            var list = new List<string>();
            foreach (var c in checkedListBox1.CheckedItems)
                list.Add(c.ToString());

            foreach (var c in list)
            {
                var ip = ipList.Where(x => x.Contains(c.ToString())).FirstOrDefault().Split(Convert.ToChar("\t"))[0];

                var pass = new SecureString();
                foreach (var chr in passTextBox.Text)
                {
                    pass.AppendChar(chr);
                }

                ExtremeMirror.PinvokeWindowsNetworking.connectToRemote(@"\\" + ip + @"\c$\", unTextBox.Text, passTextBox.Text);
                Process.Start(@"\\" + ip + @"\c$\");
            }
        }
    }
}

namespace ExtremeMirror
{
    public class PinvokeWindowsNetworking
    {
        #region Consts
        const int RESOURCE_CONNECTED = 0x00000001;
        const int RESOURCE_GLOBALNET = 0x00000002;
        const int RESOURCE_REMEMBERED = 0x00000003;

        const int RESOURCETYPE_ANY = 0x00000000;
        const int RESOURCETYPE_DISK = 0x00000001;
        const int RESOURCETYPE_PRINT = 0x00000002;

        const int RESOURCEDISPLAYTYPE_GENERIC = 0x00000000;
        const int RESOURCEDISPLAYTYPE_DOMAIN = 0x00000001;
        const int RESOURCEDISPLAYTYPE_SERVER = 0x00000002;
        const int RESOURCEDISPLAYTYPE_SHARE = 0x00000003;
        const int RESOURCEDISPLAYTYPE_FILE = 0x00000004;
        const int RESOURCEDISPLAYTYPE_GROUP = 0x00000005;

        const int RESOURCEUSAGE_CONNECTABLE = 0x00000001;
        const int RESOURCEUSAGE_CONTAINER = 0x00000002;


        const int CONNECT_INTERACTIVE = 0x00000008;
        const int CONNECT_PROMPT = 0x00000010;
        const int CONNECT_REDIRECT = 0x00000080;
        const int CONNECT_UPDATE_PROFILE = 0x00000001;
        const int CONNECT_COMMANDLINE = 0x00000800;
        const int CONNECT_CMD_SAVECRED = 0x00001000;

        const int CONNECT_LOCALDRIVE = 0x00000100;
        #endregion

        #region Errors
        const int NO_ERROR = 0;

        const int ERROR_ACCESS_DENIED = 5;
        const int ERROR_ALREADY_ASSIGNED = 85;
        const int ERROR_BAD_DEVICE = 1200;
        const int ERROR_BAD_NET_NAME = 67;
        const int ERROR_BAD_PROVIDER = 1204;
        const int ERROR_CANCELLED = 1223;
        const int ERROR_EXTENDED_ERROR = 1208;
        const int ERROR_INVALID_ADDRESS = 487;
        const int ERROR_INVALID_PARAMETER = 87;
        const int ERROR_INVALID_PASSWORD = 1216;
        const int ERROR_MORE_DATA = 234;
        const int ERROR_NO_MORE_ITEMS = 259;
        const int ERROR_NO_NET_OR_BAD_PATH = 1203;
        const int ERROR_NO_NETWORK = 1222;

        const int ERROR_BAD_PROFILE = 1206;
        const int ERROR_CANNOT_OPEN_PROFILE = 1205;
        const int ERROR_DEVICE_IN_USE = 2404;
        const int ERROR_NOT_CONNECTED = 2250;
        const int ERROR_OPEN_FILES = 2401;

        private struct ErrorClass
        {
            public int num;
            public string message;
            public ErrorClass(int num, string message)
            {
                this.num = num;
                this.message = message;
            }
        }


        // Created with excel formula:
        // ="new ErrorClass("&A1&", """&PROPER(SUBSTITUTE(MID(A1,7,LEN(A1)-6), "_", " "))&"""), "
        private static ErrorClass[] ERROR_LIST = new ErrorClass[] {
            new ErrorClass(ERROR_ACCESS_DENIED, "Error: Access Denied"),
            new ErrorClass(ERROR_ALREADY_ASSIGNED, "Error: Already Assigned"),
            new ErrorClass(ERROR_BAD_DEVICE, "Error: Bad Device"),
            new ErrorClass(ERROR_BAD_NET_NAME, "Error: Bad Net Name"),
            new ErrorClass(ERROR_BAD_PROVIDER, "Error: Bad Provider"),
            new ErrorClass(ERROR_CANCELLED, "Error: Cancelled"),
            new ErrorClass(ERROR_EXTENDED_ERROR, "Error: Extended Error"),
            new ErrorClass(ERROR_INVALID_ADDRESS, "Error: Invalid Address"),
            new ErrorClass(ERROR_INVALID_PARAMETER, "Error: Invalid Parameter"),
            new ErrorClass(ERROR_INVALID_PASSWORD, "Error: Invalid Password"),
            new ErrorClass(ERROR_MORE_DATA, "Error: More Data"),
            new ErrorClass(ERROR_NO_MORE_ITEMS, "Error: No More Items"),
            new ErrorClass(ERROR_NO_NET_OR_BAD_PATH, "Error: No Net Or Bad Path"),
            new ErrorClass(ERROR_NO_NETWORK, "Error: No Network"),
            new ErrorClass(ERROR_BAD_PROFILE, "Error: Bad Profile"),
            new ErrorClass(ERROR_CANNOT_OPEN_PROFILE, "Error: Cannot Open Profile"),
            new ErrorClass(ERROR_DEVICE_IN_USE, "Error: Device In Use"),
            new ErrorClass(ERROR_EXTENDED_ERROR, "Error: Extended Error"),
            new ErrorClass(ERROR_NOT_CONNECTED, "Error: Not Connected"),
            new ErrorClass(ERROR_OPEN_FILES, "Error: Open Files"),
        };

        private static string getErrorForNumber(int errNum)
        {
            foreach (ErrorClass er in ERROR_LIST)
            {
                if (er.num == errNum) return er.message;
            }
            return "Error: Unknown, " + errNum;
        }
        #endregion

        [DllImport("Mpr.dll")]
        private static extern int WNetUseConnection(
            IntPtr hwndOwner,
            NETRESOURCE lpNetResource,
            string lpPassword,
            string lpUserID,
            int dwFlags,
            string lpAccessName,
            string lpBufferSize,
            string lpResult
        );

        [DllImport("Mpr.dll")]
        private static extern int WNetCancelConnection2(
            string lpName,
            int dwFlags,
            bool fForce
        );

        [StructLayout(LayoutKind.Sequential)]
        private class NETRESOURCE
        {
            public int dwScope = 0;
            public int dwType = 0;
            public int dwDisplayType = 0;
            public int dwUsage = 0;
            public string lpLocalName = "";
            public string lpRemoteName = "";
            public string lpComment = "";
            public string lpProvider = "";
        }


        public static string connectToRemote(string remoteUNC, string username, string password)
        {
            return connectToRemote(remoteUNC, username, password, false);
        }

        public static string connectToRemote(string remoteUNC, string username, string password, bool promptUser)
        {
            NETRESOURCE nr = new NETRESOURCE();
            nr.dwType = RESOURCETYPE_DISK;
            nr.lpRemoteName = remoteUNC;
            //			nr.lpLocalName = "F:";

            int ret;
            if (promptUser)
                ret = WNetUseConnection(IntPtr.Zero, nr, "", "", CONNECT_INTERACTIVE | CONNECT_PROMPT, null, null, null);
            else
                ret = WNetUseConnection(IntPtr.Zero, nr, password, username, 0, null, null, null);

            if (ret == NO_ERROR) return null;
            return getErrorForNumber(ret);
        }

        public static string disconnectRemote(string remoteUNC)
        {
            int ret = WNetCancelConnection2(remoteUNC, CONNECT_UPDATE_PROFILE, false);
            if (ret == NO_ERROR) return null;
            return getErrorForNumber(ret);
        }
    }
}

namespace ImpersonationsTools
{
    public class Impersonator
    {
        #region "Consts"

        public const int LOGON32_LOGON_INTERACTIVE = 2;

        public const int LOGON32_PROVIDER_DEFAULT = 0;
        #endregion

        #region "External API"
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool RevertToSelf();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern int CloseHandle(IntPtr hObject);

        #endregion

        #region "Methods"

        //Public Sub PerformImpersonatedTask(ByVal username As String, ByVal domain As String, ByVal password As String, ByVal logonType As Integer, ByVal logonProvider As Integer, ByVal methodToPerform As Action)
        public void PerformImpersonatedTask(string username, string domain, string password, int logonType, int logonProvider, Action methodToPerform)
        {
            IntPtr token = IntPtr.Zero;
            if (RevertToSelf())
            {
                if (LogonUser(username, domain, password, logonType, logonProvider, ref token) != 0)
                {
                    dynamic identity = new WindowsIdentity(token);
                    dynamic impersonationContext = identity.Impersonate();
                    if (impersonationContext != null)
                    {
                        methodToPerform.Invoke();
                        impersonationContext.Undo();
                    }
                    // do logging
                }
                else
                {
                }
            }
            if (token != IntPtr.Zero)
            {
                CloseHandle(token);
            }
        }

        #endregion
    }

    public class UserImpersonation
    {
        const int LOGON32_LOGON_INTERACTIVE = 2;
        const int LOGON32_LOGON_NETWORK = 3;
        const int LOGON32_LOGON_BATCH = 4;
        const int LOGON32_LOGON_SERVICE = 5;
        const int LOGON32_LOGON_UNLOCK = 7;
        const int LOGON32_LOGON_NETWORK_CLEARTEXT = 8;
        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_PROVIDER_WINNT35 = 1;
        const int LOGON32_PROVIDER_WINNT40 = 2;
        const int LOGON32_PROVIDER_WINNT50 = 3;

        WindowsImpersonationContext impersonationContext;
        [DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int LogonUserA(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]

        public static extern int DuplicateToken(IntPtr ExistingTokenHandle, int ImpersonationLevel, ref IntPtr DuplicateTokenHandle);
        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]

        public static extern long RevertToSelf();
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern long CloseHandle(IntPtr handle);

        public bool impersonateUser(string userName, string domain, string password)
        {
            return impersonateValidUser(userName, domain, password);
        }

        public void undoimpersonateUser()
        {
            undoImpersonation();
        }

        private bool impersonateValidUser(string userName, string domain, string password)
        {
            bool functionReturnValue = false;

            WindowsIdentity tempWindowsIdentity = null;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;
            functionReturnValue = false;

            //if (RevertToSelf()) {
            if (LogonUserA(userName, domain, password, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_WINNT50, ref token) != 0)
            {
                if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                {
                    tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                    impersonationContext = tempWindowsIdentity.Impersonate();
                    if ((impersonationContext != null))
                    {
                        functionReturnValue = true;
                    }
                }
            }
            //}
            if (!tokenDuplicate.Equals(IntPtr.Zero))
            {
                CloseHandle(tokenDuplicate);
            }
            if (!token.Equals(IntPtr.Zero))
            {
                CloseHandle(token);
            }
            return functionReturnValue;
        }

        private void undoImpersonation()
        {
            impersonationContext.Undo();
        }
    }
}

namespace Encryption
{
    public static class StringCipher
    {
        // This constant is used to determine the keysize of the encryption algorithm in bits.
        // We divide this by 8 within the code below to get the equivalent number of bytes.
        private const int Keysize = 256;

        // This constant determines the number of iterations for the password bytes generation function.
        private const int DerivationIterations = 1000;

        public static string Encrypt(string plainText, string passPhrase)
        {
            // Salt and IV is randomly generated each time, but is preprended to encrypted cipher text
            // so that the same Salt and IV values can be used when decrypting.  
            var saltStringBytes = Generate256BitsOfRandomEntropy();
            var ivStringBytes = Generate256BitsOfRandomEntropy();
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            using (var password = new Rfc2898DeriveBytes(passPhrase, saltStringBytes, DerivationIterations))
            {
                var keyBytes = password.GetBytes(Keysize / 8);
                using (var symmetricKey = new RijndaelManaged())
                {
                    symmetricKey.BlockSize = 256;
                    symmetricKey.Mode = CipherMode.CBC;
                    symmetricKey.Padding = PaddingMode.PKCS7;
                    using (var encryptor = symmetricKey.CreateEncryptor(keyBytes, ivStringBytes))
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                            {
                                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                                cryptoStream.FlushFinalBlock();
                                // Create the final bytes as a concatenation of the random salt bytes, the random iv bytes and the cipher bytes.
                                var cipherTextBytes = saltStringBytes;
                                cipherTextBytes = cipherTextBytes.Concat(ivStringBytes).ToArray();
                                cipherTextBytes = cipherTextBytes.Concat(memoryStream.ToArray()).ToArray();
                                memoryStream.Close();
                                cryptoStream.Close();
                                return Convert.ToBase64String(cipherTextBytes);
                            }
                        }
                    }
                }
            }
        }

        public static string Decrypt(string cipherText, string passPhrase)
        {
            // Get the complete stream of bytes that represent:
            // [32 bytes of Salt] + [32 bytes of IV] + [n bytes of CipherText]
            var cipherTextBytesWithSaltAndIv = Convert.FromBase64String(cipherText);
            // Get the saltbytes by extracting the first 32 bytes from the supplied cipherText bytes.
            var saltStringBytes = cipherTextBytesWithSaltAndIv.Take(Keysize / 8).ToArray();
            // Get the IV bytes by extracting the next 32 bytes from the supplied cipherText bytes.
            var ivStringBytes = cipherTextBytesWithSaltAndIv.Skip(Keysize / 8).Take(Keysize / 8).ToArray();
            // Get the actual cipher text bytes by removing the first 64 bytes from the cipherText string.
            var cipherTextBytes = cipherTextBytesWithSaltAndIv.Skip((Keysize / 8) * 2).Take(cipherTextBytesWithSaltAndIv.Length - ((Keysize / 8) * 2)).ToArray();

            using (var password = new Rfc2898DeriveBytes(passPhrase, saltStringBytes, DerivationIterations))
            {
                var keyBytes = password.GetBytes(Keysize / 8);
                using (var symmetricKey = new RijndaelManaged())
                {
                    symmetricKey.BlockSize = 256;
                    symmetricKey.Mode = CipherMode.CBC;
                    symmetricKey.Padding = PaddingMode.PKCS7;
                    using (var decryptor = symmetricKey.CreateDecryptor(keyBytes, ivStringBytes))
                    {
                        using (var memoryStream = new MemoryStream(cipherTextBytes))
                        {
                            using (var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                            {
                                var plainTextBytes = new byte[cipherTextBytes.Length];
                                var decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
                                memoryStream.Close();
                                cryptoStream.Close();
                                return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
                            }
                        }
                    }
                }
            }
        }

        private static byte[] Generate256BitsOfRandomEntropy()
        {
            var randomBytes = new byte[32]; // 32 Bytes will give us 256 bits.
            using (var rngCsp = new RNGCryptoServiceProvider())
            {
                // Fill the array with cryptographically secure random bytes.
                rngCsp.GetBytes(randomBytes);
            }
            return randomBytes;
        }
    }
}

public class WndSearcher
{
    public static IntPtr SearchForWindow(string wndclass, string title)
    {
        SearchData sd = new SearchData { Wndclass = wndclass, Title = title };
        EnumWindows(new EnumWindowsProc(EnumProc), ref sd);
        return sd.hWnd;
    }

    public static bool EnumProc(IntPtr hWnd, ref SearchData data)
    {
        // Check classname and title
        // This is different from FindWindow() in that the code below allows partial matches
        StringBuilder sb = new StringBuilder(1024);
        GetClassName(hWnd, sb, sb.Capacity);
        if (sb.ToString().StartsWith(data.Wndclass))
        {
            sb = new StringBuilder(1024);
            GetWindowText(hWnd, sb, sb.Capacity);
            if (sb.ToString().StartsWith(data.Title))
            {
                data.hWnd = hWnd;
                return false;    // Found the wnd, halt enumeration
            }
        }
        return true;
    }

    public class SearchData
    {
        // You can put any dicks or Doms in here...
        public string Wndclass;
        public string Title;
        public IntPtr hWnd;
    }

    private delegate bool EnumWindowsProc(IntPtr hWnd, ref SearchData data);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, ref SearchData data);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
}