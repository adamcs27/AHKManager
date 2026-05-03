using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.IO.Compression;
using System.Text;
using Microsoft.Win32;

namespace AHKScriptManager
{
    // ─── Data Models ────────────────────────────────────────────────────────────
    [Serializable]
    public class ScriptEntry
    {
        public string FilePath { get; set; }
        public string Description { get; set; }
        public List<string> Groups { get; set; } = new List<string>();

        [XmlIgnore]
        public string Name => Path.GetFileNameWithoutExtension(FilePath) ?? "";

        [XmlIgnore]
        public bool IsRunning { get; set; }

        [XmlIgnore]
        public Process RunningProcess { get; set; }

        [XmlIgnore]
        public List<string> ParsedHotkeys { get; set; } = new List<string>();

        public void ParseHotkeys()
        {
            ParsedHotkeys.Clear();
            if (!File.Exists(FilePath)) return;
            try
            {
                var lines = File.ReadAllLines(FilePath);
                var hotkeyPattern = new Regex(@"^([^;,\s][^,\n]*?)::(?!\s*=)", RegexOptions.Multiline);
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith(";")) continue;
                    var m = hotkeyPattern.Match(trimmed);
                    if (m.Success)
                    {
                        var hk = m.Groups[1].Value.Trim();
                        if (!hk.Contains(" ") || hk.StartsWith("^") || hk.StartsWith("!") || hk.StartsWith("+") || hk.StartsWith("#"))
                            ParsedHotkeys.Add(hk);
                    }
                }
                ParsedHotkeys = ParsedHotkeys.Distinct().Take(8).ToList();
            }
            catch { }
        }
    }

    [Serializable]
    public class AppData
    {
        public List<ScriptEntry> Scripts { get; set; } = new List<ScriptEntry>();
        public List<string> Groups { get; set; } = new List<string>();
        public string SuspendAllHotkey { get; set; } = "";
        public bool StartWithWindows { get; set; } = false;
        public bool MinimizeToTray { get; set; } = true;
        public bool DarkMode { get; set; } = false;
        public bool ShowTrayHint { get; set; } = true;
        public List<string> RunHistory { get; set; } = new List<string>();
        public bool TourCompleted { get; set; } = false;
    }

    public static class Theme
    {
        public static bool IsDark { get; set; } = false;

        public static Color Bg => IsDark ? Color.FromArgb(45, 45, 48) : SystemColors.Control;
        public static Color Fg => IsDark ? Color.White : SystemColors.ControlText;
        public static Color TextBg => IsDark ? Color.FromArgb(30, 30, 30) : SystemColors.Window;
        public static Color TextFg => IsDark ? Color.LightGray : SystemColors.WindowText;
        public static Color BtnBg => IsDark ? Color.FromArgb(60, 60, 60) : SystemColors.Control;
        public static Color TabDark => IsDark ? Color.FromArgb(30, 30, 30) : SystemColors.ControlDark;
        public static Color TabLight => IsDark ? Color.FromArgb(80, 80, 80) : SystemColors.ControlLightLight;
        public static Color ListRunningBg => IsDark ? Color.DarkGreen : Color.LightGreen;

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        public static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        public static void Apply(Control c, bool bypassButtons = false)
        {
            if (c is Form f) { f.BackColor = Bg; f.ForeColor = Fg; }
            else if (c is RichTextBox || c is TextBox || c is ListView) 
            { 
                c.BackColor = TextBg; c.ForeColor = TextFg; 
                if (c is ListView lv) { lv.Items.Cast<ListViewItem>().ToList().ForEach(i => { i.BackColor = TextBg; i.ForeColor = TextFg; }); }
                try { SetWindowTheme(c.Handle, IsDark ? "DarkMode_Explorer" : "Explorer", null); } catch { }
            }
            else if (c is Button b && !bypassButtons) { b.BackColor = BtnBg; b.ForeColor = Fg; b.FlatStyle = IsDark ? FlatStyle.Flat : FlatStyle.Standard; if (IsDark) b.FlatAppearance.BorderColor = Color.Gray; }
            else if (c is Panel || c is FlowLayoutPanel || c is TableLayoutPanel || c is TabPage) { c.BackColor = Bg; c.ForeColor = Fg; }
            else if (c is MenuStrip || c is ToolStrip) { c.BackColor = Bg; c.ForeColor = Fg; }

            foreach (Control child in c.Controls) Apply(child, bypassButtons);
        }
    }

    // ─── Dialog Helpers ───────────────────────────────────────────────────────────
    public static class Prompts
    {
        public static string ShowDialog(string text, string caption, string defaultValue = "")
        {
            Form prompt = new Form()
            {
                Width = 450,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 9),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            
            var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(20, 20, 20, 35), AutoSize = true };
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tlp.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label textLabel = new Label() { Text = text, AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
            TextBox textBox = new TextBox() { Text = defaultValue, Width = 400, Margin = new Padding(0, 5, 0, 15) };
            Button confirmation = new Button() { Text = "OK", Width = 110, Height = 36, DialogResult = DialogResult.OK, Anchor = AnchorStyles.Right, Margin = new Padding(0, 5, 0, 10) };
            
            tlp.Controls.Add(textLabel, 0, 0);
            tlp.Controls.Add(textBox, 0, 1);
            tlp.Controls.Add(confirmation, 0, 2);
            
            prompt.Controls.Add(tlp);
            prompt.AcceptButton = confirmation;

            Theme.Apply(prompt);

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        public static bool ShowCreateScript(out string title, out string description, out string content, bool isImport)
        {
            title = ""; description = ""; content = "";
            Form prompt = new Form()
            {
                Width = 500,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = isImport ? "Import Script from Paste" : "Create New Script",
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 9),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            
            var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7, Padding = new Padding(20, 20, 20, 35), AutoSize = true };
            
            Label lblTitle = new Label { Text = "Script Title / Filename:", AutoSize = true, Margin = new Padding(0,0,0,5) };
            TextBox txtTitle = new TextBox { Width = 440, Margin = new Padding(0, 0, 0, 10) };
            
            Label lblDesc = new Label { Text = "Description (Optional):", AutoSize = true, Margin = new Padding(0,0,0,5) };
            TextBox txtDesc = new TextBox { Width = 440, Margin = new Padding(0, 0, 0, 10) };

            var headerPanel = new FlowLayoutPanel { Width = 440, Height = 35, Margin = new Padding(0) };
            Label lblContent = new Label { Text = isImport ? "Paste script content below:" : "Template contents:", AutoSize = true, Margin = new Padding(0, 8, 10, 0) };
            headerPanel.Controls.Add(lblContent);

            RichTextBox txtContent = new RichTextBox
            {
                Width = 440,
                Height = 200,
                Margin = new Padding(0, 0, 0, 15),
                Font = new Font("Consolas", 8.5f),
                Text = isImport ? "" : "#SingleInstance, Force\nSetWorkingDir %A_ScriptDir%\n\n; Your hotkeys here\n"
            };

            if (isImport)
            {
                Button btnPaste = new Button { Text = "Paste from Clipboard", AutoSize = true, Height = 30 };
                btnPaste.Click += (s, e) => { txtContent.Text = Clipboard.GetText(); };
                headerPanel.Controls.Add(btnPaste);
            }

            Button confirmation = new Button { Text = isImport ? "Import" : "Create", Width = 110, Height = 36, DialogResult = DialogResult.OK, Anchor = AnchorStyles.Right, Margin = new Padding(0, 5, 0, 10) };
            
            tlp.Controls.Add(lblTitle, 0, 0);
            tlp.Controls.Add(txtTitle, 0, 1);
            tlp.Controls.Add(lblDesc, 0, 2);
            tlp.Controls.Add(txtDesc, 0, 3);
            tlp.Controls.Add(headerPanel, 0, 4);
            tlp.Controls.Add(txtContent, 0, 5);
            tlp.Controls.Add(confirmation, 0, 6);
            
            prompt.Controls.Add(tlp);
            prompt.AcceptButton = confirmation;

            Theme.Apply(prompt);

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                title = txtTitle.Text;
                description = txtDesc.Text;
                content = txtContent.Text;
                return !string.IsNullOrWhiteSpace(title);
            }
            return false;
        }

        public static bool ShowSettings(ref string suspendHotkey, ref bool startWithWindows, ref bool minimizeToTray, ref bool darkMode)
        {
            Form prompt = new Form()
            {
                Width = 400,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Settings",
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 9),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 8, Padding = new Padding(20, 20, 20, 35), AutoSize = true };
            
            Label lblHK = new Label { Text = "Global Suspend All Hotkey:", AutoSize = true, Margin = new Padding(0, 0, 0, 0) };
            Label lblMod = new Label { Text = "(Modifiers: ^ = Ctrl | + = Shift | ! = Alt | # = Win)\nExample: ^!m for Ctrl+Alt+m", Font = new Font("Segoe UI", 8, FontStyle.Italic), AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
            TextBox txtHK = new TextBox { Width = 340, Text = suspendHotkey, Margin = new Padding(0, 0, 0, 15) };
            
            CheckBox chkStart = new CheckBox { Text = "Start With Windows", AutoSize = true, Checked = startWithWindows, Margin = new Padding(0, 0, 0, 10) };
            CheckBox chkTray = new CheckBox { Text = "Minimize to System Tray Instead of Closing", AutoSize = true, Checked = minimizeToTray, Margin = new Padding(0, 0, 0, 10) };
            CheckBox chkDark = new CheckBox { Text = "Enable Dark Mode", AutoSize = true, Checked = darkMode, Margin = new Padding(0, 0, 0, 15) };
            
            Label lblCredit = new Label { Text = "Created by @27acs on Discord", Font = new Font("Segoe UI", 8, FontStyle.Bold), ForeColor = Color.Gray, AutoSize = true, Margin = new Padding(0, 5, 0, 0) };
            Button confirmation = new Button { Text = "Save", Width = 110, Height = 36, DialogResult = DialogResult.OK, Anchor = AnchorStyles.Right, Margin = new Padding(0, 5, 0, 10) };

            tlp.Controls.Add(lblHK, 0, 0);
            tlp.Controls.Add(lblMod, 0, 1);
            tlp.Controls.Add(txtHK, 0, 2);
            tlp.Controls.Add(chkStart, 0, 3);
            tlp.Controls.Add(chkTray, 0, 4);
            tlp.Controls.Add(chkDark, 0, 5);
            tlp.Controls.Add(lblCredit, 0, 6);
            tlp.Controls.Add(confirmation, 0, 7);
            
            prompt.Controls.Add(tlp);
            prompt.AcceptButton = confirmation;

            Theme.Apply(prompt);

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                suspendHotkey = txtHK.Text;
                startWithWindows = chkStart.Checked;
                minimizeToTray = chkTray.Checked;
                darkMode = chkDark.Checked;
                return true;
            }
            return false;
        }

        public static void ShowTrayHintDialog(Icon appIcon)
        {
            Form hint = new Form()
            {
                Width = 800,
                Height = 250,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "AHK Manager",
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false,
                Font = new Font("Segoe UI", 9),
                TopMost = true,
                Icon = appIcon,
                BackColor = Theme.Bg,
                ForeColor = Theme.Fg
            };

            var tlp = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(20) };
            Label lbl = new Label { 
                Text = "AHK Manager has been minimized to the System Tray.\n\nYou can fully exit via the tray icon or change this behavior in Settings.",
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };
            Button btn = new Button { 
                Text = "OK", 
                Width = 100, 
                Height = 32, 
                DialogResult = DialogResult.OK, 
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right 
            };

            tlp.Controls.Add(lbl, 0, 0);
            tlp.Controls.Add(btn, 0, 1);
            hint.Controls.Add(tlp);
            hint.AcceptButton = btn;

            Theme.Apply(hint);
            hint.ShowDialog();
        }
    }

    // ─── Main Form ──────────────────────────────────────────────────────────────
    public class MainForm : Form
    {
        public static MainForm Instance { get; private set; }

        // Data
        private AppData _data = new AppData();
        private string _dataPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AHKScriptManager", "data.xml");

        // UI
        private MenuStrip _menuStrip;
        private ToolStrip _toolStrip;
        private TabControl _tabControl;
        private ListView _listView;
        private ToolStripTextBox _searchBox;
        private ToolStripButton _btnToolbarRun;
        private ToolStripButton _btnToolbarReload;
        private ToolStripButton _btnToolbarStop;
        private Timer _statusTimer;
        private float _currentFontSize = 9f;

        // Side Panel
        private Panel _sidePanel;
        private RichTextBox _previewBox;
        private Label _lblTitle;
        private Label _lblHotkey;
        private Label _lblDesc;
        private Button _btnRun;
        private Button _btnReload;
        private Button _btnStop;
        private Button _btnSavePreview;
        private Button _btnOpenNotepad;
        private Button _btnShare;
        private Button _btnCreateNew;
        private Button _btnImportPaste;
        private Button _btnImportShare;
        private ToolStripButton _btnSettings;
        private ToolStripButton _btnRunAll;
        private ToolStripButton _btnReloadAll;
        private ToolStripButton _btnSuspendAll;
        private ScriptEntry _selectedScript;

        // Tray & Hotkeys
        private NotifyIcon _notifyIcon;
        [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);
        [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")] private static extern bool LockWindowUpdate(IntPtr hWndLock);
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_SUSPEND_ID = 9001;
        private bool _suspended = false;

        // Syntax Highlighting Debounce
        private Timer _highlightTimer;
        private bool _isHighlighting;
        private Form _fullViewForm;
        private TabPage _draggedTab;

        public MainForm()
        {
            Instance = this;
            LoadData(); 
            Theme.IsDark = _data.DarkMode;
            if (_data.ShowTrayHint) _data.MinimizeToTray = true;

            InitializeUI();
            RefreshGroups(); // Moved from LoadData to respect init hierarchy bindings safely
            RefreshView();
            Theme.Apply(this);
            SetupStatusTimer();

            _highlightTimer = new Timer { Interval = 600 };
            _highlightTimer.Tick += (s, e) => { _highlightTimer.Stop(); HighlightSyntax(); };

            if (Environment.GetCommandLineArgs().Contains("-minimized"))
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }
            else if (!_data.TourCompleted)
            {
                this.Load += (s, e) => StartTour();
            }
        }

        private void InitializeUI()
        {
            Text = "AHK Script Manager";
            Size = new Size(1200, 650);
            MinimumSize = new Size(800, 450);
            StartPosition = FormStartPosition.CenterScreen;
            SetAppIcon();

            // Tray
            _notifyIcon = new NotifyIcon { Icon = this.Icon, Visible = true, Text = "AHK Manager" };
            _notifyIcon.DoubleClick += (s, e) => { Show(); WindowState = FormWindowState.Normal; ShowInTaskbar = true; };

            var trayMenu = new ContextMenuStrip();
            trayMenu.Opening += (s, e) => {
                trayMenu.Items.Clear();
                trayMenu.Items.Add("Open", null, (ss, ee) => { Show(); WindowState = FormWindowState.Normal; ShowInTaskbar = true; });
                trayMenu.Items.Add(new ToolStripSeparator());
                
                if (_data.RunHistory.Count > 0)
                {
                    foreach (var group in _data.RunHistory)
                    {
                        var gCopy = group;
                        trayMenu.Items.Add($"Run All: {gCopy}", null, (ss, ee) => RunGroup(gCopy));
                    }
                    trayMenu.Items.Add(new ToolStripSeparator());
                }

                trayMenu.Items.Add("Suspend / Stop All", null, (ss, ee) => ToggleSuspendAll());
                trayMenu.Items.Add("Reload All", null, OnReloadAll);
                trayMenu.Items.Add(new ToolStripSeparator());
                trayMenu.Items.Add("Exit", null, (ss, ee) => {
                    _notifyIcon.Visible = false;
                    Application.Exit();
                });
            };
            _notifyIcon.ContextMenuStrip = trayMenu;

            // MenuStrip
            _menuStrip = new MenuStrip();
            var fileMenu = new ToolStripMenuItem("File");
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Add Scripts...", null, OnBrowse));
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Import from Share String...", null, OnImportFromShareString));
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(new ToolStripMenuItem("Exit", null, (s, e) => { _notifyIcon.Visible = false; Application.Exit(); }));

            var groupMenu = new ToolStripMenuItem("Groups");
            groupMenu.DropDownItems.Add(new ToolStripMenuItem("Add Group...", null, OnAddGroup));

            var helpMenu = new ToolStripMenuItem("Help");
            helpMenu.DropDownItems.Add(new ToolStripMenuItem("Features & Quick Start", null, (s, e) => StartTour()));

            _menuStrip.Items.Add(fileMenu);
            _menuStrip.Items.Add(groupMenu);
            _menuStrip.Items.Add(helpMenu);

            // ToolStrip
            _toolStrip = new ToolStrip();
            _btnToolbarRun = new ToolStripButton("▶ Run", null, OnToolbarRun) { DisplayStyle = ToolStripItemDisplayStyle.Text, Visible = false };
            _btnToolbarReload = new ToolStripButton("↻ Reload", null, OnToolbarReload) { DisplayStyle = ToolStripItemDisplayStyle.Text, Visible = false };
            _btnToolbarStop = new ToolStripButton("⏸ Stop", null, OnToolbarStop) { DisplayStyle = ToolStripItemDisplayStyle.Text, Visible = false };
            
            _toolStrip.Items.Add(_btnToolbarRun);
            _toolStrip.Items.Add(_btnToolbarReload);
            _toolStrip.Items.Add(_btnToolbarStop);
            _toolStrip.Items.Add(new ToolStripSeparator());
            _btnRunAll = new ToolStripButton("▶ Run All", null, OnRunAll) { DisplayStyle = ToolStripItemDisplayStyle.Text };
            _btnReloadAll = new ToolStripButton("↻ Reload All", null, OnReloadAll) { DisplayStyle = ToolStripItemDisplayStyle.Text };
            _btnSuspendAll = new ToolStripButton("⏸ Suspend / Stop All", null, (s, e) => ToggleSuspendAll()) { DisplayStyle = ToolStripItemDisplayStyle.Text };
            
            _toolStrip.Items.Add(_btnRunAll);
            _toolStrip.Items.Add(_btnReloadAll);
            _toolStrip.Items.Add(_btnSuspendAll);
            _toolStrip.Items.Add(new ToolStripSeparator());
            
            _toolStrip.Items.Add(new ToolStripLabel("Search:"));
            _searchBox = new ToolStripTextBox { Width = 150 };
            _searchBox.TextChanged += (s, e) => RefreshView();
            _toolStrip.Items.Add(_searchBox);
            
            _toolStrip.Items.Add(new ToolStripSeparator());
            _btnSettings = new ToolStripButton("⚙ Settings", null, OnSettings) { DisplayStyle = ToolStripItemDisplayStyle.Text };
            _toolStrip.Items.Add(_btnSettings);

            var contentPanel = new Panel { Dock = DockStyle.Fill };

            // Bottom Action Bar
            var bottomBar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(5, 5, 5, 20) };
            _btnCreateNew = new Button { Text = "Create New Script", AutoSize = true, Height = 32, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            _btnImportPaste = new Button { Text = "Import from Paste", AutoSize = true, Height = 32, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            _btnImportShare = new Button { Text = "Import from Share", AutoSize = true, Height = 32, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            
            _btnCreateNew.Click += OnCreateNewScript;
            _btnImportPaste.Click += OnImportFromPaste;
            _btnImportShare.Click += OnImportFromShareString;

            bottomBar.Controls.Add(_btnCreateNew);
            bottomBar.Controls.Add(_btnImportPaste);
            bottomBar.Controls.Add(_btnImportShare);
            
            bottomBar.Controls.Add(_btnShare);

            // Side Panel setup
            _sidePanel = new Panel { Dock = DockStyle.Right, Width = 350, Visible = false, BackColor = SystemColors.ControlLight, Padding = new Padding(10) };
            
            _lblTitle = new Label { Dock = DockStyle.Top, Height = 42, Font = new Font("Segoe UI", 12, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
            _lblHotkey = new Label { Dock = DockStyle.Top, Height = 32, Font = new Font("Segoe UI", 9.5f, FontStyle.Italic), AutoEllipsis = true, UseCompatibleTextRendering = true };
            _lblDesc = new Label { Dock = DockStyle.Top, AutoSize = true, MaximumSize = new Size(330, 0), Font = new Font("Segoe UI", 9), Padding = new Padding(0, 0, 0, 10) };
            
            _btnShare = new Button { 
                Text = "📤 Share", 
                Dock = DockStyle.Top, 
                Height = 32, 
                Font = new Font("Segoe UI", 8, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 10),
                Cursor = Cursors.Hand
            };
            _btnShare.Click += (s, e) => { if (_selectedScript != null) OnShareScript(_selectedScript); };
            
            var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 45, WrapContents = false, Padding = new Padding(0, 0, 0, 5) };
            _btnRun = new Button { Text = "Run", Width = 95, Height = 36, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            _btnReload = new Button { Text = "Reload", Width = 95, Height = 36, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            _btnStop = new Button { Text = "Stop", Width = 95, Height = 36, Font = new Font("Segoe UI", 9, FontStyle.Bold) };

            var divider = new Label { Dock = DockStyle.Top, Height = 2, BackColor = SystemColors.ControlDarkDark };

            var previewToolbar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 35, WrapContents = false, Padding = new Padding(0, 5, 0, 5) };
            _btnSavePreview = new Button { Text = "Save", Width = 80, Height = 25 };
            _btnOpenNotepad = new Button { Text = "Open Full View", Width = 110, Height = 25 };
            previewToolbar.Controls.Add(_btnSavePreview);
            previewToolbar.Controls.Add(_btnOpenNotepad);

            _previewBox = new RichTextBox { 
                Dock = DockStyle.Fill, 
                Font = new Font("Consolas", 9f), 
                BackColor = SystemColors.Window,
                DetectUrls = false,
                HideSelection = false
            };
            _previewBox.TextChanged += (s, e) => {
                if (_isHighlightingProcess(_previewBox)) return;
                _highlightTimer.Stop();
                _highlightTimer.Start();
            };

            _btnRun.Click += (s, e) => { if (_selectedScript != null) RunScript(_selectedScript); };
            _btnReload.Click += (s, e) => { if (_selectedScript != null) ReloadScript(_selectedScript); };
            _btnStop.Click += (s, e) => { if (_selectedScript != null) StopScript(_selectedScript); };
            _btnOpenNotepad.Click += (s, e) => { if (_selectedScript != null) ShowFullView(_selectedScript); };
            
            _previewBox.KeyDown += (s, e) => {
                if (e.Control && e.KeyCode == Keys.S)
                {
                    SaveCurrentPreview();
                    e.SuppressKeyPress = true;
                }
            };
            this.KeyPreview = true;
            this.KeyDown += (s, e) => {
                if (e.Control && e.KeyCode == Keys.F)
                {
                    _searchBox.Focus();
                    e.SuppressKeyPress = true;
                }
            };
            _btnSavePreview.Click += (s, e) => SaveCurrentPreview();

            btnPanel.Controls.Add(_btnRun);
            btnPanel.Controls.Add(_btnReload);
            btnPanel.Controls.Add(_btnStop);

            _sidePanel.Controls.Add(_previewBox);
            _sidePanel.Controls.Add(previewToolbar);
            _sidePanel.Controls.Add(divider);
            _sidePanel.Controls.Add(btnPanel);
            _sidePanel.Controls.Add(_btnShare);
            _sidePanel.Controls.Add(_lblDesc);
            _sidePanel.Controls.Add(_lblHotkey);
            _sidePanel.Controls.Add(_lblTitle);

            // TabControl
            _tabControl = new TabControl { 
                Dock = DockStyle.Fill,
                ItemSize = new Size(120, 28),
                SizeMode = TabSizeMode.Fixed,
                Font = new Font("Segoe UI", 9f)
            };
            _tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
            _tabControl.Padding = new Point(12, 6);
            _tabControl.DrawItem += (s, e) =>
            {
                var g = e.Graphics;
                var rect = e.Bounds;
                var tabText = _tabControl.TabPages[e.Index].Text;
                bool isSelected = e.Index == _tabControl.SelectedIndex;

                Color backColor = isSelected ? Theme.TabLight : Theme.TabDark;
                Color textColor = Theme.Fg;

                using (var b = new SolidBrush(backColor))
                    g.FillRectangle(b, rect);

                TextRenderer.DrawText(g, tabText, _tabControl.Font, rect, textColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            };
            _tabControl.SelectedIndexChanged += (s, e) => {
                if (_tabControl.SelectedTab != null) {
                    _listView.Parent = _tabControl.SelectedTab;
                    RefreshView();
                }
            };
            
            // Tab Drag & Drop + Context Menu
            _tabControl.AllowDrop = true;
            _tabControl.MouseDown += (s, e) => {
                _draggedTab = null;
                for (int i = 0; i < _tabControl.TabCount; i++) {
                    if (_tabControl.GetTabRect(i).Contains(e.Location)) {
                        if (e.Button == MouseButtons.Left) {
                            _draggedTab = _tabControl.TabPages[i];
                        } else if (e.Button == MouseButtons.Right) {
                            if (i > 0) ShowTabContextMenu(i, e.Location);
                        }
                        break;
                    }
                }
            };
            _tabControl.MouseMove += (s, e) => {
                if (e.Button == MouseButtons.Left && _draggedTab != null) {
                    _tabControl.DoDragDrop(_draggedTab, DragDropEffects.Move);
                }
            };
            _tabControl.DragOver += (s, e) => {
                if (e.Data.GetDataPresent(typeof(TabPage))) e.Effect = DragDropEffects.Move;
                else e.Effect = DragDropEffects.None;
            };
            _tabControl.DragDrop += (s, e) => {
                if (e.Data.GetDataPresent(typeof(TabPage))) {
                    TabPage draggedTab = (TabPage)e.Data.GetData(typeof(TabPage));
                    Point pt = _tabControl.PointToClient(new Point(e.X, e.Y));
                    for (int i = 0; i < _tabControl.TabCount; i++) {
                        if (_tabControl.GetTabRect(i).Contains(pt)) {
                            if (i == 0) return; // Don't move to "All"
                            
                            int oldIdx = _tabControl.TabPages.IndexOf(draggedTab) - 1;
                            int newIdx = i - 1;
                            
                            if (oldIdx >= 0 && newIdx >= 0 && oldIdx != newIdx) {
                                var groupName = _data.Groups[oldIdx];
                                _data.Groups.RemoveAt(oldIdx);
                                _data.Groups.Insert(newIdx, groupName);
                                SaveData();
                                RefreshGroups();
                                _tabControl.SelectedIndex = i;
                            }
                            break;
                        }
                    }
                }
            };

            // ListView
            _listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                AllowDrop = true,
                MultiSelect = true,
                OwnerDraw = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                HideSelection = false
            };

            _listView.DrawColumnHeader += (s, e) => e.DrawDefault = true;
            _listView.DrawItem += (s, e) => e.DrawDefault = true;
            _listView.DrawSubItem += (s, e) => e.DrawDefault = true;
            _listView.Paint += (s, e) => {
                using (var pen = new Pen(Theme.IsDark ? Color.FromArgb(80, 80, 80) : Color.DarkGray))
                {
                    // Draw a 1px line exactly at the bottom border of the header
                    // Header height is roughly 28-30px.
                    e.Graphics.DrawLine(pen, 0, 27, _listView.Width, 27);
                }
            };
            
            _listView.Columns.Add("Status", 80);
            _listView.Columns.Add("Name", 180);
            _listView.Columns.Add("Run Hotkey", 175);
            _listView.Columns.Add("Description", 250);
            _listView.Columns.Add("Groups", 120);

            _listView.DragEnter += OnDragEnter;
            _listView.DragDrop += OnDragDrop;
            _listView.MouseClick += OnListMouseClick;
            _listView.DoubleClick += OnListDoubleClick;
            _listView.SelectedIndexChanged += OnListSelectedIndexChanged;
            _listView.KeyDown += OnListKeyDown;
            _listView.Font = new Font(_listView.Font.FontFamily, _currentFontSize);

            contentPanel.Controls.Add(_tabControl);
            contentPanel.Controls.Add(_sidePanel);
            contentPanel.Controls.Add(bottomBar);

            Controls.Add(contentPanel);
            Controls.Add(_toolStrip);
            Controls.Add(_menuStrip);
            MainMenuStrip = _menuStrip;
        }

        private void SetAppIcon()
        {
            try
            {
                this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch { this.Icon = SystemIcons.Application; }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        private HashSet<RichTextBox> _activeHighlighters = new HashSet<RichTextBox>();
        private bool _isHighlightingProcess(RichTextBox box) => _activeHighlighters.Contains(box);

        private void HighlightRichTextBox(RichTextBox box)
        {
            if (box.Text.Length > 200000) return;

            _activeHighlighters.Add(box);
            int selStart = box.SelectionStart;
            int selLen = box.SelectionLength;

            LockWindowUpdate(box.Handle);

            box.SelectAll();
            box.SelectionColor = Theme.TextFg;
            box.SelectionFont = new Font(box.Font, FontStyle.Regular);

            string text = box.Text;

            var keywords = Regex.Matches(text, @"\b(return|run|sleep|send|click|loop|if|else|msgbox|exitapp)\b", RegexOptions.IgnoreCase);
            foreach (Match m in keywords) { box.Select(m.Index, m.Length); box.SelectionColor = Theme.IsDark ? Color.DeepSkyBlue : Color.Blue; }

            var hotkeys = Regex.Matches(text, @"^.*::", RegexOptions.Multiline);
            foreach (Match m in hotkeys) { box.Select(m.Index, m.Length); box.SelectionColor = Theme.IsDark ? Color.Orange : Color.DarkOrange; box.SelectionFont = new Font(box.Font, FontStyle.Bold); }

            var comments = Regex.Matches(text, @"(?:^|\s);.*?$", RegexOptions.Multiline);
            foreach (Match m in comments) { box.Select(m.Index, m.Length); box.SelectionColor = Theme.IsDark ? Color.LightGreen : Color.ForestGreen; box.SelectionFont = new Font(box.Font, FontStyle.Italic); }

            box.Select(selStart, selLen);
            box.SelectionColor = Theme.TextFg;
            box.SelectionFont = new Font(box.Font, FontStyle.Regular);

            LockWindowUpdate(IntPtr.Zero);
            _activeHighlighters.Remove(box);
        }

        private void HighlightSyntax() => HighlightRichTextBox(_previewBox);

        private void SaveCurrentPreview()
        {
            if (_selectedScript != null)
            {
                try
                {
                    File.WriteAllText(_selectedScript.FilePath, _previewBox.Text);
                    _selectedScript.ParseHotkeys();
                    RefreshView();
                    MessageBox.Show("Script saved successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Failed to save: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        public void ShowFullView(ScriptEntry script)
        {
            if (!File.Exists(script.FilePath)) return;

            if (_fullViewForm != null && !_fullViewForm.IsDisposed) {
                _fullViewForm.Close();
            }

            var fullForm = new Form
            {
                Text = script.Name + " - Full View",
                Width = 600,
                Height = 600,
                StartPosition = FormStartPosition.Manual,
                ShowIcon = false,
                Font = new Font("Segoe UI", 9)
            };
            
            fullForm.Location = new Point(this.Right, this.Top);

            var moveHandler = new EventHandler((s, e) => {
                if (!fullForm.IsDisposed) fullForm.Location = new Point(this.Right, this.Top);
            });
            this.Move += moveHandler;
            this.Resize += moveHandler;

            fullForm.FormClosed += (s, e) => {
                this.Move -= moveHandler;
                this.Resize -= moveHandler;
                _fullViewForm = null;
            };

            var topBar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(10) };
            var btnSave = new Button { Text = "Save", Width = 80, Height = 35 };
            var btnUndo = new Button { Text = "Undo", Width = 80, Height = 35 };
            var btnRedo = new Button { Text = "Redo", Width = 80, Height = 35 };
            var btnNotepad = new Button { Text = "Open in Notepad", AutoSize = true, Height = 35 };

            _fullViewForm = fullForm;

            var rtb = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10.5f),
                HideSelection = false,
                Text = File.ReadAllText(script.FilePath)
            };

            Action saveAction = () => {
                File.WriteAllText(script.FilePath, rtb.Text);
                script.ParseHotkeys();
                RefreshView();
                if (_selectedScript == script) _previewBox.Text = rtb.Text;
                MessageBox.Show("Script saved successfully.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            btnUndo.Click += (s, e) => rtb.Undo();
            btnRedo.Click += (s, e) => rtb.Redo();
            btnNotepad.Click += (s, e) => { EditScript(script); };
            btnSave.Click += (s, e) => saveAction();
            
            rtb.KeyDown += (s, e) => {
                if (e.Control && e.KeyCode == Keys.S)
                {
                    saveAction();
                    e.SuppressKeyPress = true;
                }
            };

            var tmr = new Timer { Interval = 600 };
            tmr.Tick += (s, e) => { tmr.Stop(); HighlightRichTextBox(rtb); };

            rtb.TextChanged += (s, e) => {
                if (_isHighlightingProcess(rtb)) return;
                tmr.Stop();
                tmr.Start();
            };

            topBar.Controls.Add(btnSave);
            topBar.Controls.Add(btnUndo);
            topBar.Controls.Add(btnRedo);
            topBar.Controls.Add(btnNotepad);

            fullForm.Controls.Add(rtb);
            fullForm.Controls.Add(topBar);

            Theme.Apply(fullForm);

            fullForm.Shown += (s, e) => HighlightRichTextBox(rtb);
            fullForm.Show(this);
        }

        private void SetupStatusTimer()
        {
            _statusTimer = new Timer { Interval = 1500 };
            _statusTimer.Tick += (s, e) => CheckProcessStatuses();
            _statusTimer.Start();
        }

        private void CheckProcessStatuses()
        {
            if (this.WindowState == FormWindowState.Minimized) return;
            bool changed = false;
            var runningPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE Name LIKE '%AutoHotkey%'"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var cmdLine = obj["CommandLine"]?.ToString();
                        if (cmdLine != null)
                        {
                            foreach (var s in _data.Scripts)
                            {
                                if (cmdLine.IndexOf(s.FilePath, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    cmdLine.IndexOf(s.Name + ".ahk", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    runningPaths.Add(s.FilePath);
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            foreach (var script in _data.Scripts)
            {
                bool isRunningNow = runningPaths.Contains(script.FilePath);
                if (script.IsRunning != isRunningNow)
                {
                    script.IsRunning = isRunningNow;
                    changed = true;
                }
            }

            if (changed) RefreshStatusesInView();
        }

        private void RefreshStatusesInView()
        {
            if (this.WindowState == FormWindowState.Minimized) return;
            foreach (ListViewItem item in _listView.Items)
            {
                if (item.Tag is ScriptEntry s)
                {
                    item.SubItems[0].Text = s.IsRunning ? "▶ Running" : "⏸ Stopped";
                    item.UseItemStyleForSubItems = true;
                    item.BackColor = s.IsRunning ? Theme.ListRunningBg : Theme.TextBg;
                    item.ForeColor = s.IsRunning ? (Theme.IsDark ? Color.White : Color.Black) : Theme.TextFg;
                }
            }

            if (_sidePanel.Visible && _selectedScript != null)
            {
                _btnRun.Enabled = !_selectedScript.IsRunning;
                _btnStop.Enabled = _selectedScript.IsRunning;
                _btnReload.Enabled = _selectedScript.IsRunning;
            }
            UpdateToolbarButtons();
        }

        public void RefreshView()
        {
            var filterText = _searchBox.Text.ToLower();
            var selectedGroup = _tabControl.SelectedTab?.Name ?? "All";

            var previouslySelected = _listView.SelectedItems.Count > 0 ? _listView.SelectedItems[0].Tag as ScriptEntry : null;

            _listView.BeginUpdate();
            _listView.Items.Clear();

            var scripts = _data.Scripts.AsEnumerable();
            if (selectedGroup != "All")
                scripts = scripts.Where(s => s.Groups.Contains(selectedGroup));

            if (!string.IsNullOrWhiteSpace(filterText))
                scripts = scripts.Where(s => s.Name.ToLower().Contains(filterText) || (s.Description ?? "").ToLower().Contains(filterText));

            foreach (var s in scripts)
            {
                var item = new ListViewItem(s.IsRunning ? "▶ Running" : "⏸ Stopped");
                item.SubItems.Add(s.Name);
                item.SubItems.Add(string.Join(", ", s.ParsedHotkeys));
                item.SubItems.Add(s.Description ?? "");
                item.SubItems.Add(string.Join(", ", s.Groups));
                item.Tag = s;
                item.UseItemStyleForSubItems = true;
                item.BackColor = s.IsRunning ? Theme.ListRunningBg : Theme.TextBg;
                item.ForeColor = s.IsRunning ? (Theme.IsDark ? Color.White : Color.Black) : Theme.TextFg;
                if (previouslySelected == s) item.Selected = true;
                _listView.Items.Add(item);
            }

            _listView.EndUpdate();
            OnListSelectedIndexChanged(null, null);
        }

        public void RefreshGroups()
        {
            var current = _tabControl.SelectedTab?.Name ?? "All";
            _tabControl.TabPages.Clear();
            
            _tabControl.TabPages.Add("All", "   All   ");
            foreach (var g in _data.Groups) _tabControl.TabPages.Add(g, $"   {g}   ");

            foreach (TabPage tp in _tabControl.TabPages) { tp.BackColor = Theme.Bg; tp.ForeColor = Theme.Fg; }

            var target = _tabControl.TabPages[current] ?? _tabControl.TabPages["All"];
            if (target != null)
            {
                _tabControl.SelectedTab = target;
                _listView.Parent = target;
            }
        }

        private void ShowTabContextMenu(int index, Point loc)
        {
            var cms = new ContextMenuStrip();
            var groupName = _data.Groups[index - 1];
            
            cms.Items.Add($"Delete Group '{groupName}'", null, (s, e) => {
                if (MessageBox.Show($"Are you sure you want to delete the group '{groupName}'?\nScripts will not be deleted, only the group category.", "Delete Group", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    _data.Groups.Remove(groupName);
                    foreach (var script in _data.Scripts) {
                        script.Groups.Remove(groupName);
                    }
                    SaveData();
                    RefreshGroups();
                    RefreshView();
                }
            });
            
            cms.Show(_tabControl, loc);
        }

        private void OnListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && (e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add))
            {
                _currentFontSize += 1f;
                if (_currentFontSize > 36f) _currentFontSize = 36f;
                _listView.Font = new Font(_listView.Font.FontFamily, _currentFontSize);
                e.Handled = e.SuppressKeyPress = true;
            }
            else if (e.Control && (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract))
            {
                _currentFontSize -= 1f;
                if (_currentFontSize < 6f) _currentFontSize = 6f;
                _listView.Font = new Font(_listView.Font.FontFamily, _currentFontSize);
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private void OnListSelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateToolbarButtons();
            if (_listView.SelectedItems.Count == 1 && _listView.SelectedItems[0].Tag is ScriptEntry s)
            {
                _selectedScript = s;
                _lblTitle.Text = s.Name;
                _lblHotkey.Text = s.ParsedHotkeys.Count > 0 ? "Hotkeys: " + string.Join(", ", s.ParsedHotkeys) : "No discrete hotkeys";
                _lblDesc.Text = !string.IsNullOrEmpty(s.Description) ? s.Description : "(No Description)";

                try
                {
                    if (File.Exists(s.FilePath))
                    {
                        var info = new FileInfo(s.FilePath);
                        if (info.Length < 100000)
                        {
                            _activeHighlighters.Add(_previewBox);
                            _previewBox.Text = File.ReadAllText(s.FilePath);
                            _activeHighlighters.Remove(_previewBox);
                            
                            _highlightTimer.Stop();
                            _highlightTimer.Start(); // highlight freshly loaded
                        }
                        else
                            _previewBox.Text = "... File too large for preview ...";
                    }
                    else _previewBox.Text = "... File missing ...";
                }
                catch { _previewBox.Text = ""; }

                _btnRun.Enabled = !s.IsRunning;
                _btnStop.Enabled = s.IsRunning;
                _btnReload.Enabled = s.IsRunning;
                
                _sidePanel.Visible = true;
            }
            else if (_listView.SelectedItems.Count > 1)
            {
                _selectedScript = null;
                _sidePanel.Visible = false;
            }
            // If SelectedItems.Count == 0, we do nothing to let the preview panel remain active
        }

        private void OnListMouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (_listView.HitTest(e.Location).Item?.Tag is ScriptEntry script)
                {
                    var cms = BuildContextMenu(script);
                    cms.Show(_listView, e.Location);
                }
            }
        }

        private void OnListDoubleClick(object sender, EventArgs e)
        {
            if (_listView.SelectedItems.Count > 0 && _listView.SelectedItems[0].Tag is ScriptEntry s)
            {
                if (s.IsRunning) StopScript(s); else RunScript(s);
            }
        }

        private ContextMenuStrip BuildContextMenu(ScriptEntry script)
        {
            var cms = new ContextMenuStrip();
            cms.Items.Add("Run", null, (s, e) => RunScript(script));
            cms.Items.Add("Reload", null, (s, e) => ReloadScript(script));
            cms.Items.Add("Stop", null, (s, e) => StopScript(script));

            cms.Items.Add("Edit", null, (s, e) => EditScript(script));
            cms.Items.Add("Share Script", null, (s, e) => OnShareScript(script));
            cms.Items.Add(new ToolStripSeparator());
            cms.Items.Add("Set Description...", null, (s, e) =>
            {
                string desc = Prompts.ShowDialog("Enter description:", "Description", script.Description ?? "");
                if (desc != "")
                {
                    script.Description = desc;
                    SaveData();
                    RefreshView();
                }
            });

            var groupItem = new ToolStripMenuItem("Send to Group");
            foreach (var g in _data.Groups)
            {
                var gCopy = g;
                var gi = new ToolStripMenuItem(gCopy);
                gi.Checked = script.Groups.Contains(gCopy);
                gi.Click += (ss, ee) =>
                {
                    if (script.Groups.Contains(gCopy)) script.Groups.Remove(gCopy);
                    else script.Groups.Add(gCopy);
                    SaveData();
                    RefreshView();
                };
                groupItem.DropDownItems.Add(gi);
            }
            
            groupItem.DropDownItems.Add(new ToolStripSeparator());
            var newGroupItem = new ToolStripMenuItem("Add to New Group...");
            newGroupItem.Click += (ss, ee) => {
                string g = Prompts.ShowDialog("Enter new group name:", "New Group");
                if (!string.IsNullOrWhiteSpace(g)) {
                    if (!_data.Groups.Contains(g.Trim())) _data.Groups.Add(g.Trim());
                    if (!script.Groups.Contains(g.Trim())) script.Groups.Add(g.Trim());
                    SaveData();
                    RefreshGroups();
                    RefreshView();
                }
            };
            groupItem.DropDownItems.Add(newGroupItem);

            cms.Items.Add(groupItem);
            cms.Items.Add(new ToolStripSeparator());
            cms.Items.Add("Open Folder", null, (s, e) => {
                if (File.Exists(script.FilePath)) Process.Start("explorer.exe", $"/select,\"{script.FilePath}\"");
            });
            cms.Items.Add(new ToolStripSeparator());
            cms.Items.Add("Remove", null, (s, e) => {
                if (script.IsRunning) StopScript(script);
                if (script == _selectedScript) { _selectedScript = null; _sidePanel.Visible = false; }
                _data.Scripts.Remove(script);
                SaveData();
                RefreshView();
            });

            return cms;
        }

        public void RunScript(ScriptEntry script)
        {
            if (script.IsRunning) return;

            if (!File.Exists(script.FilePath))
            {
                MessageBox.Show("File not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(script.FilePath) { UseShellExecute = true });
                script.IsRunning = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not launch script. Is AutoHotkey installed?\n" + ex.Message, "Launch Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            RefreshStatusesInView();
        }

        private void StopScript(ScriptEntry script)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name LIKE '%AutoHotkey%'"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        var cmdLine = obj["CommandLine"]?.ToString();
                        if (cmdLine != null && (cmdLine.IndexOf(script.FilePath, StringComparison.OrdinalIgnoreCase) >= 0 || 
                                                cmdLine.IndexOf(script.Name + ".ahk", StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            int pid = Convert.ToInt32(obj["ProcessId"]);
                            try { Process.GetProcessById(pid).Kill(); } catch { }
                        }
                    }
                }
            }
            catch { }
            script.IsRunning = false;
            RefreshStatusesInView();
        }

        public void ReloadScript(ScriptEntry script)
        {
            StopScript(script);
            System.Threading.Thread.Sleep(300);
            RunScript(script);
        }

        public void EditScript(ScriptEntry script)
        {
            if (!File.Exists(script.FilePath)) return;
            try 
            { 
                Process.Start(new ProcessStartInfo("notepad.exe", $"\"{script.FilePath}\"") { UseShellExecute = true }); 
            }
            catch (Exception ex) 
            { 
                MessageBox.Show("Failed to open notepad: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
        }

        private void OnRunAll(object sender, EventArgs e)
        {
            string groupName = _tabControl.SelectedTab?.Name ?? "All";
            if (groupName == "All")
            {
                if (MessageBox.Show("Are you sure you want to run ALL scripts across all groups?", "Run All", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    return;
            }
            else
            {
                // Update History
                _data.RunHistory.Remove(groupName);
                _data.RunHistory.Insert(0, groupName);
                if (_data.RunHistory.Count > 3) _data.RunHistory.RemoveAt(3);
                SaveData();
            }

            foreach (ListViewItem item in _listView.Items)
                if (item.Tag is ScriptEntry s && !s.IsRunning) RunScript(s);
        }

        private void RunGroup(string groupName)
        {
            foreach (var s in _data.Scripts)
            {
                if (s.Groups.Contains(groupName) && !s.IsRunning) RunScript(s);
            }
        }

        private void OnReloadAll(object sender, EventArgs e)
        {
            foreach (ListViewItem item in _listView.Items)
                if (item.Tag is ScriptEntry s && s.IsRunning) ReloadScript(s);
        }

        private void OnToolbarRun(object sender, EventArgs e)
        {
            foreach (ListViewItem item in _listView.SelectedItems)
                if (item.Tag is ScriptEntry s) RunScript(s);
        }

        private void OnToolbarReload(object sender, EventArgs e)
        {
            foreach (ListViewItem item in _listView.SelectedItems)
                if (item.Tag is ScriptEntry s) ReloadScript(s);
        }

        private void OnToolbarStop(object sender, EventArgs e)
        {
            foreach (ListViewItem item in _listView.SelectedItems)
                if (item.Tag is ScriptEntry s) StopScript(s);
        }

        private void OnShareScript(ScriptEntry s)
        {
            if (!File.Exists(s.FilePath)) return;
            try
            {
                string content = File.ReadAllText(s.FilePath);
                string raw = $"{s.Name}|{s.Description}|{content}";
                
                byte[] data = Encoding.UTF8.GetBytes(raw);
                using (var ms = new MemoryStream())
                {
                    using (var gz = new GZipStream(ms, CompressionMode.Compress))
                    {
                        gz.Write(data, 0, data.Length);
                    }
                    string shareString = "AHKM_v1:" + Convert.ToBase64String(ms.ToArray());
                    Clipboard.SetText(shareString);
                    MessageBox.Show("Share string copied to clipboard!\n\nYour friends can use 'File > Import from Share String' to add it.", "Share Script", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { MessageBox.Show("Failed to share: " + ex.Message); }
        }

        private void OnImportFromShareString(object sender, EventArgs e)
        {
            string input = Prompts.ShowDialog("Paste the AHKM sharing string here:", "Import Script");
            if (string.IsNullOrWhiteSpace(input)) return;

            try
            {
                if (!input.StartsWith("AHKM_v1:")) throw new Exception("Invalid sharing string format.");
                
                string base64 = input.Substring(8);
                byte[] compressed = Convert.FromBase64String(base64);
                
                using (var ms = new MemoryStream(compressed))
                using (var gz = new GZipStream(ms, CompressionMode.Decompress))
                using (var resultMs = new MemoryStream())
                {
                    gz.CopyTo(resultMs);
                    string raw = Encoding.UTF8.GetString(resultMs.ToArray());
                    string[] parts = raw.Split(new[] { '|' }, 3);
                    
                    if (parts.Length < 3) throw new Exception("Data corruption in share string.");
                    
                    string name = parts[0];
                    string desc = parts[1];
                    string content = parts[2];
                    
                    GenerateScriptAndAdd(name, desc, content);
                }
            }
            catch (Exception ex) { MessageBox.Show("Import failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void UpdateToolbarButtons()
        {
            int count = _listView.SelectedItems.Count;
            if (count == 0)
            {
                _btnToolbarRun.Visible = false;
                _btnToolbarReload.Visible = false;
                _btnToolbarStop.Visible = false;
            }
            else
            {
                bool anyRunning = false;
                foreach (ListViewItem item in _listView.SelectedItems)
                    if (item.Tag is ScriptEntry s && s.IsRunning) { anyRunning = true; break; }

                _btnToolbarRun.Visible = true;
                _btnToolbarReload.Visible = true;
                _btnToolbarStop.Visible = anyRunning;
                
                _btnToolbarRun.Text = count > 1 ? "▶ Run Selected" : "▶ Run";
                _btnToolbarReload.Text = count > 1 ? "↻ Reload Selected" : "↻ Reload";
                _btnToolbarStop.Text = count > 1 ? "⏸ Stop Selected" : "⏸ Stop";
            }
        }

        private void OnBrowse(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog { Filter = "AHK Scripts (*.ahk)|*.ahk|All (*.*)|*.*", Multiselect = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK) AddScripts(ofd.FileNames);
            }
        }
        
        private void OnCreateNewScript(object sender, EventArgs e)
        {
            if (Prompts.ShowCreateScript(out string title, out string desc, out string defaultContent, false))
            {
                GenerateScriptAndAdd(title, desc, defaultContent);
            }
        }

        private void OnImportFromPaste(object sender, EventArgs e)
        {
            if (Prompts.ShowCreateScript(out string title, out string desc, out string pastedContent, true))
            {
                GenerateScriptAndAdd(title, desc, pastedContent);
            }
        }
        
        private void GenerateScriptAndAdd(string title, string description, string content)
        {
            try
            {
                string safeTitle = string.Join("_", title.Split(Path.GetInvalidFileNameChars()));
                if (!safeTitle.EndsWith(".ahk", StringComparison.OrdinalIgnoreCase)) safeTitle += ".ahk";

                string scriptsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AHKScripts");
                Directory.CreateDirectory(scriptsDir);
                
                string fullPath = Path.Combine(scriptsDir, safeTitle);
                if (File.Exists(fullPath)) 
                {
                    if (MessageBox.Show("File already exists. Overwrite?", "Warning", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
                }

                File.WriteAllText(fullPath, content);
                
                var entry = new ScriptEntry { FilePath = fullPath, Description = description };
                entry.ParseHotkeys();
                _data.Scripts.Add(entry);
                SaveData();
                RefreshView();

                // Select the new script so preview opens correctly
                _listView.SelectedItems.Clear();
                foreach (ListViewItem item in _listView.Items)
                {
                    if (item.Tag == entry)
                    {
                        item.Selected = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to create script physically: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddScripts(files.Where(f => f.EndsWith(".ahk", StringComparison.OrdinalIgnoreCase)).ToArray());
        }

        private void AddScripts(string[] paths)
        {
            bool any = false;
            foreach (var path in paths)
            {
                if (_data.Scripts.Any(s => s.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase))) continue;
                var entry = new ScriptEntry { FilePath = path };
                entry.ParseHotkeys();
                _data.Scripts.Add(entry);
                any = true;
            }
            if (any) { SaveData(); RefreshView(); }
        }

        private void OnAddGroup(object sender, EventArgs e)
        {
            string g = Prompts.ShowDialog("Enter new group name:", "New Group");
            if (!string.IsNullOrWhiteSpace(g) && !_data.Groups.Contains(g.Trim()))
            {
                _data.Groups.Add(g.Trim());
                SaveData();
                RefreshGroups();
            }
        }

        private void OnSettings(object sender, EventArgs e)
        {
            bool wasStartWithWindows = _data.StartWithWindows;
            bool wasMinimizeToTray = _data.MinimizeToTray;
            bool wasDarkMode = _data.DarkMode;
            string hk = _data.SuspendAllHotkey;

            if (Prompts.ShowSettings(ref hk, ref wasStartWithWindows, ref wasMinimizeToTray, ref wasDarkMode))
            {
                _data.SuspendAllHotkey = hk;
                _data.StartWithWindows = wasStartWithWindows;
                _data.MinimizeToTray = wasMinimizeToTray;
                
                if (_data.DarkMode != wasDarkMode)
                {
                    _data.DarkMode = wasDarkMode;
                    Theme.IsDark = wasDarkMode;
                    Theme.Apply(this);
                    HighlightSyntax(); // Re-trigger mapping
                    if (_fullViewForm != null && !_fullViewForm.IsDisposed) Theme.Apply(_fullViewForm);
                    _tabControl.Invalidate();
                }

                SaveData();
                RegisterSuspendHotkey();
                ManageRegistryStartup();
            }
        }

        // ─── Guided Tour ────────────────────────────────────────────────────────────
        private Form _tourForm;
        private Form _highlightForm;
        private Label _tourText;
        private int _tourStep = 0;
        private object _tourTarget; 

        private void StartTour()
        {
            if (_tourForm != null) return;

            _tourForm = new Form {
                FormBorderStyle = FormBorderStyle.None,
                MinimumSize = new Size(320, 100),
                MaximumSize = new Size(400, 600),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Theme.IsDark ? Color.FromArgb(45, 45, 48) : Color.White,
                TopMost = true,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual,
                Padding = new Padding(15)
            };
            // Add a border
            _tourForm.Paint += (s, e) => {
                using (var pen = new Pen(Color.Gold, 2))
                    e.Graphics.DrawRectangle(pen, 0, 0, _tourForm.Width - 1, _tourForm.Height - 1);
            };

            // Internal layout to stack text and buttons
            var mainLayout = new FlowLayoutPanel {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            _tourForm.Controls.Add(mainLayout);

            _tourText = new Label { 
                AutoSize = true,
                MaximumSize = new Size(340, 0),
                ForeColor = Theme.IsDark ? Color.White : Color.Black,
                Font = new Font("Segoe UI", 10),
                TextAlign = ContentAlignment.TopLeft,
                Margin = new Padding(0, 0, 0, 10)
            };
            mainLayout.Controls.Add(_tourText);

            var navPanel = new FlowLayoutPanel { 
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Width = 340,
                Margin = new Padding(0, 5, 0, 0)
            };
            var btnNext = new Button { Text = "Next >", Width = 80, Height = 32 };
            var btnBack = new Button { Text = "< Back", Width = 80, Height = 32 };
            var btnClose = new Button { Text = "Exit Tour", Width = 85, Height = 32 };

            btnNext.Click += (s, e) => { _tourStep++; ShowTourStep(); };
            btnBack.Click += (s, e) => { if (_tourStep > 0) { _tourStep--; ShowTourStep(); } };
            btnClose.Click += (s, e) => EndTour();

            navPanel.Controls.Add(btnNext);
            navPanel.Controls.Add(btnBack);
            navPanel.Controls.Add(btnClose);
            mainLayout.Controls.Add(navPanel);

            _highlightForm = new Form {
                FormBorderStyle = FormBorderStyle.None,
                BackColor = Color.Magenta,
                TransparencyKey = Color.Magenta,
                TopMost = true,
                ShowInTaskbar = false,
                StartPosition = FormStartPosition.Manual
            };
            // Make click-through (WS_EX_TRANSPARENT = 0x20)
            int exStyle = (int)GetWindowLong(_highlightForm.Handle, -20);
            SetWindowLong(_highlightForm.Handle, -20, (IntPtr)(exStyle | 0x20));

            _highlightForm.Paint += (s, e) => {
                using (var pen = new Pen(Color.Gold, 4))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, _highlightForm.Width - 1, _highlightForm.Height - 1);
                }
            };

            _tourForm.Show();
            _highlightForm.Show();

            _tourStep = 0;
            ShowTourStep();
            
            this.Move += (s, e) => { UpdateTourPositions(); };
            this.Resize += (s, e) => { UpdateTourPositions(); };
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        static extern IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        private void UpdateTourPositions()
        {
            if (_tourForm == null) return;
            
            var rect = GetTargetRectScreen();
            if (!rect.IsEmpty)
            {
                _highlightForm.Bounds = rect;
            }

            // Position tour window outside: preferably to the right of the main form
            int x = this.Right + 10;
            int y = this.Top + 100;

            // If it goes off screen on the right, put it on the left
            if (x + _tourForm.Width > Screen.FromControl(this).WorkingArea.Right)
                x = this.Left - _tourForm.Width - 10;
            
            // If it's still weird, put it inside at the bottom
            if (x < Screen.FromControl(this).WorkingArea.Left)
                x = this.Left + 10;

            _tourForm.Location = new Point(x, y);
        }

        private void EndTour()
        {
            if (_tourForm != null) { _tourForm.Close(); _tourForm = null; }
            if (_highlightForm != null) { _highlightForm.Close(); _highlightForm = null; }
            _data.TourCompleted = true;
            SaveData();
        }

        private Control _spotlightControl; // Obsolete but keeping for now or removing
        private void ShowTourStep()
        {
            string msg = "";
            Control target = null;

            switch (_tourStep)
            {
                case 0:
                    msg = "Drag and drop scripts here, or add them with File > Add Scripts";
                    target = _listView;
                    break;
                case 1:
                    msg = "Create groups for scripts and organize them. Right-click any script > 'Send to Group' to move it.";
                    target = _tabControl;
                    break;
                case 2:
                    msg = "Run All: Runs every script in view\nReload All: Restarts all running scripts\nSuspend/Stop All: Halts all execution";
                    target = _toolStrip;
                    break;
                case 3:
                    msg = "Settings: Set your global 'Suspend All' keybind here for emergency stops.";
                    target = _toolStrip; // Highlighting settings button on strip
                    break;
                case 4:
                    msg = "Create New: Build a fresh script from scratch directly in the manager.";
                    target = _btnCreateNew;
                    break;
                case 5:
                    msg = "Import from Paste: Instantly load script code currently in your clipboard.";
                    target = _btnImportPaste;
                    break;
                case 6:
                    msg = "Import from Share: Paste an AHKM string to instantly get a script and its description from a friend.";
                    target = _btnImportShare;
                    break;
                default:
                    EndTour();
                    return;
            }

            _tourText.Text = msg;
            _tourTarget = target;
            
            // Special cases for ToolStripItems
            if (_tourStep == 2 && _btnRunAll != null) _tourTarget = _btnRunAll;
            if (_tourStep == 3 && _btnSettings != null) _tourTarget = _btnSettings;

            // Update positions
            UpdateTourPositions();
        }

        private Rectangle GetTargetRectScreen()
        {
            if (_tourTarget == null) return Rectangle.Empty;

            Rectangle rect = Rectangle.Empty;
            if (_tourTarget is Control ctrl)
            {
                rect = ctrl.RectangleToScreen(ctrl.ClientRectangle);
            }
            else if (_tourTarget is ToolStripItem item)
            {
                rect = item.Owner.RectangleToScreen(item.Bounds);
            }
            
            if (rect.IsEmpty) return Rectangle.Empty;

            // Step-specific overrides
            if (_tourStep == 1 && _tourTarget == _tabControl)
            {
                // Highlight only the top part (tabs)
                rect = new Rectangle(rect.X, rect.Y, rect.Width, 35);
            }
            else if (_tourStep == 2 && _btnRunAll != null && _btnSuspendAll != null)
            {
                // Union of Run All, Reload All, and Suspend All
                var r1 = _btnRunAll.Owner.RectangleToScreen(_btnRunAll.Bounds);
                var r2 = _btnSuspendAll.Owner.RectangleToScreen(_btnSuspendAll.Bounds);
                rect = Rectangle.Union(r1, r2);
            }

            rect.Inflate(5, 5);
            return rect;
        }


        private void ManageRegistryStartup()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    if (_data.StartWithWindows)
                        key.SetValue("AHKScriptManager", $"\"{Application.ExecutablePath}\" -minimized");
                    else
                        key.DeleteValue("AHKScriptManager", false);
                }
            }
            catch { }
        }

        private void RegisterSuspendHotkey()
        {
            UnregisterHotKey(Handle, HOTKEY_SUSPEND_ID);
            if (string.IsNullOrEmpty(_data.SuspendAllHotkey)) return;
            ParseHotkeyString(_data.SuspendAllHotkey, out int mod, out int vk);
            if (vk != 0) RegisterHotKey(Handle, HOTKEY_SUSPEND_ID, mod, vk);
        }

        private void ParseHotkeyString(string hk, out int mod, out int vk)
        {
            mod = 0; vk = 0;
            hk = hk.ToLower();
            if (hk.Contains("^")) mod |= 0x0002;
            if (hk.Contains("!")) mod |= 0x0001;
            if (hk.Contains("+")) mod |= 0x0004;
            if (hk.Contains("#")) mod |= 0x0008;

            var keyChar = hk.Replace("^", "").Replace("!", "").Replace("+", "").Replace("#", "").Trim();
            
            if (keyChar.StartsWith("f") && int.TryParse(keyChar.Substring(1), out int fNum) && fNum >= 1 && fNum <= 24)
            {
                vk = 0x6F + fNum;
            }
            else if (keyChar.Length == 1)
            {
                if (char.IsLetter(keyChar[0])) vk = char.ToUpper(keyChar[0]);
                else if (char.IsDigit(keyChar[0])) vk = (int)keyChar[0];
            }
            else
            {
                if (Enum.TryParse(keyChar, true, out Keys parsedKey))
                    vk = (int)parsedKey;
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_SUSPEND_ID) ToggleSuspendAll();
            base.WndProc(ref m);
        }

        private void ToggleSuspendAll()
        {
            _suspended = !_suspended;
            
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessId FROM Win32_Process WHERE Name LIKE '%AutoHotkey%'"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        try { Process.GetProcessById(Convert.ToInt32(obj["ProcessId"])).Kill(); } catch { }
                    }
                }
            }
            catch { }

            foreach (var s in _data.Scripts) s.IsRunning = false;
            RefreshStatusesInView();

            if (_notifyIcon != null)
                _notifyIcon.ShowBalloonTip(2000, "AHK Manager", "All AHK scripts killed/suspended.", ToolTipIcon.Info);
        }

        public void SaveData()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_dataPath));
                using (var sw = new StreamWriter(_dataPath)) new XmlSerializer(typeof(AppData)).Serialize(sw, _data);
            }
            catch { }
        }

        private void LoadData()
        {
            try
            {
                if (File.Exists(_dataPath))
                {
                    using (var sr = new StreamReader(_dataPath)) _data = (AppData)new XmlSerializer(typeof(AppData)).Deserialize(sr);
                    foreach (var s in _data.Scripts) s.ParseHotkeys();
                }
            }
            catch { _data = new AppData(); }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            RegisterSuspendHotkey();
            ManageRegistryStartup();
            
            // If started with windows, start minimized (by simulating form load minification or checking args if supported)
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (WindowState == FormWindowState.Minimized) Hide();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing && _data.MinimizeToTray)
            {
                // Intercept closing, just minimize to tray
                e.Cancel = true;
                WindowState = FormWindowState.Minimized;
                Hide();

                if (_data.ShowTrayHint)
                {
                    _data.ShowTrayHint = false;
                    SaveData();
                    Prompts.ShowTrayHintDialog(this.Icon);
                }
                return;
            }

            _statusTimer?.Stop();
            _notifyIcon?.Dispose();
            UnregisterHotKey(Handle, HOTKEY_SUSPEND_ID);
            
            try
            {
                foreach (var s in _data.Scripts)
                {
                    if (s.IsRunning) StopScript(s);
                }
            }
            catch { }
        }
    }

    // ─── Entry Point ─────────────────────────────────────────────────────────────
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var current = Process.GetCurrentProcess();
            var others = Process.GetProcessesByName(current.ProcessName)
                                .Where(p => p.Id != current.Id)
                                .ToList();

            if (others.Any())
            {
                if (MessageBox.Show("An instance of AHK Manager is already running, would you still like to open?", 
                                    "Already Running", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    foreach (var p in others)
                    {
                        try 
                        { 
                            p.Kill(); 
                            p.WaitForExit(3000); 
                        } catch { }
                    }
                }
                else
                {
                    return;
                }
            }

            Application.Run(new MainForm());
        }
    }
}
