/*
 * Copyright (c) 2009 Jim Radford http://www.jimradford.com
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions: 
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using log4net;
using SuperPutty.Data;
using SuperPutty.Gui;
using SuperPutty.Utils;

namespace SuperPutty
{
    public partial class DlgFindPutty : Form
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(DlgFindPutty));

        private string OrigSettingsFolder { get; }
        private string OrigDefaultLayoutName { get; set; }

        private BindingList<KeyboardShortcut> Shortcuts { get; }

        public DlgFindPutty()
        {
            InitializeComponent();

            string puttyExe = SuperPuTTY.Settings.PuttyExe;
            string pscpExe = SuperPuTTY.Settings.PscpExe;

            bool firstExecution = String.IsNullOrEmpty(puttyExe);
            textBoxFilezillaLocation.Text = getPathExe(@"\FileZilla FTP Client\filezilla.exe", SuperPuTTY.Settings.FileZillaExe, firstExecution);
            textBoxWinSCPLocation.Text = getPathExe(@"\WinSCP\WinSCP.exe", SuperPuTTY.Settings.WinSCPExe, firstExecution);

            // check for location of putty/pscp
            if (!String.IsNullOrEmpty(puttyExe) && File.Exists(puttyExe))
            {
                textBoxPuttyLocation.Text = puttyExe;
                if (!String.IsNullOrEmpty(pscpExe) && File.Exists(pscpExe))
                {
                    textBoxPscpLocation.Text = pscpExe;
                }
            }
            else if(!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("ProgramFiles(x86)")))
            {
                if (File.Exists(Environment.GetEnvironmentVariable("ProgramFiles(x86)") + @"\PuTTY\putty.exe"))
                {
                    textBoxPuttyLocation.Text = Environment.GetEnvironmentVariable("ProgramFiles(x86)") + @"\PuTTY\putty.exe";
                    openFileDialog1.InitialDirectory = Environment.GetEnvironmentVariable("ProgramFiles(x86)");
                }

                if (File.Exists(Environment.GetEnvironmentVariable("ProgramFiles(x86)") + @"\PuTTY\pscp.exe"))
                {

                    textBoxPscpLocation.Text = Environment.GetEnvironmentVariable("ProgramFiles(x86)") + @"\PuTTY\pscp.exe";
                }
            }
            else if (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("ProgramFiles")))
            {
                if (File.Exists(Environment.GetEnvironmentVariable("ProgramFiles") + @"\PuTTY\putty.exe"))
                {
                    textBoxPuttyLocation.Text = Environment.GetEnvironmentVariable("ProgramFiles") + @"\PuTTY\putty.exe";
                    openFileDialog1.InitialDirectory = Environment.GetEnvironmentVariable("ProgramFiles");
                }

                if (File.Exists(Environment.GetEnvironmentVariable("ProgramFiles") + @"\PuTTY\pscp.exe"))
                {
                    textBoxPscpLocation.Text = Environment.GetEnvironmentVariable("ProgramFiles") + @"\PuTTY\pscp.exe";
                }
            }            
            else
            {
                openFileDialog1.InitialDirectory = Application.StartupPath;
            }

            if (String.IsNullOrEmpty(SuperPuTTY.Settings.MinttyExe))
            {
                if (File.Exists(@"C:\cygwin\bin\mintty.exe"))
                {
                    textBoxMinttyLocation.Text = @"C:\cygwin\bin\mintty.exe";
                }
            }
            else
            {
                textBoxMinttyLocation.Text = SuperPuTTY.Settings.MinttyExe;
            }
            
            // super putty settings (sessions and layouts)
            if (String.IsNullOrEmpty(SuperPuTTY.Settings.SettingsFolder))
            {
                // Set a default
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SuperPuTTY");
                if (!Directory.Exists(dir))
                {
                    Log.InfoFormat("Creating default settings dir: {0}", dir);
                    Directory.CreateDirectory(dir);
                }
                textBoxSettingsFolder.Text = dir;
            }
            else
            {
                textBoxSettingsFolder.Text = SuperPuTTY.Settings.SettingsFolder;
            }
            OrigSettingsFolder = SuperPuTTY.Settings.SettingsFolder;

            // tab text
            foreach(String s in Enum.GetNames(typeof(frmSuperPutty.TabTextBehavior)))
            {
                comboBoxTabText.Items.Add(s);
            }
            comboBoxTabText.SelectedItem = SuperPuTTY.Settings.TabTextBehavior;

            // tab switcher
            ITabSwitchStrategy selectedItem = null;
            foreach (ITabSwitchStrategy strat in TabSwitcher.Strategies)
            {
                comboBoxTabSwitching.Items.Add(strat);
                if (strat.GetType().FullName == SuperPuTTY.Settings.TabSwitcher)
                {
                    selectedItem = strat;
                }
            }
            comboBoxTabSwitching.SelectedItem = selectedItem ?? TabSwitcher.Strategies[0];

            // activator types
            comboBoxActivatorType.Items.Add(typeof(KeyEventWindowActivator).FullName);
            comboBoxActivatorType.Items.Add(typeof(CombinedWindowActivator).FullName);
            comboBoxActivatorType.Items.Add(typeof(SetFGWindowActivator).FullName);
            comboBoxActivatorType.Items.Add(typeof(RestoreWindowActivator).FullName);
            comboBoxActivatorType.Items.Add(typeof(SetFGAttachThreadWindowActivator).FullName);
            comboBoxActivatorType.SelectedItem = SuperPuTTY.Settings.WindowActivator;

            // search types
            foreach (string name in Enum.GetNames(typeof(SessionTreeview.SearchMode)))
            {
                comboSearchMode.Items.Add(name);
            }
            comboSearchMode.SelectedItem = SuperPuTTY.Settings.SessionsSearchMode;

            // default layouts
            InitLayouts();

            checkSingleInstanceMode.Checked = SuperPuTTY.Settings.SingleInstanceMode;
            checkConstrainPuttyDocking.Checked = SuperPuTTY.Settings.RestrictContentToDocumentTabs;
            checkRestoreWindow.Checked = SuperPuTTY.Settings.RestoreWindowLocation;
            checkExitConfirmation.Checked = SuperPuTTY.Settings.ExitConfirmation;
            checkExpandTree.Checked = SuperPuTTY.Settings.ExpandSessionsTreeOnStartup;
            checkMinimizeToTray.Checked = SuperPuTTY.Settings.MinimizeToTray;
            checkSessionsTreeShowLines.Checked = SuperPuTTY.Settings.SessionsTreeShowLines;
            checkConfirmTabClose.Checked = SuperPuTTY.Settings.MultipleTabCloseConfirmation;
            checkEnableControlTabSwitching.Checked = SuperPuTTY.Settings.EnableControlTabSwitching;
            checkEnableKeyboardShortcuts.Checked = SuperPuTTY.Settings.EnableKeyboadShortcuts;
            btnFont.Font = SuperPuTTY.Settings.SessionsTreeFont;
            btnFont.Text = ToShortString(SuperPuTTY.Settings.SessionsTreeFont);
            numericUpDownOpacity.Value = (decimal) SuperPuTTY.Settings.Opacity * 100;
            checkQuickSelectorCaseSensitiveSearch.Checked = SuperPuTTY.Settings.QuickSelectorCaseSensitiveSearch;
            checkShowDocumentIcons.Checked = SuperPuTTY.Settings.ShowDocumentIcons;
            checkRestrictFloatingWindows.Checked = SuperPuTTY.Settings.DockingRestrictFloatingWindows;
            checkSessionsShowSearch.Checked = SuperPuTTY.Settings.SessionsShowSearch;
            checkPuttyEnableNewSessionMenu.Checked = SuperPuTTY.Settings.PuttyPanelShowNewSessionMenu;
            checkBoxCheckForUpdates.Checked = SuperPuTTY.Settings.AutoUpdateCheck;
            textBoxHomeDirPrefix.Text = SuperPuTTY.Settings.PscpHomePrefix;
            textBoxRootDirPrefix.Text = SuperPuTTY.Settings.PscpRootHomePrefix;
            checkSessionTreeFoldersFirst.Checked = SuperPuTTY.Settings.SessiontreeShowFoldersFirst;
            checkBoxPersistTsHistory.Checked = SuperPuTTY.Settings.PersistCommandBarHistory;
            numericUpDown1.Value = SuperPuTTY.Settings.SaveCommandHistoryDays;
            checkBoxAllowPuttyPWArg.Checked = SuperPuTTY.Settings.AllowPlainTextPuttyPasswordArg;
            textBoxPuttyDefaultParameters.Text = SuperPuTTY.Settings.PuttyDefaultParameters;

            if (SuperPuTTY.IsFirstRun)
            {
                ShowIcon = true;
                ShowInTaskbar = true;
            }

            // shortcuts
            Shortcuts = new BindingList<KeyboardShortcut>();
            foreach (KeyboardShortcut ks in SuperPuTTY.Settings.LoadShortcuts())
            {
                Shortcuts.Add(ks);
            }
            dataGridViewShortcuts.DataSource = Shortcuts;
        }


        /// <summary>
        /// return the path of the exe. 
        /// return settingValue if it is a valid path, or if searchPath is false, else search and return the default location of pathInProgramFile.
        /// </summary>
        /// <param name="pathInProgramFile">relative path of file (in ProgramFiles or ProgramFiles(x86))</param>
        /// <param name="settingValue">path stored in settings </param>
        /// <param name="searchPath">boolean </param>
        /// <returns>The path of the exe</returns>
        private String getPathExe(String pathInProgramFile, String settingValue, Boolean searchPath)
        {
            if ((!String.IsNullOrEmpty(settingValue) && File.Exists(settingValue)) || !searchPath)
            {
                return settingValue;
            }

            if (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("ProgramFiles(x86)")))
            {
                if (File.Exists(Environment.GetEnvironmentVariable("ProgramFiles(x86)") + pathInProgramFile))
                {
                    return Environment.GetEnvironmentVariable("ProgramFiles(x86)") + pathInProgramFile;
                }
            }

            if (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("ProgramFiles")))
            {
                if (File.Exists(Environment.GetEnvironmentVariable("ProgramFiles") + pathInProgramFile))
                {
                    return Environment.GetEnvironmentVariable("ProgramFiles") + pathInProgramFile;
                }
            }

            return "";
        }


        private void InitLayouts()
        {
            String defaultLayout;
            List<String> layouts = new List<string>();
            if (SuperPuTTY.IsFirstRun)
            {
                layouts.Add(String.Empty);
                // HACK: first time so layouts directory not set yet so layouts don't exist...
                //       preload <AutoRestore> so we can set it as default
                layouts.Add(LayoutData.AutoRestore);

                defaultLayout = LayoutData.AutoRestore;
            }
            else
            {
                layouts.Add(String.Empty);
                // auto restore is in the layouts collection already
                layouts.AddRange(SuperPuTTY.Layouts.Select(layout => layout.Name));

                defaultLayout = SuperPuTTY.Settings.DefaultLayoutName;
            }
            comboBoxLayouts.DataSource = layouts;
            comboBoxLayouts.SelectedItem = defaultLayout;
            OrigDefaultLayoutName = defaultLayout;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            BeginInvoke(new MethodInvoker(delegate { textBoxPuttyLocation.Focus(); }));
        }
       
        private void buttonOk_Click(object sender, EventArgs e)
        {
            List<String> errors = new List<string>();

            if (String.IsNullOrEmpty(textBoxFilezillaLocation.Text) || File.Exists(textBoxFilezillaLocation.Text))
            {
                SuperPuTTY.Settings.FileZillaExe = textBoxFilezillaLocation.Text;
            }

            if (String.IsNullOrEmpty(textBoxWinSCPLocation.Text) || File.Exists(textBoxWinSCPLocation.Text))
            {
                SuperPuTTY.Settings.WinSCPExe = textBoxWinSCPLocation.Text;
            }

            if (String.IsNullOrEmpty(textBoxPscpLocation.Text) || File.Exists(textBoxPscpLocation.Text))
            {
                SuperPuTTY.Settings.PscpExe = textBoxPscpLocation.Text;
            }

            string settingsDir = textBoxSettingsFolder.Text;
            if (String.IsNullOrEmpty(settingsDir) || !Directory.Exists(settingsDir))
            {
                errors.Add("Settings Folder must be set to valid directory");
            }
            else
            {
                SuperPuTTY.Settings.SettingsFolder = settingsDir;
            }

            if (comboBoxLayouts.SelectedValue != null)
            {
                SuperPuTTY.Settings.DefaultLayoutName = (string) comboBoxLayouts.SelectedValue;
            }

            if (!String.IsNullOrEmpty(textBoxPuttyLocation.Text) && File.Exists(textBoxPuttyLocation.Text))
            {
                SuperPuTTY.Settings.PuttyExe = textBoxPuttyLocation.Text;
            }
            else
            {
                errors.Insert(0, "PuTTY is required to properly use this application.");
            }

            string mintty = textBoxMinttyLocation.Text;
            if (!string.IsNullOrEmpty(mintty) && File.Exists(mintty))
            {
                SuperPuTTY.Settings.MinttyExe = mintty;
            }

            if (errors.Count == 0)
            {
                SuperPuTTY.Settings.SingleInstanceMode = checkSingleInstanceMode.Checked;
                SuperPuTTY.Settings.RestrictContentToDocumentTabs = checkConstrainPuttyDocking.Checked;
                SuperPuTTY.Settings.MultipleTabCloseConfirmation= checkConfirmTabClose.Checked;
                SuperPuTTY.Settings.RestoreWindowLocation = checkRestoreWindow.Checked;
                SuperPuTTY.Settings.ExitConfirmation = checkExitConfirmation.Checked;
                SuperPuTTY.Settings.ExpandSessionsTreeOnStartup = checkExpandTree.Checked;
                SuperPuTTY.Settings.EnableControlTabSwitching = checkEnableControlTabSwitching.Checked;
                SuperPuTTY.Settings.EnableKeyboadShortcuts = checkEnableKeyboardShortcuts.Checked;
                SuperPuTTY.Settings.MinimizeToTray = checkMinimizeToTray.Checked;
                SuperPuTTY.Settings.TabTextBehavior = (string) comboBoxTabText.SelectedItem;
                SuperPuTTY.Settings.TabSwitcher = comboBoxTabSwitching.SelectedItem.GetType().FullName;
                SuperPuTTY.Settings.SessionsTreeShowLines = checkSessionsTreeShowLines.Checked;
                SuperPuTTY.Settings.SessionsTreeFont = btnFont.Font;
                SuperPuTTY.Settings.WindowActivator = (string) comboBoxActivatorType.SelectedItem;
                SuperPuTTY.Settings.Opacity = (double) numericUpDownOpacity.Value / 100.0;
                SuperPuTTY.Settings.SessionsSearchMode = (string) comboSearchMode.SelectedItem;
                SuperPuTTY.Settings.QuickSelectorCaseSensitiveSearch = checkQuickSelectorCaseSensitiveSearch.Checked;
                SuperPuTTY.Settings.ShowDocumentIcons = checkShowDocumentIcons.Checked;
                SuperPuTTY.Settings.DockingRestrictFloatingWindows = checkRestrictFloatingWindows.Checked;
                SuperPuTTY.Settings.SessionsShowSearch = checkSessionsShowSearch.Checked;
                SuperPuTTY.Settings.PuttyPanelShowNewSessionMenu = checkPuttyEnableNewSessionMenu.Checked;
                SuperPuTTY.Settings.AutoUpdateCheck = checkBoxCheckForUpdates.Checked;
                SuperPuTTY.Settings.PscpHomePrefix = textBoxHomeDirPrefix.Text;
                SuperPuTTY.Settings.PscpRootHomePrefix = textBoxRootDirPrefix.Text;
                SuperPuTTY.Settings.SessiontreeShowFoldersFirst = checkSessionTreeFoldersFirst.Checked;
                SuperPuTTY.Settings.PersistCommandBarHistory = checkBoxPersistTsHistory.Checked;
                SuperPuTTY.Settings.SaveCommandHistoryDays = (int)numericUpDown1.Value;
                SuperPuTTY.Settings.AllowPlainTextPuttyPasswordArg = checkBoxAllowPuttyPWArg.Checked;
                SuperPuTTY.Settings.PuttyDefaultParameters = textBoxPuttyDefaultParameters.Text;

                // save shortcuts
                KeyboardShortcut[] shortcuts = new KeyboardShortcut[Shortcuts.Count];
                Shortcuts.CopyTo(shortcuts, 0);
                SuperPuTTY.Settings.UpdateFromShortcuts(shortcuts);

                SuperPuTTY.Settings.Save();

                // @TODO - move this to a better place...maybe event handler after opening
                if (OrigSettingsFolder != SuperPuTTY.Settings.SettingsFolder)
                {
                    SuperPuTTY.LoadLayouts();
                    SuperPuTTY.LoadSessions();
                }
                else if (OrigDefaultLayoutName != SuperPuTTY.Settings.DefaultLayoutName)
                {
                    SuperPuTTY.LoadLayouts();
                }

                DialogResult = DialogResult.OK;
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                foreach (string s in errors)
                {
                    sb.Append(s).AppendLine().AppendLine();
                }
                if (MessageBox.Show(sb.ToString(), @"Errors", MessageBoxButtons.RetryCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
                {
                    DialogResult = DialogResult.Cancel;
                }
            }
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = @"PuTTY|putty.exe|KiTTY|kitty*.exe";
            openFileDialog1.FileName = "putty.exe";
            if (File.Exists(textBoxPuttyLocation.Text))
            {
                openFileDialog1.FileName = Path.GetFileName(textBoxPuttyLocation.Text);
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(textBoxPuttyLocation.Text);
                openFileDialog1.FilterIndex = openFileDialog1.FileName != null && openFileDialog1.FileName.ToLower().StartsWith("putty") ? 1 : 2;
            }
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                if (!String.IsNullOrEmpty(openFileDialog1.FileName))
                    textBoxPuttyLocation.Text = openFileDialog1.FileName;
            }            
        }

        private void buttonBrowsePscp_Click(object sender, EventArgs e)
        {
            DialogBrowseExe("PScp|pscp.exe", "pscp.exe", textBoxPscpLocation);
        }

        private void btnBrowseMintty_Click(object sender, EventArgs e)
        {
            DialogBrowseExe("MinTTY|mintty.exe", "mintty.exe", textBoxMinttyLocation);
        }

        private void buttonBrowseFilezilla_Click(object sender, EventArgs e)
        {
            DialogBrowseExe("filezilla|filezilla.exe", "filezilla.exe", textBoxFilezillaLocation);
        }

        private void buttonBrowseWinSCP_Click(object sender, EventArgs e)
        {
            DialogBrowseExe("WinSCP|WinSCP.exe", "WinSCP.exe", textBoxWinSCPLocation);
        }

        private void DialogBrowseExe(String filter, string filename, TextBox textbox)
        {
            openFileDialog1.Filter = filter;
            openFileDialog1.FileName = filename;

            if (File.Exists(textbox.Text))
            {
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(textbox.Text);
            }
            if (openFileDialog1.ShowDialog(this) == DialogResult.OK)
            {
                if (!String.IsNullOrEmpty(openFileDialog1.FileName))
                    textbox.Text = openFileDialog1.FileName;
            }

        }

        //Search automaticaly the path of FileZilla when doubleClick when it is empty
        private void textBoxFilezillaLocation_DoubleClick(object sender, EventArgs e)
        {
            textBoxFilezillaLocation.Text = getPathExe(@"\FileZilla FTP Client\filezilla.exe", SuperPuTTY.Settings.FileZillaExe, true);
        }

        //Search automaticaly the path of WinSCP when doubleClick when it is empty
        private void textBoxWinSCPLocation_DoubleClick(object sender, EventArgs e)
        {
            textBoxWinSCPLocation.Text = getPathExe(@"\WinSCP\WinSCP.exe", SuperPuTTY.Settings.WinSCPExe, true);
        }


        /// <summary>
        /// Check that putty can be found.  If not, prompt the user
        /// </summary>
        public static void PuttyCheck()
        {
            if (String.IsNullOrEmpty(SuperPuTTY.Settings.PuttyExe) || SuperPuTTY.IsFirstRun || !File.Exists(SuperPuTTY.Settings.PuttyExe))
            {
                // first time, try to import old putty settings from registry
                SuperPuTTY.Settings.ImportFromRegistry();
                DlgFindPutty dialog = new DlgFindPutty();
                if (dialog.ShowDialog(SuperPuTTY.MainForm) == DialogResult.Cancel)
                {
                    Environment.Exit(1);
                }
            }

            if (String.IsNullOrEmpty(SuperPuTTY.Settings.PuttyExe))
            {
                MessageBox.Show(@"Cannot find PuTTY installation. Please visit http://www.chiark.greenend.org.uk/~sgtatham/putty/download.html to download a copy",
                    @"PuTTY Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                Environment.Exit(1);
            }

            if (SuperPuTTY.IsFirstRun && SuperPuTTY.Sessions.Count == 0)
            {
                // first run, got nothing...try to import from registry
                SuperPuTTY.ImportSessionsFromSuperPutty1030();
            }
        }

        private void buttonBrowseLayoutsFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog(this) == DialogResult.OK) 
            {
                textBoxSettingsFolder.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnFont_Click(object sender, EventArgs e)
        {
            fontDialog.Font = btnFont.Font;
            if (fontDialog.ShowDialog(this) == DialogResult.OK)
            {
                btnFont.Font = fontDialog.Font;
                btnFont.Text = ToShortString(fontDialog.Font);
            }
        }

        private void dataGridViewShortcuts_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 || e.ColumnIndex == -1) { return; }

            Log.InfoFormat("Shortcuts grid click: row={0}, col={1}", e.RowIndex, e.ColumnIndex);
            DataGridViewColumn col = dataGridViewShortcuts.Columns[e.ColumnIndex];
            DataGridViewRow row = dataGridViewShortcuts.Rows[e.RowIndex];
            KeyboardShortcut ks = (KeyboardShortcut) row.DataBoundItem;

            if (col == colEdit)
            {
                KeyboardShortcutEditor editor = new KeyboardShortcutEditor
                {
                    StartPosition = FormStartPosition.CenterParent
                };
                if (DialogResult.OK == editor.ShowDialog(this, ks))
                {
                    Shortcuts.ResetItem(Shortcuts.IndexOf(ks));
                    Log.InfoFormat("Edited shortcut: {0}", ks);
                }
            }
            else if (col == colClear)
            {
                ks.Clear();
                Shortcuts.ResetItem(Shortcuts.IndexOf(ks));
                Log.InfoFormat("Cleared shortcut: {0}", ks);
            }
        }

        static string ToShortString(Font font)
        {
            return $"{font.FontFamily.Name}, {font.Size} pt, {font.Style}";
        }       
    }

}
