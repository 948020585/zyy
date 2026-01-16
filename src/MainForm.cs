using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace CertPhotoSorter
{
    internal sealed class MainForm : Form
    {
        private readonly TextBox _excelPath = new TextBox();
        private readonly TextBox _photoRoot = new TextBox();
        private readonly TextBox _outputRoot = new TextBox();
        private readonly ComboBox _worksheet = new ComboBox();
        private readonly ComboBox _matchMode = new ComboBox();
        private readonly ToggleSwitch _dryRun = new ToggleSwitch();
        private readonly ToggleSwitch _updateExcel = new ToggleSwitch();
        private readonly Button _runButton = new Button();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly TextBox _log = new TextBox();

        public MainForm()
        {
            Text = Texts.AppTitle;
            Width = 1100;
            Height = 720;
            MinimumSize = new Size(980, 640);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            var sidebarBack = Color.FromArgb(243, 244, 246);
            var mainBack = Color.FromArgb(249, 250, 251);
            var cardBack = Color.White;
            var textColor = Color.FromArgb(17, 24, 39);
            var mutedText = Color.FromArgb(107, 114, 128);
            var primary = Color.FromArgb(59, 130, 246);

            BackColor = mainBack;

            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 2;
            root.RowCount = 1;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(root);

            var sidebar = new Panel();
            sidebar.Dock = DockStyle.Fill;
            sidebar.BackColor = sidebarBack;
            root.Controls.Add(sidebar, 0, 0);

            var main = new Panel();
            main.Dock = DockStyle.Fill;
            main.BackColor = mainBack;
            root.Controls.Add(main, 1, 0);

            var sidebarLayout = new TableLayoutPanel();
            sidebarLayout.Dock = DockStyle.Fill;
            sidebarLayout.ColumnCount = 1;
            sidebarLayout.RowCount = 5;
            sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
            sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            sidebar.Controls.Add(sidebarLayout);

            var brandPanel = new Panel();
            brandPanel.Dock = DockStyle.Fill;
            brandPanel.Padding = new Padding(16, 18, 16, 0);
            brandPanel.BackColor = sidebarBack;

            var brand = new Label();
            brand.Text = Texts.AppTitle;
            brand.Dock = DockStyle.Fill;
            brand.Font = new Font(Font.FontFamily, 12F, FontStyle.Bold);
            brand.ForeColor = textColor;
            brandPanel.Controls.Add(brand);
            sidebarLayout.Controls.Add(brandPanel, 0, 0);

            var navRun = MakeNavButton("\u8FD0\u884C", true, sidebarBack, textColor, (s, e) => { });
            sidebarLayout.Controls.Add(navRun, 0, 1);

            var navAbout = MakeNavButton("\u5173\u4E8E", false, sidebarBack, textColor, (s, e) => ShowAbout());
            sidebarLayout.Controls.Add(navAbout, 0, 2);

            var sidebarHint = new Label();
            sidebarHint.Text = "\u4EC5\u672C\u5730\u5904\u7406\uFF0C\u65E0\u9700\u8054\u7F51";
            sidebarHint.Dock = DockStyle.Fill;
            sidebarHint.Padding = new Padding(16, 0, 16, 0);
            sidebarHint.TextAlign = ContentAlignment.MiddleLeft;
            sidebarHint.ForeColor = mutedText;
            sidebarLayout.Controls.Add(sidebarHint, 0, 4);

            var mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.ColumnCount = 1;
            mainLayout.RowCount = 2;
            mainLayout.Padding = new Padding(24, 18, 24, 18);
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            main.Controls.Add(mainLayout);

            var headerTitle = new Label();
            headerTitle.Text = "\u8FD0\u884C\u914D\u7F6E";
            headerTitle.Dock = DockStyle.Fill;
            headerTitle.Font = new Font(Font.FontFamily, 14F, FontStyle.Bold);
            headerTitle.ForeColor = textColor;
            headerTitle.TextAlign = ContentAlignment.MiddleLeft;
            mainLayout.Controls.Add(headerTitle, 0, 0);

            var contentLayout = new TableLayoutPanel();
            contentLayout.Dock = DockStyle.Fill;
            contentLayout.ColumnCount = 1;
            contentLayout.RowCount = 2;
            contentLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 460));
            contentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            contentLayout.BackColor = mainBack;
            mainLayout.Controls.Add(contentLayout, 0, 1);

            var settingsCard = new Panel();
            settingsCard.Dock = DockStyle.Fill;
            settingsCard.BackColor = cardBack;
            settingsCard.Padding = new Padding(22);
            settingsCard.Margin = new Padding(0, 0, 0, 18);
            contentLayout.Controls.Add(settingsCard, 0, 0);

            var settingsCardLayout = new TableLayoutPanel();
            settingsCardLayout.Dock = DockStyle.Fill;
            settingsCardLayout.ColumnCount = 1;
            settingsCardLayout.RowCount = 2;
            settingsCardLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            settingsCardLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            settingsCard.Controls.Add(settingsCardLayout);

            var settingsTitle = new Label();
            settingsTitle.Text = "\u57FA\u672C\u53C2\u6570";
            settingsTitle.Dock = DockStyle.Fill;
            settingsTitle.Font = new Font(Font.FontFamily, 10F, FontStyle.Bold);
            settingsTitle.ForeColor = textColor;
            settingsTitle.TextAlign = ContentAlignment.MiddleLeft;
            settingsCardLayout.Controls.Add(settingsTitle, 0, 0);

            var form = new TableLayoutPanel();
            form.Dock = DockStyle.Fill;
            form.ColumnCount = 3;
            form.RowCount = 8;
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            for (int i = 0; i < form.RowCount; i++)
            {
                form.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            }
            settingsCardLayout.Controls.Add(form, 0, 1);

            _excelPath.Dock = DockStyle.Fill;
            _excelPath.BorderStyle = BorderStyle.FixedSingle;

            _photoRoot.Dock = DockStyle.Fill;
            _photoRoot.BorderStyle = BorderStyle.FixedSingle;

            _outputRoot.Dock = DockStyle.Fill;
            _outputRoot.BorderStyle = BorderStyle.FixedSingle;

            _worksheet.Dock = DockStyle.Fill;
            _worksheet.DropDownStyle = ComboBoxStyle.DropDown;
            _worksheet.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            _worksheet.AutoCompleteSource = AutoCompleteSource.ListItems;
            ResetWorksheetCombo();

            _matchMode.Dock = DockStyle.Fill;
            _matchMode.DropDownStyle = ComboBoxStyle.DropDownList;
            _matchMode.Items.Add(Texts.UiMatchModeIdOnly);
            _matchMode.Items.Add(Texts.UiMatchModeNameAndId);
            _matchMode.SelectedIndex = 0;

            _runButton.Text = Texts.UiRun;
            _runButton.Dock = DockStyle.Fill;
            _runButton.Height = 36;
            _runButton.FlatStyle = FlatStyle.Flat;
            _runButton.FlatAppearance.BorderSize = 0;
            _runButton.BackColor = primary;
            _runButton.ForeColor = Color.White;
            _runButton.UseVisualStyleBackColor = false;
            _runButton.Click += RunButtonOnClick;

            _progress.Dock = DockStyle.Fill;
            _progress.Minimum = 0;
            _progress.Maximum = 100;
            _progress.Margin = new Padding(0, 10, 10, 0);
            _progress.Style = ProgressBarStyle.Continuous;

            _log.Dock = DockStyle.Fill;
            _log.Multiline = true;
            _log.ScrollBars = ScrollBars.Both;
            _log.ReadOnly = true;
            _log.WordWrap = false;
            _log.BorderStyle = BorderStyle.None;
            _log.BackColor = Color.White;

            form.Controls.Add(MakeFieldLabel(Texts.UiLabelExcel, mutedText), 0, 0);
            form.Controls.Add(_excelPath, 1, 0);
            form.Controls.Add(MakeButton(Texts.UiChooseFile, ChooseExcel), 2, 0);

            form.Controls.Add(MakeFieldLabel(Texts.UiLabelPhotos, mutedText), 0, 1);
            form.Controls.Add(_photoRoot, 1, 1);
            form.Controls.Add(MakeButton(Texts.UiChooseFolder, ChoosePhotoRoot), 2, 1);

            form.Controls.Add(MakeFieldLabel(Texts.UiLabelOutput, mutedText), 0, 2);
            form.Controls.Add(_outputRoot, 1, 2);
            form.Controls.Add(MakeButton(Texts.UiChooseFolder, ChooseOutputRoot), 2, 2);

            form.Controls.Add(MakeFieldLabel(Texts.UiLabelWorksheet, mutedText), 0, 3);
            form.Controls.Add(_worksheet, 1, 3);
            form.Controls.Add(MakeButton(Texts.UiLoadSheets, LoadWorksheets), 2, 3);

            form.Controls.Add(MakeFieldLabel(Texts.UiLabelMatchMode, mutedText), 0, 4);
            form.Controls.Add(_matchMode, 1, 4);
            form.Controls.Add(new Panel { Dock = DockStyle.Fill }, 2, 4);

            var dryRunLabel = MakeFieldLabel(Texts.UiDryRun, mutedText);
            form.Controls.Add(dryRunLabel, 0, 5);
            form.SetColumnSpan(dryRunLabel, 2);
            _dryRun.Anchor = AnchorStyles.Right;
            _dryRun.Margin = new Padding(0, 10, 0, 0);
            form.Controls.Add(_dryRun, 2, 5);

            var updateExcelLabel = MakeFieldLabel(Texts.UiUpdateExcel, mutedText);
            form.Controls.Add(updateExcelLabel, 0, 6);
            form.SetColumnSpan(updateExcelLabel, 2);
            _updateExcel.Anchor = AnchorStyles.Right;
            _updateExcel.Margin = new Padding(0, 10, 0, 0);
            form.Controls.Add(_updateExcel, 2, 6);

            form.Controls.Add(_progress, 0, 7);
            form.SetColumnSpan(_progress, 2);
            form.Controls.Add(_runButton, 2, 7);

            var logCard = new Panel();
            logCard.Dock = DockStyle.Fill;
            logCard.BackColor = cardBack;
            logCard.Padding = new Padding(22);
            contentLayout.Controls.Add(logCard, 0, 1);

            var logCardLayout = new TableLayoutPanel();
            logCardLayout.Dock = DockStyle.Fill;
            logCardLayout.ColumnCount = 1;
            logCardLayout.RowCount = 2;
            logCardLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            logCardLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            logCard.Controls.Add(logCardLayout);

            var logTitle = new Label();
            logTitle.Text = "\u8FD0\u884C\u65E5\u5FD7";
            logTitle.Dock = DockStyle.Fill;
            logTitle.Font = new Font(Font.FontFamily, 10F, FontStyle.Bold);
            logTitle.ForeColor = textColor;
            logTitle.TextAlign = ContentAlignment.MiddleLeft;
            logCardLayout.Controls.Add(logTitle, 0, 0);

            var logBox = new Panel();
            logBox.Dock = DockStyle.Fill;
            logBox.BackColor = Color.FromArgb(229, 231, 235);
            logBox.Padding = new Padding(1);

            var logInner = new Panel();
            logInner.Dock = DockStyle.Fill;
            logInner.BackColor = Color.White;
            logInner.Padding = new Padding(10, 8, 10, 8);
            logInner.Controls.Add(_log);

            logBox.Controls.Add(logInner);
            logCardLayout.Controls.Add(logBox, 0, 1);
        }

        private static Button MakeButton(string text, EventHandler onClick)
        {
            var button = new Button();
            button.Text = text;
            button.Dock = DockStyle.Fill;
            button.Height = 32;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Color.FromArgb(75, 85, 99);
            button.ForeColor = Color.White;
            button.UseVisualStyleBackColor = false;
            button.Click += onClick;
            return button;
        }

        private static Button MakeNavButton(string text, bool selected, Color sidebarBack, Color textColor, EventHandler onClick)
        {
            var button = new Button();
            button.Text = text;
            button.Dock = DockStyle.Fill;
            button.Height = 40;
            button.TextAlign = ContentAlignment.MiddleLeft;
            button.Padding = new Padding(16, 0, 0, 0);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(229, 231, 235);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(209, 213, 219);
            button.BackColor = selected ? Color.FromArgb(229, 231, 235) : sidebarBack;
            button.ForeColor = textColor;
            button.UseVisualStyleBackColor = false;
            button.Click += onClick;
            return button;
        }

        private static Label MakeFieldLabel(string text, Color foreColor)
        {
            var label = new Label();
            label.Text = text;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.ForeColor = foreColor;
            label.AutoEllipsis = true;
            return label;
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                this,
                Texts.AppTitle + Environment.NewLine +
                Environment.NewLine +
                "- Excel \u8BFB\u53D6\u652F\u6301 OLEDB/NPOI\uFF08\u65E0\u9700\u5B89\u88C5 Office/WPS \u63D2\u4EF6\uFF09" + Environment.NewLine +
                "- \u9ED8\u8BA4\u4E0D\u56DE\u5199 Excel\uFF0C\u52FE\u9009\u201C\u56DE\u5199Excel\uFF08\u53EF\u9009\uFF09\u201D\u624D\u4F1A\u5C1D\u8BD5\u56DE\u5199" + Environment.NewLine,
                "\u5173\u4E8E",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void ChooseExcel(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = Texts.UiChooseFile;
                dlg.Filter = "Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm|All (*.*)|*.*";
                if (dlg.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _excelPath.Text = dlg.FileName;
                if (string.IsNullOrWhiteSpace(_outputRoot.Text))
                {
                    var dir = Path.GetDirectoryName(dlg.FileName);
                    if (!string.IsNullOrWhiteSpace(dir))
                    {
                        _outputRoot.Text = Path.Combine(dir, Texts.DefaultOutputFolder);
                    }
                }

                TryLoadWorksheets();
            }
        }

        private void ChoosePhotoRoot(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = Texts.UiLabelPhotos;
                if (dlg.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _photoRoot.Text = dlg.SelectedPath;
            }
        }

        private void ChooseOutputRoot(object sender, EventArgs e)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                dlg.Description = Texts.UiLabelOutput;
                if (dlg.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _outputRoot.Text = dlg.SelectedPath;
            }
        }

        private void RunButtonOnClick(object sender, EventArgs e)
        {
            var excel = (_excelPath.Text ?? string.Empty).Trim();
            var photos = (_photoRoot.Text ?? string.Empty).Trim();
            var output = (_outputRoot.Text ?? string.Empty).Trim();
            var worksheet = (_worksheet.Text ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(excel) || !File.Exists(excel))
            {
                MessageBox.Show(this, Texts.MsgNeedExcel, Texts.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(photos) || !Directory.Exists(photos))
            {
                MessageBox.Show(this, Texts.MsgNeedPhotoRoot, Texts.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(output))
            {
                output = Path.Combine(Path.GetDirectoryName(excel) ?? Environment.CurrentDirectory, Texts.DefaultOutputFolder);
                _outputRoot.Text = output;
            }

            var settings = new RunSettings
            {
                ExcelPath = excel,
                PhotoRoot = photos,
                OutputRoot = output,
                Worksheet = string.IsNullOrWhiteSpace(worksheet) || string.Equals(worksheet, Texts.UiAutoDetect, StringComparison.Ordinal) ? null : worksheet,
                DryRun = _dryRun.Checked,
                UpdateExcel = _updateExcel.Checked,
                MatchMode = _matchMode.SelectedIndex == 1 ? MatchMode.NameAndId : MatchMode.IdOnly
            };

            _runButton.Enabled = false;
            _progress.Value = 0;
            _log.Clear();
            AppendLog("Start...");

            var worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += (o, args2) =>
            {
                var result = Processor.Execute(
                    settings,
                    msg => worker.ReportProgress(-1, msg),
                    percent => worker.ReportProgress(percent, null));
                args2.Result = result;
            };

            worker.ProgressChanged += (o, args2) =>
            {
                if (args2.ProgressPercentage >= 0 && args2.ProgressPercentage <= 100)
                {
                    _progress.Value = args2.ProgressPercentage;
                }

                var msg = args2.UserState as string;
                if (!string.IsNullOrWhiteSpace(msg))
                {
                    AppendLog(msg);
                }
            };

            worker.RunWorkerCompleted += (o, args2) =>
            {
                _runButton.Enabled = true;

                if (args2.Error != null)
                {
                    AppendLog("Failed: " + args2.Error.Message);
                    MessageBox.Show(this, args2.Error.ToString(), Texts.MsgFailed, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var result = args2.Result as RunResult;
                if (result == null)
                {
                    MessageBox.Show(this, Texts.MsgDone, Texts.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                AppendLog("Done.");
                AppendLog("Report: " + result.ReportPath);

                MessageBox.Show(
                    this,
                    Texts.MsgDone + Environment.NewLine + Environment.NewLine + result.ReportPath,
                    Texts.AppTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            };

            worker.RunWorkerAsync();
        }

        private void LoadWorksheets(object sender, EventArgs e)
        {
            TryLoadWorksheets();
        }

        private void TryLoadWorksheets()
        {
            var excel = (_excelPath.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(excel) || !File.Exists(excel))
            {
                ResetWorksheetCombo();
                return;
            }

            try
            {
                string provider;
                var sheets = ExcelReader.ListWorksheetNames(excel, out provider);
                PopulateWorksheetCombo(sheets);
                AppendLog("工作表已读取：" + sheets.Count + " (" + provider + ")");
            }
            catch (Exception ex)
            {
                ResetWorksheetCombo();
                AppendLog("读取工作表失败：" + ex.Message);
            }
        }

        private void ResetWorksheetCombo()
        {
            _worksheet.BeginUpdate();
            _worksheet.Items.Clear();
            _worksheet.Items.Add(Texts.UiAutoDetect);
            _worksheet.SelectedIndex = 0;
            _worksheet.EndUpdate();
        }

        private void PopulateWorksheetCombo(List<string> sheets)
        {
            _worksheet.BeginUpdate();
            _worksheet.Items.Clear();
            _worksheet.Items.Add(Texts.UiAutoDetect);

            if (sheets != null)
            {
                for (int i = 0; i < sheets.Count; i++)
                {
                    var sheet = sheets[i] ?? string.Empty;
                    if (sheet.EndsWith("$", StringComparison.Ordinal))
                    {
                        sheet = sheet.Substring(0, sheet.Length - 1);
                    }

                    if (sheet.Length == 0)
                    {
                        continue;
                    }

                    _worksheet.Items.Add(sheet);
                }
            }

            _worksheet.SelectedIndex = 0;
            _worksheet.EndUpdate();
        }

        private void AppendLog(string message)
        {
            _log.AppendText(DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine);
        }
    }
}
