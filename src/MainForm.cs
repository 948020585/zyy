using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private readonly CheckBox _dryRun = new CheckBox();
        private readonly Button _runButton = new Button();
        private readonly ProgressBar _progress = new ProgressBar();
        private readonly TextBox _log = new TextBox();

        public MainForm()
        {
            Text = Texts.AppTitle;
            Width = 900;
            Height = 650;

            var layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 3;
            layout.RowCount = 8;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));

            _excelPath.Dock = DockStyle.Fill;
            _photoRoot.Dock = DockStyle.Fill;
            _outputRoot.Dock = DockStyle.Fill;
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

            _dryRun.Text = Texts.UiDryRun;
            _dryRun.AutoSize = true;

            _runButton.Text = Texts.UiRun;
            _runButton.Width = 120;
            _runButton.Click += RunButtonOnClick;

            _progress.Dock = DockStyle.Fill;
            _progress.Minimum = 0;
            _progress.Maximum = 100;

            _log.Dock = DockStyle.Fill;
            _log.Multiline = true;
            _log.ScrollBars = ScrollBars.Both;
            _log.ReadOnly = true;
            _log.WordWrap = false;

            layout.Controls.Add(new Label { Text = Texts.UiLabelExcel, AutoSize = true }, 0, 0);
            layout.Controls.Add(_excelPath, 1, 0);
            layout.Controls.Add(MakeButton(Texts.UiChooseFile, ChooseExcel), 2, 0);

            layout.Controls.Add(new Label { Text = Texts.UiLabelPhotos, AutoSize = true }, 0, 1);
            layout.Controls.Add(_photoRoot, 1, 1);
            layout.Controls.Add(MakeButton(Texts.UiChooseFolder, ChoosePhotoRoot), 2, 1);

            layout.Controls.Add(new Label { Text = Texts.UiLabelOutput, AutoSize = true }, 0, 2);
            layout.Controls.Add(_outputRoot, 1, 2);
            layout.Controls.Add(MakeButton(Texts.UiChooseFolder, ChooseOutputRoot), 2, 2);

            layout.Controls.Add(new Label { Text = Texts.UiLabelWorksheet, AutoSize = true }, 0, 3);
            layout.Controls.Add(_worksheet, 1, 3);
            layout.Controls.Add(MakeButton(Texts.UiLoadSheets, LoadWorksheets), 2, 3);

            layout.Controls.Add(new Label { Text = Texts.UiLabelMatchMode, AutoSize = true }, 0, 4);
            layout.Controls.Add(_matchMode, 1, 4);

            layout.Controls.Add(_dryRun, 1, 5);
            layout.Controls.Add(_runButton, 2, 5);

            layout.Controls.Add(_progress, 1, 6);
            layout.SetColumnSpan(_progress, 2);

            layout.Controls.Add(_log, 0, 7);
            layout.SetColumnSpan(_log, 3);
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            Controls.Add(layout);
        }

        private static Button MakeButton(string text, EventHandler onClick)
        {
            var button = new Button();
            button.Text = text;
            button.Width = 120;
            button.Click += onClick;
            return button;
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
