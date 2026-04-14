using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
#if USE_IMAGESHARP
using SixLabors.ImageSharp.Formats.Gif;
#endif

namespace GrandBattleSupport
{
    public partial class MainForm : Form
    {
        private const int Days = 6;
        private const int TeamCount = 4;
        // チーム色
        private readonly Color[] _teamColors = new[]
        {
            Color.FromArgb(220, 231, 76, 60),  // Team A (赤)
            Color.FromArgb(220, 52, 152, 219), // Team B (青)
            Color.FromArgb(220, 46, 204, 113), // Team C (緑)
            Color.FromArgb(220, 241, 196, 15), // Team D (黄)
        };

        private readonly string[] _teamNames = new[] { "A", "B", "C", "D" };
        private readonly Image?[] _teamIcons = new Image?[TeamCount];
        private static readonly string[] TeamIconFiles =
        {
            Path.Combine("image", "flag_a.png"),
            Path.Combine("image", "flag_b.png"),
            Path.Combine("image", "flag_c.png"),
            Path.Combine("image", "flag_d.png")
        };

        private Dictionary<int, Node> _nodes = new();
        private OwnerState _owner = new(Days, TeamCount);

        // UI
        private SplitContainer _split = null!;
        private PictureBox _mapBox = null!;
        private Panel _rightPanel = null!;
        private NumericUpDown _daySelector = null!;
        private GroupBox _teamGroup = null!;
        private RadioButton[] _teamRadios = null!;
        private TextBox[] _teamNameBoxes = null!; // ギルド名入力欄
        private DataGridView _grid = null!;
        private GroupBox _totalsGroup = null!;
        private Label[] _totalLabels = null!;
        private Label _status = null!;
        private Button _copyPrevButton = null!;
        private Button _saveButton = null!;
        private Button _saveAsButton = null!;
        private Button _exportGifButton = null!;

        // 選択状態
        private int SelectedDay => (int)_daySelector.Value; // 1..6
        private int SelectedTeam => Array.FindIndex(_teamRadios, r => r.Checked); // 0..3

        // クリック判定（円の半径）
        private const int HitRadius = 22;
        private const int DrawRadius = 14;
        private const int IconDrawSize = 30;

        // ファイル名
        private const string NodesJsonFile = "nodes.json";
        // 保存フォルダ
        private const string SaveFolder = "save";
        // map.jpg 取得のためのパス（実行ファイルからの相対パス）
        private static readonly string MapImageFile = Path.Combine("image", "map.jpg");

        public MainForm()
        {
            Text = "メメントモリ グラバトポイント計算";
            Width = 1400;
            Height = 900;
            StartPosition = FormStartPosition.CenterScreen;

            BuildUi();
            LoadDataOrFallback();
            LoadTeamIcons();
            LoadOwnersIfExists();
            InitializeGridRows();
            RefreshLeftGridForDay();
            RefreshTotals();
            _mapBox.Invalidate();
        }

        // 名前を付けて保存（保存一覧ダイアログ）
        private void SaveOwnersToFileAs()
        {
            Directory.CreateDirectory(SaveFolder);

            using var dlg = new Form { Text = "保存", Width = 700, Height = 420, StartPosition = FormStartPosition.CenterParent };

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8) };
            var nameBox = new TextBox { Width = 420, Location = new Point(8, 8) };
            var btnSave = new Button { Text = "保存", Location = new Point(440, 6), AutoSize = true };
            topPanel.Controls.Add(nameBox);
            topPanel.Controls.Add(btnSave);

            var sep = new Label { Dock = DockStyle.Top, Height = 2, BackColor = SystemColors.ControlDark }; 

            var grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "No", HeaderText = "No", Width = 50 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "名前", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Updated", HeaderText = "更新日時", Width = 160 });
            var loadCol = new DataGridViewButtonColumn { Name = "Load", HeaderText = "操作", Text = "読込", UseColumnTextForButtonValue = true, Width = 80 };
            var delCol = new DataGridViewButtonColumn { Name = "Delete", HeaderText = "", Text = "削除", UseColumnTextForButtonValue = true, Width = 80 };
            grid.Columns.Add(loadCol);
            grid.Columns.Add(delCol);

            void RefreshList()
            {
                grid.Rows.Clear();
                var files = Directory.GetFiles(SaveFolder, "*.json").OrderByDescending(f => File.GetLastWriteTimeUtc(f)).ToArray();
                int idx = 1;
                foreach (var f in files)
                {
                    var pkg = ReadOwnerListFromFile(f, out var name, out var savedAt, out var guildNames);
                    var displayName = name ?? Path.GetFileNameWithoutExtension(f);
                    var dt = savedAt ?? File.GetLastWriteTime(f);
                    grid.Rows.Add(idx.ToString(), displayName, dt.ToString("MM/dd HH:mm"));
                    grid.Rows[grid.Rows.Count - 1].Tag = f;
                    idx++;
                }
            }

            btnSave.Click += (_, __) =>
            {
                var saveName = (nameBox.Text ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(saveName))
                {
                    MessageBox.Show("保存名を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    var list = new List<OwnerDto>();
                    foreach (var key in _owner.GetAllKeys())
                    {
                        list.Add(new OwnerDto { Day = key.day, NodeId = key.nodeId, TeamId = _owner.GetOwner(key.day, key.nodeId) });
                    }

                    var pkg = new SavePackage { Name = saveName, SavedAt = DateTime.Now, Data = list, GuildNames = _teamNames };
                    var safe = string.Concat(saveName.Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch));
                    var fn = DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + (string.IsNullOrEmpty(safe) ? "save" : safe) + ".json";
                    var path = Path.Combine(SaveFolder, fn);
                    var json = JsonSerializer.Serialize(pkg, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(path, json);
                    SetStatus($"所持状態を {path} に保存しました");
                    RefreshList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "保存失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            grid.CellContentClick += (_, e) =>
            {
                if (e.RowIndex < 0) return;
                var row = grid.Rows[e.RowIndex];
                var path = row.Tag as string;
                if (path is null) return;

                if (grid.Columns[e.ColumnIndex].Name == "Load")
                {
                    try
                    {
                        string? tmpName = null;
                        DateTime? tmpSaved = null;
                        string[]? tmpGuildNames = null;
                        var arr = ReadOwnerListFromFile(path, out tmpName, out tmpSaved, out tmpGuildNames);
                        if (arr is null) throw new InvalidOperationException("ファイルの内容が不正です");
                        _owner.Clear();
                        foreach (var o in arr) _owner.SetOwner(o.Day, o.NodeId, o.TeamId);
                        if (tmpGuildNames is not null)
                        {
                            for (int i = 0; i < TeamCount && i < tmpGuildNames.Length; i++)
                            {
                                _teamNames[i] = tmpGuildNames[i];
                                // update textboxes (this will also trigger TextChanged handlers to refresh displays)
                                if (_teamNameBoxes is not null && _teamNameBoxes.Length > i && _teamNameBoxes[i] is not null)
                                    _teamNameBoxes[i].Text = tmpGuildNames[i];
                            }
                            UpdateTeamNameDisplays();
                        }
                        RefreshLeftGridForDay();
                        RefreshTotals();
                        _mapBox.Invalidate();
                        SetStatus($"{path} を読み込みました");
                        dlg.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "読込失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (grid.Columns[e.ColumnIndex].Name == "Delete")
                {
                    if (MessageBox.Show($"{path} を削除しますか？", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                    try
                    {
                        File.Delete(path);
                        RefreshList();
                        SetStatus($"{path} を削除しました");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "削除失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            };

            dlg.Controls.Add(grid);
            dlg.Controls.Add(sep);
            dlg.Controls.Add(topPanel);

            RefreshList();
            dlg.ShowDialog(this);
        }

        private void ExportGifDialog()
        {
            Directory.CreateDirectory(SaveFolder);
            using var sfd = new SaveFileDialog { Filter = "GIF|*.gif", FileName = "export.gif", InitialDirectory = SaveFolder };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExportGif(sfd.FileName);
                SetStatus($"GIF を {sfd.FileName} に出力しました");
            }
            catch (Exception ex)
            {
                try
                {
                    var full = ex.ToString();
                    File.WriteAllText(Path.Combine(SaveFolder, "gif_error.log"), full);
                }
                catch { }

                SetStatus("GIF 出力に失敗しました: " + ex.Message);
                MessageBox.Show(ex.ToString(), "GIF 出力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportGif(string path)
        {
#if USE_IMAGESHARP
            if (_mapBox.Image is null) throw new InvalidOperationException($"{MapImageFile} が読み込めていません");

            var bitmaps = new List<Bitmap>();
            for (int day = 1; day <= Days; day++)
            {
                var bmp = RenderDayToBitmap(day);
                bitmaps.Add(bmp);
            }

            // Use ImageSharp exporter
            ImageSharpGifExporter.Export(bitmaps, path);

            // dispose bitmaps
            foreach (var b in bitmaps) b.Dispose();
#else
            throw new NotSupportedException("GIF 出力には ImageSharp が必要です。プロジェクトに SixLabors.ImageSharp と SixLabors.ImageSharp.Drawing を追加し、コンパイルシンボル USE_IMAGESHARP を定義してください。");
#endif
        }

        // Render a day to a System.Drawing.Bitmap at the original map image size
        private Bitmap RenderDayToBitmap(int day)
        {
            if (_mapBox.Image is null) throw new InvalidOperationException($"{MapImageFile} が読み込めていません");
            var baseImg = new Bitmap(_mapBox.Image);
            var bmp = new Bitmap(baseImg.Width, baseImg.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmp);
            g.DrawImage(baseImg, 0, 0, bmp.Width, bmp.Height);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            DrawDayTeamTotals(g, bmp.Size, day);

            // Draw markers
            foreach (var node in _nodes.Values)
            {
                var owner = _owner.GetOwner(day, node.Id);
                if (owner is null) continue;

                DrawTeamMarker(g, node.Position, owner.Value, node.Point);
            }

            return bmp;
        }

        private void DrawDayTeamTotals(Graphics g, Size canvasSize, int day)
        {
            var totals = CalcTotalPoints(day, TeamCount, _nodes, _owner);
            using var teamLineFont = new Font(Font.FontFamily, 20, FontStyle.Bold);
            using var dayFont = new Font(Font.FontFamily, 28, FontStyle.Bold);

            var teamLines = new string[TeamCount];
            for (int i = 0; i < TeamCount; i++)
            {
                teamLines[i] = $"{_teamNames[i]}：{totals[i]}PT";
            }
            var dayText = $"{day}日目";

            var maxWidth = 0f;
            foreach (var line in teamLines)
            {
                var size = g.MeasureString(line, teamLineFont);
                if (size.Width > maxWidth) maxWidth = size.Width;
            }
            var daySize = g.MeasureString(dayText, dayFont);
            if (daySize.Width > maxWidth) maxWidth = daySize.Width;

            var margin = 16f;
            var padding = 12f;
            var lineHeight = g.MeasureString("A", teamLineFont).Height;
            var dayHeight = daySize.Height;
            var panelWidth = maxWidth + padding * 2;
            var panelHeight = lineHeight * TeamCount + dayHeight + padding * 2 + 6f;

            var panelRect = new RectangleF(
                canvasSize.Width - panelWidth - margin,
                canvasSize.Height - panelHeight - margin,
                panelWidth,
                panelHeight);
            using var panelBrush = new SolidBrush(Color.FromArgb(140, Color.Black));
            g.FillRectangle(panelBrush, panelRect);

            for (int i = 0; i < TeamCount; i++)
            {
                var y = panelRect.Y + padding + i * lineHeight;
                g.DrawString(teamLines[i], teamLineFont, Brushes.White, panelRect.X + padding, y);
            }

            var dayY = panelRect.Y + padding + TeamCount * lineHeight + 4f;
            g.DrawString(dayText, dayFont, Brushes.White, panelRect.X + padding, dayY);
        }

        private void BuildUi()
        {
            _split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                SplitterDistance = 760, // 左パネルは隠すので実質右側全体を使う
                Panel1MinSize = 0,
                FixedPanel = FixedPanel.Panel1
            };
            // 左パネルは折りたたむ（元の Panel1 を使わない）
            _split.Panel1Collapsed = true;
            Controls.Add(_split);

            // RIGHT: MAP + コントロール + 合計（ここに左パネル内容を下部に追加する）
            _rightPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8) };
            _split.Panel2.Controls.Add(_rightPanel);

            // 上部コントロール行
            var topControls = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 80, // 少し高さを増やす
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            _rightPanel.Controls.Add(topControls);

            // 日数選択
            var dayBox = new GroupBox
            {
                Text = "表示日",
                Width = 190,
                Height = 64
            };
            topControls.Controls.Add(dayBox);

            _daySelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = Days,
                Value = 1,
                Width = 100,
                Location = new Point(16, 26)
            };
            _daySelector.ValueChanged += (_, __) =>
            {
                _mapBox.Invalidate();
                RefreshLeftGridForDay();
                if (_copyPrevButton is not null)
                    _copyPrevButton.Enabled = SelectedDay > 1;
                SetStatus($"表示日を {SelectedDay} 日目に変更");
            };
            dayBox.Controls.Add(_daySelector);

            // 前日のコピーボタン（表示日の右側）
            _copyPrevButton = new Button
            {
                Text = "前日コピー",
                AutoSize = true,
                Height = 32,
                Width = 100,
                Margin = new Padding(8, 24, 8, 8),
                Enabled = SelectedDay > 1
            };
            _copyPrevButton.Click += (_, __) => CopyPreviousDayOwners();
            topControls.Controls.Add(_copyPrevButton);

            // 現在の入力値を保存するボタン
            _saveButton = new Button
            {
                Text = "一時保存",
                AutoSize = true,
                Height = 32,
                Width = 80,
                Margin = new Padding(4, 24, 8, 8)
            };
            _saveButton.Click += (_, __) => SaveOwnersToFile();
            topControls.Controls.Add(_saveButton);

            // 保存（名前を付けて）
            _saveAsButton = new Button
            {
                Text = "保存管理",
                AutoSize = true,
                Height = 32,
                Width = 140,
                Margin = new Padding(4, 24, 8, 8)
            };
            _saveAsButton.Click += (_, __) => SaveOwnersToFileAs();
            topControls.Controls.Add(_saveAsButton);

            // 読込ボタンは保存管理ダイアログで扱うため削除

            // GIF 出力
            _exportGifButton = new Button
            {
                Text = "GIF出力",
                AutoSize = true,
                Height = 32,
                Width = 100,
                Margin = new Padding(4, 24, 8, 8)
            };
            _exportGifButton.Click += (_, __) => ExportGifDialog();
            topControls.Controls.Add(_exportGifButton);

            // チーム選択
            _teamGroup = new GroupBox
            {
                Text = "入力チーム（クリックで割当）",
                Width = 430,
                Height = 64
            };
            topControls.Controls.Add(_teamGroup);

            _teamRadios = new RadioButton[TeamCount];
            for (int i = 0; i < TeamCount; i++)
            {
                _teamRadios[i] = new RadioButton
                {
                    Text = _teamNames[i],
                    AutoSize = true,
                    Location = new Point(16 + i * 90, 28),
                    Checked = (i == 0)
                };
                _teamGroup.Controls.Add(_teamRadios[i]);
            }

            // 操作説明
            var help = new Label
            {
                AutoSize = true,
                Text = "左クリック：選択チームに割当 / 同じ所持なら解除\n右クリック：解除",
                Padding = new Padding(6, 10, 6, 0)
            };
            topControls.Controls.Add(help);

            // MAP（PictureBox） — ここが上部のメイン領域
            var mapScrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            _rightPanel.Controls.Add(mapScrollPanel);

            _mapBox = new PictureBox
            {
                SizeMode = PictureBoxSizeMode.AutoSize
            };
            _mapBox.Paint += MapBox_Paint;
            _mapBox.MouseClick += MapBox_MouseClick;

            mapScrollPanel.Controls.Add(_mapBox);

            // 下段の一覧パネル
            var leftPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 220,
                Padding = new Padding(8)
            };
            _rightPanel.Controls.Add(leftPanel);

            var leftTitle = new Label
            {
                Text = "日別の所持拠点一覧（MAPクリックで更新）",
                AutoSize = true,
                Dock = DockStyle.Fill
            };

            var namePanel = new Panel
            {
                Height = 32,
                Padding = new Padding(4),
                Dock = DockStyle.Fill
            };

            _teamNameBoxes = new TextBox[TeamCount];
            for (int i = 0; i < TeamCount; i++)
            {
                var tb = new TextBox
                {
                    Width = 180,
                    Location = new Point(8 + i * 190, 4),
                    Text = _teamNames[i]
                };
                int idx = i;
                tb.TextChanged += (_, __) =>
                {
                    _teamNames[idx] = tb.Text;
                    UpdateTeamNameDisplays();
                };
                _teamNameBoxes[i] = tb;
                namePanel.Controls.Add(tb);
            }

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ScrollBars = ScrollBars.Vertical,
                AllowUserToResizeColumns = false,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };
            var leftLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0)
            };
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            leftLayout.Controls.Add(leftTitle, 0, 0);
            leftLayout.Controls.Add(namePanel, 0, 1);
            leftLayout.Controls.Add(_grid, 0, 2);

            leftPanel.Controls.Add(leftLayout);

            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Day", HeaderText = "日" });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "A", HeaderText = $"{_teamNames[0]} 所持拠点" });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "B", HeaderText = $"{_teamNames[1]} 所持拠点" });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "C", HeaderText = $"{_teamNames[2]} 所持拠点" });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "D", HeaderText = $"{_teamNames[3]} 所持拠点" });

            if (_grid.Columns.Contains("Day"))
            {
                _grid.Columns["Day"].FillWeight = 8;
                _grid.Columns["Day"].MinimumWidth = 60;
            }
            var teamFill = 23;
            foreach (var colName in new[] { "A", "B", "C", "D" })
            {
                if (_grid.Columns.Contains(colName))
                {
                    _grid.Columns[colName].FillWeight = teamFill;
                    _grid.Columns[colName].MinimumWidth = 20;
                }
            }

            // 下部：合計（最下部に配置）
            _totalsGroup = new GroupBox
            {
                Text = "6日間 合計ポイント",
                Dock = DockStyle.Bottom,
                Height = 92
            };
            _rightPanel.Controls.Add(_totalsGroup);

            _totalLabels = new Label[TeamCount];
            for (int i = 0; i < TeamCount; i++)
            {
                _totalLabels[i] = new Label
                {
                    AutoSize = true,
                    Location = new Point(16 + i * 200, 34),
                    Font = new Font(Font.FontFamily, 12, FontStyle.Bold),
                    Text = $"{_teamNames[i]}: 0"
                };
                _totalsGroup.Controls.Add(_totalLabels[i]);
            }

            // 一番下：ステータス
            _status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                Text = "",
                ForeColor = Color.DimGray
            };
            Controls.Add(_status);
        }

        // 前日の所持状態を選択日へコピー
        private void CopyPreviousDayOwners()
        {
            int day = SelectedDay;
            if (day <= 1)
            {
                SetStatus("前日のコピーは2日目以降で利用できます");
                return;
            }

            int prev = day - 1;

            // prev日の各拠点の所有をコピー（null もコピー）
            foreach (var node in _nodes.Values)
            {
                var ownerPrev = _owner.GetOwner(prev, node.Id);
                _owner.SetOwner(day, node.Id, ownerPrev);
            }

            RefreshLeftGridForDay();
            RefreshTotals();
            _mapBox.Invalidate();

            SetStatus($"{prev}日目の状態を {day}日目にコピーしました");
        }

        // 所持状態を JSON ファイルに保存
        private void SaveOwnersToFile()
        {
            try
            {
                var list = new List<OwnerDto>();
                // days 1..Days
                foreach (var key in _owner.GetAllKeys())
                {
                    list.Add(new OwnerDto { Day = key.day, NodeId = key.nodeId, TeamId = _owner.GetOwner(key.day, key.nodeId) });
                }

                // 保存時にギルド名も含める（パッケージ形式）
                var pkg = new SavePackage { Name = null, SavedAt = DateTime.Now, Data = list, GuildNames = _teamNames };
                var json = JsonSerializer.Serialize(pkg, new JsonSerializerOptions { WriteIndented = true });
                // 一時保存として save フォルダ内に書き出す（拡張子 .tmp を付与）
                Directory.CreateDirectory(SaveFolder);
                var path = Path.Combine(SaveFolder, "owners.json.tmp");
                File.WriteAllText(path, json);
                SetStatus($"一時保存しました: {path}");
            }
            catch (Exception ex)
            {
                SetStatus("保存に失敗しました: " + ex.Message);
            }
        }

        private void LoadOwnersIfExists()
        {
            try
            {
                var path1 = Path.Combine(SaveFolder, "owners.json");
                var pathTmp = Path.Combine(SaveFolder, "owners.json.tmp");

                // Prefer the most recently modified of owners.json and owners.json.tmp
                string? path = null;
                if (File.Exists(path1) && File.Exists(pathTmp))
                {
                    path = File.GetLastWriteTimeUtc(path1) >= File.GetLastWriteTimeUtc(pathTmp) ? path1 : pathTmp;
                }
                else if (File.Exists(path1))
                {
                    path = path1;
                }
                else if (File.Exists(pathTmp))
                {
                    path = pathTmp;
                }

                if (path is null) return;

                var arr = ReadOwnerListFromFile(path, out var name, out var savedAt, out var guildNames);
                if (arr is null) return;
                _owner.Clear();
                foreach (var o in arr)
                {
                    _owner.SetOwner(o.Day, o.NodeId, o.TeamId);
                }

                if (guildNames is not null)
                {
                    for (int i = 0; i < TeamCount && i < guildNames.Length; i++)
                    {
                        _teamNames[i] = guildNames[i];
                        if (_teamNameBoxes is not null && _teamNameBoxes.Length > i && _teamNameBoxes[i] is not null)
                            _teamNameBoxes[i].Text = guildNames[i];
                    }
                    UpdateTeamNameDisplays();
                }

                SetStatus($"{path} から所持状態を読み込みました");
            }
            catch (Exception ex)
            {
                SetStatus("owners.json の読み込み失敗: " + ex.Message);
            }
        }

        private sealed class OwnerDto
        {
            public int Day { get; set; }
            public int NodeId { get; set; }
            public int? TeamId { get; set; }
        }

        private sealed class SavePackage
        {
            public string? Name { get; set; }
            public DateTime SavedAt { get; set; }
            public List<OwnerDto>? Data { get; set; }
            public string[]? GuildNames { get; set; }
        }

        // 保存形式（配列形式/パッケージ形式）の両方に対応して読込
        private static List<OwnerDto>? ReadOwnerListFromFile(string path, out string? name, out DateTime? savedAt, out string[]? guildNames)
        {
            name = null;
            savedAt = null;
            guildNames = null;
            var json = File.ReadAllText(path);
            try
            {
                var pkg = JsonSerializer.Deserialize<SavePackage>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (pkg is not null && pkg.Data is not null)
                {
                    name = pkg.Name;
                    savedAt = pkg.SavedAt == default ? (DateTime?)File.GetLastWriteTime(path) : pkg.SavedAt;
                    if (pkg.GuildNames is not null)
                        guildNames = pkg.GuildNames;
                    return pkg.Data;
                }
            }
            catch { }

            try
            {
                var arr = JsonSerializer.Deserialize<List<OwnerDto>>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (arr is not null)
                {
                    savedAt = File.GetLastWriteTime(path);
                    return arr;
                }
            }
            catch { }

            return null;
        }

        private void LoadDataOrFallback()
        {
            // map.jpg
            if (File.Exists(MapImageFile))
            {
                using var fs = new FileStream(MapImageFile, FileMode.Open, FileAccess.Read);
                _mapBox.Image = Image.FromStream(fs);
            }
            else
            {
                _mapBox.Image = null;
                SetStatus($"警告: {MapImageFile} が見つかりません（実行フォルダに置いてください）");
            }

            // nodes.json
            if (File.Exists(NodesJsonFile))
            {
                var json = File.ReadAllText(NodesJsonFile);
                var list = JsonSerializer.Deserialize<List<NodeDto>>(json, JsonOpts());
                if (list is null || list.Count == 0)
                    throw new InvalidOperationException("nodes.json の内容が空です");

                _nodes = list.ToDictionary(
                    x => x.Id,
                    x => new Node
                    {
                        Id = x.Id,
                        Name = x.Name ?? $"Node{x.Id}",
                        Type = x.Type,
                        Point = x.Point,
                        Position = new Point(x.Position.X, x.Position.Y),
                        Neighbors = x.Neighbors?.ToList() ?? new List<int>()
                    });

                SetStatus($"nodes.json を読み込みました（拠点数: {_nodes.Count}）");
            }
            else
            {
                // フォールバック（最小サンプル）
                _nodes = CreateFallbackNodes();
                SetStatus($"警告: {NodesJsonFile} が見つからないため、サンプル拠点で起動しました。");
            }
        }

        private void LoadTeamIcons()
        {
            for (int i = 0; i < TeamCount; i++)
            {
                _teamIcons[i]?.Dispose();
                _teamIcons[i] = null;

                var iconPath = TeamIconFiles[i];
                if (!File.Exists(iconPath)) continue;

                using var fs = new FileStream(iconPath, FileMode.Open, FileAccess.Read);
                using var src = Image.FromStream(fs);
                _teamIcons[i] = new Bitmap(src);
            }
        }

        private void InitializeGridRows()
        {
            _grid.Rows.Clear();
            for (int d = 1; d <= Days; d++)
            {
                _grid.Rows.Add($"{d}日目", "", "", "", "");
            }
        }

        private void RefreshLeftGridForDay()
        {
            // 6日分すべての行を更新（最小構成なので毎回全更新でOK）
            for (int day = 1; day <= Days; day++)
            {
                var row = _grid.Rows[day - 1];
                row.Cells[0].Value = $"{day}日目";
                for (int team = 0; team < TeamCount; team++)
                {
                    row.Cells[1 + team].Value = BuildTeamOwnedText(day, team);
                }
            }
        }

        private string BuildTeamOwnedText(int day, int teamId)
        {
            var ownedNodes = _nodes.Values
                .Where(n => _owner.GetOwner(day, n.Id) == teamId)
                .OrderByDescending(n => n.Point)
                .ThenBy(n => n.Id)
                .ToList();

            if (ownedNodes.Count == 0) return string.Empty;

            // 当日のみのポイント合計を先頭に表示
            var total = ownedNodes.Sum(n => n.Point);

            // 城や教会は ID を表示しない（寺院は ID を表示）
            var parts = ownedNodes.Select(n =>
                (n.Type == NodeType.Castle || n.Type == NodeType.Church)
                    ? $"{n.Name}({n.Point})"
                    : $"{n.Id}:{n.Name}({n.Point})"
            );

            return $"{total}PT：{string.Join(", ", parts)}";
        }

        private void RefreshTotals()
        {
            var totals = CalcTotalPoints(Days, TeamCount, _nodes, _owner);
            for (int i = 0; i < TeamCount; i++)
            {
                _totalLabels[i].Text = $"{_teamNames[i]}: {totals[i]} PT";
                _totalLabels[i].ForeColor = _teamColors[i];
            }
        }

        // チーム名変更時にラジオボタン、グリッド列ヘッダ、合計ラベルを更新
        private void UpdateTeamNameDisplays()
        {
            var colNames = new[] { "A", "B", "C", "D" };
            for (int i = 0; i < TeamCount; i++)
            {
                if (_teamRadios is not null && _teamRadios.Length > i && _teamRadios[i] is not null)
                    _teamRadios[i].Text = _teamNames[i];

                if (_grid.Columns.Contains(colNames[i]))
                    _grid.Columns[colNames[i]].HeaderText = $"{_teamNames[i]} 所持拠点";
            }

            // 合計ラベルは RefreshTotals で正しく更新する
            RefreshTotals();
        }

        private void MapBox_MouseClick(object? sender, MouseEventArgs e)
        {
            if (_nodes.Count == 0) return;

            // PictureBox(ZOOM) → 画像座標へ変換
            if (_mapBox.Image is null)
            {
                SetStatus($"{MapImageFile} が読み込めていません");
                return;
            }

            var imgPoint = TranslateToImagePoint(_mapBox, e.Location);
            if (imgPoint is null)
            {
                SetStatus("クリック位置が画像範囲外です");
                return;
            }

            var hit = HitTestNode(_nodes, imgPoint.Value, HitRadius);
            if (hit is null)
            {
                SetStatus($"拠点にヒットしませんでした（x={imgPoint.Value.X}, y={imgPoint.Value.Y}）");
                return;
            }

            int nodeId = hit.Value;
            int day = SelectedDay;

            if (e.Button == MouseButtons.Right)
            {
                _owner.SetOwner(day, nodeId, null);
                AfterOwnerChanged(day, nodeId, "解除（右クリック）");
                return;
            }

            // 左クリック：選択チームに割当 / 同じなら解除
            int team = SelectedTeam;
            if (team < 0) team = 0;

            var current = _owner.GetOwner(day, nodeId);
            if (current == team)
            {
                _owner.SetOwner(day, nodeId, null);
                AfterOwnerChanged(day, nodeId, $"解除（{_teamNames[team]} だったのでトグル）");
            }
            else
            {
                _owner.SetOwner(day, nodeId, team);
                AfterOwnerChanged(day, nodeId, $"割当 → {_teamNames[team]}");
            }
        }

        private void AfterOwnerChanged(int day, int nodeId, string action)
        {
            RefreshLeftGridForDay();
            RefreshTotals();
            _mapBox.Invalidate();

            var node = _nodes[nodeId];
            SetStatus($"{day}日目: {node.Id}:{node.Name}({node.Point}) を {action}");
        }

        private void MapBox_Paint(object? sender, PaintEventArgs e)
        {
            if (_mapBox.Image is null || _nodes.Count == 0) return;

            // ZOOM表示で描画がズレるので、画像座標→画面座標に変換して描く
            foreach (var node in _nodes.Values)
            {
                var owner = _owner.GetOwner(SelectedDay, node.Id);
                if (owner is null) continue;

                var screenPoint = TranslateToControlPoint(_mapBox, node.Position);
                if (screenPoint is null) continue;

                DrawTeamMarker(e.Graphics, screenPoint.Value, owner.Value, node.Point);
            }
        }

        private void DrawTeamMarker(Graphics g, Point center, int teamId, int point)
        {
            if (teamId < 0 || teamId >= TeamCount) return;

            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            var icon = _teamIcons[teamId];
            if (icon is not null)
            {
                var rect = new Rectangle(center.X - IconDrawSize / 2, center.Y - IconDrawSize / 2, IconDrawSize, IconDrawSize);
                g.DrawImage(icon, rect);
            }
            else
            {
                using var brush = new SolidBrush(_teamColors[teamId]);
                g.FillEllipse(brush, center.X - DrawRadius, center.Y - DrawRadius, DrawRadius * 2, DrawRadius * 2);
                g.DrawEllipse(Pens.Black, center.X - DrawRadius, center.Y - DrawRadius, DrawRadius * 2, DrawRadius * 2);
            }

            var text = point.ToString();
            var sz = g.MeasureString(text, Font);
            g.DrawString(text, Font, Brushes.White, center.X - sz.Width / 2, center.Y - sz.Height / 2);
        }

        // =============================
        // ヒットテスト（画像座標で判定）
        // =============================
        private static int? HitTestNode(IReadOnlyDictionary<int, Node> nodes, Point p, int radius)
        {
            int? bestId = null;
            double best = double.MaxValue;
            int r2 = radius * radius;

            foreach (var n in nodes.Values)
            {
                int dx = n.Position.X - p.X;
                int dy = n.Position.Y - p.Y;
                int d2 = dx * dx + dy * dy;
                if (d2 > r2) continue;
                if (d2 < best)
                {
                    best = d2;
                    bestId = n.Id;
                }
            }
            return bestId;
        }

        // =========================================
        // PictureBox(ZOOM) 画像座標 <-> 画面座標変換
        // =========================================
        private static Point? TranslateToImagePoint(PictureBox pb, Point controlPoint)
        {
            if (pb.Image is null) return null;

            var img = pb.Image;
            var rect = GetImageDisplayRectangle(pb);

            if (!rect.Contains(controlPoint))
                return null;

            // 表示矩形内の相対位置を画像サイズへマップ
            float xRatio = (float)img.Width / rect.Width;
            float yRatio = (float)img.Height / rect.Height;

            int x = (int)((controlPoint.X - rect.X) * xRatio);
            int y = (int)((controlPoint.Y - rect.Y) * yRatio);

            return new Point(x, y);
        }

        private static Point? TranslateToControlPoint(PictureBox pb, Point imagePoint)
        {
            if (pb.Image is null) return null;

            var img = pb.Image;
            var rect = GetImageDisplayRectangle(pb);

            if (img.Width == 0 || img.Height == 0) return null;

            float xRatio = (float)rect.Width / img.Width;
            float yRatio = (float)rect.Height / img.Height;

            int x = (int)(rect.X + imagePoint.X * xRatio);
            int y = (int)(rect.Y + imagePoint.Y * yRatio);

            return new Point(x, y);
        }

        private static Rectangle GetImageDisplayRectangle(PictureBox pb)
        {
            // SizeMode.Zoom の表示矩形を計算
            if (pb.Image is null) return Rectangle.Empty;

            int imgW = pb.Image.Width;
            int imgH = pb.Image.Height;

            int boxW = pb.ClientSize.Width;
            int boxH = pb.ClientSize.Height;

            if (imgW == 0 || imgH == 0 || boxW == 0 || boxH == 0)
                return Rectangle.Empty;

            float imgAspect = (float)imgW / imgH;
            float boxAspect = (float)boxW / boxH;

            int drawW, drawH;
            if (imgAspect > boxAspect)
            {
                // 横がフィット
                drawW = boxW;
                drawH = (int)(boxW / imgAspect);
            }
            else
            {
                // 縦がフィット
                drawH = boxH;
                drawW = (int)(boxH * imgAspect);
            }

            int x = (boxW - drawW) / 2;
            int y = (boxH - drawH) / 2;

            return new Rectangle(x, y, drawW, drawH);
        }

        // =============================
        // ポイント計算
        // =============================
        private static int[] CalcTotalPoints(int days, int teamCount, IReadOnlyDictionary<int, Node> nodes, OwnerState owner)
        {
            var totals = new int[teamCount];

            for (int day = 1; day <= days; day++)
            {
                foreach (var node in nodes.Values)
                {
                    var t = owner.GetOwner(day, node.Id);
                    if (t is null) continue;
                    totals[t.Value] += node.Point;
                }
            }
            return totals;
        }

        private void SetStatus(string text)
        {
            _status.Text = text;
        }

        // =============================
        // JSON / Model
        // =============================
        private static JsonSerializerOptions JsonOpts()
        {
            var opts = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };
            opts.Converters.Add(new JsonStringEnumConverter());
            return opts;
        }

        public enum NodeType { Temple, Castle, Church }

        public sealed class Node
        {
            public int Id { get; init; }
            public string Name { get; init; } = "";
            public NodeType Type { get; init; }
            public int Point { get; init; }
            public Point Position { get; init; }
            public List<int> Neighbors { get; init; } = new();
        }

        public sealed class NodeDto
        {
            public int Id { get; set; }
            public string? Name { get; set; }
            public NodeType Type { get; set; }
            public int Point { get; set; }
            public PosDto Position { get; set; } = new();
            public int[]? Neighbors { get; set; }
        }

        public sealed class PosDto
        {
            public int X { get; set; }
            public int Y { get; set; }
        }

        // owner[day][nodeId] = teamId?
        public sealed class OwnerState
        {
            private readonly int _days;
            private readonly int _teamCount;

            // day index: 1..days, nodeId: any int
            private readonly Dictionary<(int day, int nodeId), int> _map = new();

            public OwnerState(int days, int teamCount)
            {
                _days = days;
                _teamCount = teamCount;
            }

            public int? GetOwner(int day, int nodeId)
            {
                if (day < 1 || day > _days) throw new ArgumentOutOfRangeException(nameof(day));
                return _map.TryGetValue((day, nodeId), out var t) ? t : (int?)null;
            }

            public void SetOwner(int day, int nodeId, int? teamId)
            {
                if (day < 1 || day > _days) throw new ArgumentOutOfRangeException(nameof(day));
                if (teamId is not null && (teamId < 0 || teamId >= _teamCount))
                    throw new ArgumentOutOfRangeException(nameof(teamId));

                var key = (day, nodeId);
                if (teamId is null)
                    _map.Remove(key);
                else
                    _map[key] = teamId.Value;
            }

            // 所持情報のキー一覧を取得
            public IEnumerable<(int day, int nodeId)> GetAllKeys()
            {
                return _map.Keys.ToList();
            }

            // すべてクリア
            public void Clear()
            {
                _map.Clear();
            }
        }

        // =============================
        // フォールバック拠点（nodes.json無い場合）
        // =============================
        private static Dictionary<int, Node> CreateFallbackNodes()
        {
            // とりあえず3つだけ置く（起動確認用）
            return new List<Node>
            {
                new Node { Id = 1, Name = "アイン(例)", Type = NodeType.Temple, Point = 4, Position = new Point(1000, 700) },
                new Node { Id = 2, Name = "城(例)", Type = NodeType.Castle, Point = 2, Position = new Point(850, 500) },
                new Node { Id = 3, Name = "教会(例)", Type = NodeType.Church, Point = 1, Position = new Point(1150, 520) },
            }.ToDictionary(n => n.Id, n => n);
        }
    }
}