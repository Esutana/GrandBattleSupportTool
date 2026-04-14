using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using System.Drawing.Imaging;
#if USE_IMAGESHARP
using SixLabors.ImageSharp.Formats.Gif;
using SixLabors.ImageSharp.PixelFormats;
using ImageSharpImage = SixLabors.ImageSharp.Image;
#endif

namespace GrandBattleSupport
{
    public partial class MainForm : Form
    {
        private const int Days = 6;
        private const int TeamCount = 4;
        private const int TeamCount2 = 5;
        // チーム色（自由に変えてOK）
        private readonly Color[] _teamColors = new[]
        {
            Color.FromArgb(220, 231, 76, 60),  // Team A (赤)
            Color.FromArgb(220, 52, 152, 219), // Team B (青)
            Color.FromArgb(220, 46, 204, 113), // Team C (緑)
            Color.FromArgb(220, 241, 196, 15), // Team D (黄)
        };

        private readonly string[] _teamNames = new[] { "A", "B", "C", "D" };

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
        private Button _copyPrevButton = null!; // ← 追加
        private Button _saveButton = null!; // ← 現在値保存ボタン
        private Button _saveAsButton = null!; // 保存（名前を付けて）
        private Button _loadButton = null!; // 読込
        private Button _exportGifButton = null!; // GIF 出力

        // 選択状態
        private int SelectedDay => (int)_daySelector.Value; // 1..6
        private int SelectedTeam => Array.FindIndex(_teamRadios, r => r.Checked); // 0..3

        // クリック判定（円の半径）
        private const int HitRadius = 22;
        private const int DrawRadius = 14;

        // ファイル名
        private const string NodesJsonFile = "nodes.json";
        // map.jpg 取得のためのパス（実行ファイルからの相対パス）
        private static readonly string MapImageFile = Path.Combine("image", "map.jpg");

        public MainForm()
        {
            Text = "メメントモリ グラバトポイント計算（最小動作）";
            Width = 1400;
            Height = 900;
            StartPosition = FormStartPosition.CenterScreen;

            BuildUi();
            LoadDataOrFallback();
            LoadOwnersIfExists();
            InitializeGridRows();
            RefreshLeftGridForDay();
            RefreshTotals();
            _mapBox.Invalidate();
        }

        // 名前を付けて保存（ファイル選択ダイアログ）
        private void SaveOwnersToFileAs()
        {
            using var sfd = new SaveFileDialog { Filter = "JSON files|*.json|All files|*.*", FileName = "owners.json" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var list = new List<OwnerDto>();
                foreach (var key in _owner.GetAllKeys())
                {
                    list.Add(new OwnerDto { Day = key.day, NodeId = key.nodeId, TeamId = _owner.GetOwner(key.day, key.nodeId) });
                }
                var json = JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(sfd.FileName, json);
                SetStatus($"所持状態を {sfd.FileName} に保存しました");
            }
            catch (Exception ex)
            {
                SetStatus("保存に失敗しました: " + ex.Message);
            }
        }

        private void LoadOwnersFromFileDialog()
        {
            using var ofd = new OpenFileDialog { Filter = "JSON files|*.json|All files|*.*", FileName = "owners.json" };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            try
            {
                var json = File.ReadAllText(ofd.FileName);
                var list = JsonSerializer.Deserialize<List<OwnerDto>>(json, JsonOpts());
                if (list is null) return;

                // 既存状態をクリア
                _owner.Clear();

                foreach (var o in list)
                {
                    _owner.SetOwner(o.Day, o.NodeId, o.TeamId);
                }

                RefreshLeftGridForDay();
                RefreshTotals();
                _mapBox.Invalidate();
                SetStatus($"{ofd.FileName} から所持状態を読み込みました");
            }
            catch (Exception ex)
            {
                SetStatus("読込に失敗しました: " + ex.Message);
            }
        }

        private void ExportGifDialog()
        {
            using var sfd = new SaveFileDialog { Filter = "GIF|*.gif", FileName = "export.gif" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExportGif(sfd.FileName);
                SetStatus($"GIF を {sfd.FileName} に出力しました");
            }
            catch (Exception ex)
            {
                // Write full exception to log and show a message box for easier debugging
                try
                {
                    var full = ex.ToString();
                    File.WriteAllText("gif_error.log", full);
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

            // Draw markers
            foreach (var node in _nodes.Values)
            {
                var owner = _owner.GetOwner(day, node.Id);
                if (owner is null) continue;

                var teamColor = _teamColors[owner.Value];
                using var brush = new SolidBrush(teamColor);
                int r = DrawRadius;
                g.FillEllipse(brush, node.Position.X - r, node.Position.Y - r, r * 2, r * 2);
                g.DrawEllipse(Pens.Black, node.Position.X - r, node.Position.Y - r, r * 2, r * 2);

                // point text
                var text = node.Point.ToString();
                var sz = g.MeasureString(text, Font);
                g.DrawString(text, Font, Brushes.White, node.Position.X - sz.Width / 2, node.Position.Y - sz.Height / 2);
            }

            // Optionally, draw day number
            using var dayFont = new Font(Font.FontFamily, 48, FontStyle.Bold);
            var dayText = $"{day}日目";
            var daySz = g.MeasureString(dayText, dayFont);
            var margin = 16;
            var rect = new RectangleF(bmp.Width - (float)daySz.Width - margin, bmp.Height - (float)daySz.Height - margin, daySz.Width, daySz.Height);
            using var bgBrush = new SolidBrush(Color.FromArgb(128, Color.Black));
            g.FillRectangle(bgBrush, rect);
            g.DrawString(dayText, dayFont, Brushes.White, rect.Location);

            return bmp;
        }

#if USE_IMAGESHARP
        // Convert Bitmap to Rgba32[] for ImageSharp
        private static Rgba32[] BitmapToRgba32(Bitmap bmp)
        {
            var w = bmp.Width;
            var h = bmp.Height;
            var arr = new Rgba32[w * h];

            var rect = new Rectangle(0, 0, w, h);
            var data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                var bytes = new byte[Math.Abs(data.Stride) * h];
                System.Runtime.InteropServices.Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
                int stride = data.Stride;
                for (int y = 0; y < h; y++)
                {
                    for (int x = 0; x < w; x++)
                    {
                        int idx = y * stride + x * 4;
                        byte b = bytes[idx + 0];
                        byte g = bytes[idx + 1];
                        byte r = bytes[idx + 2];
                        byte a = bytes[idx + 3];
                        arr[y * w + x] = new Rgba32(r, g, b, a);
                    }
                }
            }
            finally
            {
                bmp.UnlockBits(data);
            }

            return arr;
        }
#else
        private static object BitmapToRgba32(Bitmap bmp)
        {
            throw new NotSupportedException("Bitmap->ImageSharp 変換には ImageSharp が必要です。プロジェクトにパッケージを追加し、USE_IMAGESHARP を定義してください。");
        }
#endif

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
                RefreshLeftGridForDay(); // 一応最新反映
                // ボタンの有効/無効を切り替え
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
                Text = "保存",
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
                Text = "名前を付けて保存",
                AutoSize = true,
                Height = 32,
                Width = 140,
                Margin = new Padding(4, 24, 8, 8)
            };
            _saveAsButton.Click += (_, __) => SaveOwnersToFileAs();
            topControls.Controls.Add(_saveAsButton);

            // 読込
            _loadButton = new Button
            {
                Text = "読込",
                AutoSize = true,
                Height = 32,
                Width = 80,
                Margin = new Padding(4, 24, 8, 8)
            };
            _loadButton.Click += (_, __) => LoadOwnersFromFileDialog();
            topControls.Controls.Add(_loadButton);

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

            // LEFT相当のパネルを MAP の下面に置く（ここが以前の左パネルの中身）
            var leftPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 220,
                Padding = new Padding(8)
            };
            _rightPanel.Controls.Add(leftPanel);

            // Use a TableLayoutPanel so we can place the guild name inputs directly above the grid rows
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
                    // 配列要素を更新して表示を反映
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
            // Compose with TableLayoutPanel: title (Auto), name inputs (Auto), grid (Fill)
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

                var json = JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText("owners.json", json);
                SetStatus("所持状態を owners.json に保存しました");
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
                if (!File.Exists("owners.json")) return;
                var json = File.ReadAllText("owners.json");
                var list = JsonSerializer.Deserialize<List<OwnerDto>>(json, JsonOpts());
                if (list is null) return;
                foreach (var o in list)
                {
                    _owner.SetOwner(o.Day, o.NodeId, o.TeamId);
                }
                SetStatus("owners.json から所持状態を読み込みました");
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
            var owned = _nodes.Values
                .Where(n => _owner.GetOwner(day, n.Id) == teamId)
                .OrderByDescending(n => n.Point)
                .ThenBy(n => n.Id)
                .Select(n => $"{n.Id}:{n.Name}({n.Point})");

            return string.Join(", ", owned);
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

                using var brush = new SolidBrush(_teamColors[owner.Value]);
                var r = DrawRadius;
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillEllipse(brush, screenPoint.Value.X - r, screenPoint.Value.Y - r, r * 2, r * 2);
                e.Graphics.DrawEllipse(Pens.Black, screenPoint.Value.X - r, screenPoint.Value.Y - r, r * 2, r * 2);

                // 点数の小表示
                var text = node.Point.ToString();
                var sz = e.Graphics.MeasureString(text, Font);
                e.Graphics.DrawString(text, Font, Brushes.White,
                    screenPoint.Value.X - sz.Width / 2,
                    screenPoint.Value.Y - sz.Height / 2);
            }
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