using DevExpress.Data;
using DevExpress.Data.Filtering;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraTab;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace SPR1
{
    public partial class frmSPR1 : Syb3.Form.RB1.Syb3FormRB_1cn_SR
    {
        private DataTable dtReport = null;
        private string activeConnString = string.Empty;
        private string activeEstateFilter = "";
        private string activePeriodFilter = "ALL";
        private PanelControl pnlFilterBar;
        private Dictionary<string, SimpleButton> filterButtons = new Dictionary<string, SimpleButton>();

        private BarButtonItem btnRibbonResetFilter;
        private BarButtonItem btnRibbonRefreshData;
        private BarButtonItem btnRibbonClearColFilter;

        private class SmoothScrollableControl : XtraScrollableControl
        {
            public SmoothScrollableControl()
            {
                this.SetStyle(ControlStyles.AllPaintingInWmPaint |
                              ControlStyles.UserPaint |
                              ControlStyles.OptimizedDoubleBuffer |
                              ControlStyles.ResizeRedraw, true);
                this.DoubleBuffered = true;
                this.UpdateStyles();
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams cp = base.CreateParams;
                    cp.ExStyle |= 0x02000000;
                    return cp;
                }
            }
        }

        private class SmoothDashboardCanvas : PanelControl
        {
            public SmoothDashboardCanvas()
            {
                this.SetStyle(ControlStyles.AllPaintingInWmPaint |
                              ControlStyles.UserPaint |
                              ControlStyles.OptimizedDoubleBuffer |
                              ControlStyles.ResizeRedraw, true);
                this.DoubleBuffered = true;
                this.UpdateStyles();
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams cp = base.CreateParams;
                    cp.ExStyle |= 0x02000000;
                    return cp;
                }
            }
        }

        private SmoothScrollableControl scrollDashboard;
        private SmoothDashboardCanvas pnlDashboardCanvas;
        private PanelControl pnlBannerHeader;
        private LabelControl lblBannerTitle;
        private LabelControl lblBannerSubtitle;

        private PanelControl pnlBannerRightContainer;
        private LabelControl lblBannerPeriodeVal;
        private LabelControl lblBannerTphVal;
        private LabelControl lblBannerMatangVal;
        private LabelControl lblBannerMentahVal;
        private LabelControl lblBannerAfkirVal;

        private PanelControl pnlSummaryContainer;
        private LabelControl lblKpiTotalTphVal;
        private LabelControl lblKpiTotalJanjangVal;
        private LabelControl lblKpiMatangVal;
        private LabelControl lblKpiMentahVal;
        private LabelControl lblKpiAfkirVal;

        private PanelControl pnlSectionEstateHeader;
        private PanelControl pnlRowEstates;
        private Dictionary<string, EstateCardControl> estateCardMap = new Dictionary<string, EstateCardControl>();

        private PanelControl pnlReportPhotoPanel;
        private SplitterControl splitterReportPhoto;
        private XtraTabControl tabRightPhoto;
        private XtraTabPage pagePhotoPreview;
        private PictureEdit picConditionViewer;
        private LabelControl lblTransactionTitle;
        private LabelControl lblEstateAndBlockInfo;
        private LabelControl lblQualitySummary;
        private LabelControl lblDetailsCount;
        private LabelControl lblKraniInfo;
        private LabelControl lblSyncStatusBadge;
        private LabelControl lblTimestampText;
        private LabelControl lblPhotoStatus;
        private SimpleButton btnClosePhotoPanel;

        private static readonly Color ColorPrimary = Color.FromArgb(0, 83, 150);
        private static readonly Color ColorPrimaryDark = Color.FromArgb(0, 56, 102);
        private static readonly Color ColorAccent = Color.FromArgb(29, 99, 237);
        private static readonly Color ColorBrandLight = Color.FromArgb(239, 246, 255);
        private static readonly Color ColorSlate900 = Color.FromArgb(15, 23, 42);
        private static readonly Color ColorSlate800 = Color.FromArgb(30, 41, 59);
        private static readonly Color ColorSlate600 = Color.FromArgb(71, 85, 105);
        private static readonly Color ColorSlate400 = Color.FromArgb(148, 163, 184);
        private static readonly Color ColorBorder = Color.FromArgb(226, 232, 240);
        private static readonly Color ColorCanvasBg = Color.FromArgb(248, 250, 252);

        private static readonly Color ColorMatang = Color.FromArgb(16, 185, 129);
        private static readonly Color ColorMatangBg = Color.FromArgb(220, 252, 231);
        private static readonly Color ColorMatangBorder = Color.FromArgb(134, 239, 172);

        private static readonly Color ColorMentah = Color.FromArgb(217, 119, 6);
        private static readonly Color ColorMentahBg = Color.FromArgb(254, 243, 199);
        private static readonly Color ColorMentahBorder = Color.FromArgb(253, 230, 138);

        private static readonly Color ColorAfkir = Color.FromArgb(220, 38, 38);
        private static readonly Color ColorAfkirBg = Color.FromArgb(254, 226, 226);
        private static readonly Color ColorAfkirBorder = Color.FromArgb(252, 165, 165);

        public static GraphicsPath CreateRoundedRectangle(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            if (radius <= 0)
            {
                path.AddRectangle(rect);
                return path;
            }
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        public static void EnableDoubleBuffering(Control c)
        {
            if (c == null) return;
            try
            {
                var prop = typeof(Control).GetProperty("DoubleBuffered",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                prop?.SetValue(c, true, null);
            }
            catch { }

            foreach (Control child in c.Controls)
            {
                EnableDoubleBuffering(child);
            }
        }

        private class EstateDef
        {
            public int Index { get; set; }
            public string EstateCode { get; set; }
            public string EstateName { get; set; }
            public string Description { get; set; }
        }

        private readonly List<EstateDef> estateDefinitions = new List<EstateDef>
        {
            new EstateDef { Index = 0, EstateCode = "MRE", EstateName = "MRE Estate", Description = "Kebun Inti MRE" },
            new EstateDef { Index = 1, EstateCode = "LBE", EstateName = "LBE Estate", Description = "Kebun Inti LBE" },
            new EstateDef { Index = 2, EstateCode = "BKE", EstateName = "BKE Estate", Description = "Kebun Inti BKE" },
            new EstateDef { Index = 3, EstateCode = "AGE", EstateName = "AGE Estate", Description = "Kebun Inti AGE" }
        };

        public frmSPR1()
        {
            InitializeComponent();
        }

        public frmSPR1(Syb3init1.mySyb3Init param_AppInit1, string[] param_App2, Object[] param_App3)
            : base(param_AppInit1)
        {
            InitializeComponent();
            _Syb3InitPub1 = param_AppInit1;
            _AppParam3 = param_App3;
        }

        private void frmSPR1_Load(object sender, EventArgs e)
        {
            if (this.DesignMode) return;
            Application.Idle += new EventHandler(frmSPR1_Idle);

            Syb3FormRB_1cn_reLoad1(sender, e, "EO", null, false, false, _Syb3InitPub1, _Syb3InitLoc1);

            _Syb3InitLoc1[_Syb3InitPub1.vfirstField] = "ID";

            if (_5Connect_BUnitDB != null)
            {
                RepositoryItemComboBox repoUnit = new RepositoryItemComboBox();
                repoUnit.Items.Add("ESTATE");
                repoUnit.TextEditStyle = TextEditStyles.DisableTextEditor;
                _5Connect_BUnitDB.Edit = repoUnit;
                _5Connect_BUnitDB.EditValue = "ESTATE";
                _5Connect_BUnitDB.Enabled = false;
            }

            SetupSheetsLayout();
            InitRibbonActionButtons();
            EnableDoubleBuffering(this);
        }

        private void frmSPR1_Shown(object sender, EventArgs e)
        {
            Syb3DBload(sender, "SPECTRA_DB", _tblConsolBA);
            if (_Syb3InitPub1 != null)
            {
                _Syb3InitPub1.loadWait1 = false;
            }

            EnsureBottomSheetsConfigured();
            if (scrollDashboard != null)
            {
                scrollDashboard.AutoScrollPosition = new Point(0, 0);
            }
            AdjustDashboardLayout();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void SetupSheetsLayout()
        {
            if (_panelCtrl1 != null) _panelCtrl1.Visible = false;
            if (_splitter1 != null) _splitter1.Visible = false;
            if (_panelCtrl2 != null) _panelCtrl2.Dock = DockStyle.Fill;

            EnsureBottomSheetsConfigured();
            BuildDashboardSheet();
            BuildReportSheet();
        }

        private void EnsureBottomSheetsConfigured()
        {
            if (_Rb1cn_xtraTabControl1 == null) return;

            _Rb1cn_xtraTabControl1.ShowTabHeader = DevExpress.Utils.DefaultBoolean.True;
            _Rb1cn_xtraTabControl1.HeaderLocation = TabHeaderLocation.Bottom;
            _Rb1cn_xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.False;
            _Rb1cn_xtraTabControl1.AppearancePage.Header.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            _Rb1cn_xtraTabControl1.AppearancePage.HeaderActive.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            if (_Rb1cn_xtraTabPage1 != null)
            {
                _Rb1cn_xtraTabPage1.Text = "Dashboard Rekapitulasi";
                _Rb1cn_xtraTabPage1.PageVisible = true;
                _Rb1cn_xtraTabPage1.PageEnabled = true;
                _Rb1cn_xtraTabPage1.ShowCloseButton = DevExpress.Utils.DefaultBoolean.False;

                if (!_Rb1cn_xtraTabControl1.TabPages.Contains(_Rb1cn_xtraTabPage1))
                {
                    _Rb1cn_xtraTabControl1.TabPages.Add(_Rb1cn_xtraTabPage1);
                }
            }

            if (_Rb1cn_xtraTabPage2 != null)
            {
                _Rb1cn_xtraTabPage2.Text = "Report Kualitas Buah";
                _Rb1cn_xtraTabPage2.PageVisible = true;
                _Rb1cn_xtraTabPage2.PageEnabled = true;
                _Rb1cn_xtraTabPage2.ShowCloseButton = DevExpress.Utils.DefaultBoolean.False;

                if (!_Rb1cn_xtraTabControl1.TabPages.Contains(_Rb1cn_xtraTabPage2))
                {
                    _Rb1cn_xtraTabControl1.TabPages.Add(_Rb1cn_xtraTabPage2);
                }
            }

            for (int i = 0; i < _Rb1cn_xtraTabControl1.TabPages.Count; i++)
            {
                var page = _Rb1cn_xtraTabControl1.TabPages[i];
                if (page != _Rb1cn_xtraTabPage1 && page != _Rb1cn_xtraTabPage2)
                {
                    page.PageVisible = false;
                }
            }

            _Rb1cn_xtraTabControl1.SelectedPageChanged -= XtraTabControl_SelectedPageChanged;
            _Rb1cn_xtraTabControl1.SelectedPageChanged += XtraTabControl_SelectedPageChanged;
        }

        private void XtraTabControl_SelectedPageChanged(object sender, TabPageChangedEventArgs e)
        {
            if (e.Page == _Rb1cn_xtraTabPage2)
            {
                ApplyFilter();
            }
            else if (e.Page == _Rb1cn_xtraTabPage1)
            {
                if (scrollDashboard != null)
                {
                    scrollDashboard.AutoScrollPosition = new Point(0, 0);
                }
                AdjustDashboardLayout();
                EnableDoubleBuffering(scrollDashboard);
            }
        }

        private void BuildDashboardSheet()
        {
            _Rb1cn_xtraTabPage1.Controls.Clear();

            scrollDashboard = new SmoothScrollableControl();
            scrollDashboard.Dock = DockStyle.Fill;
            scrollDashboard.AutoScroll = true;
            scrollDashboard.Appearance.BackColor = ColorCanvasBg;
            scrollDashboard.Appearance.Options.UseBackColor = true;
            scrollDashboard.VerticalScroll.SmallChange = 30;
            scrollDashboard.VerticalScroll.LargeChange = 90;

            pnlDashboardCanvas = new SmoothDashboardCanvas();
            pnlDashboardCanvas.BorderStyle = BorderStyles.NoBorder;
            pnlDashboardCanvas.Appearance.BackColor = ColorCanvasBg;
            pnlDashboardCanvas.Appearance.Options.UseBackColor = true;
            pnlDashboardCanvas.Location = new Point(0, 0);
            pnlDashboardCanvas.MouseEnter += (s, e) => scrollDashboard.Focus();

            pnlBannerHeader = new PanelControl();
            pnlBannerHeader.BorderStyle = BorderStyles.NoBorder;
            pnlBannerHeader.Height = 84;
            pnlBannerHeader.Appearance.BackColor = ColorCanvasBg;
            pnlBannerHeader.Appearance.Options.UseBackColor = true;
            pnlBannerHeader.Paint += (s, pe) =>
            {
                pe.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle rect = new Rectangle(0, 0, pnlBannerHeader.Width - 1, pnlBannerHeader.Height - 1);
                if (rect.Width <= 10 || rect.Height <= 10) return;

                using (GraphicsPath path = CreateRoundedRectangle(rect, 12))
                {
                    using (LinearGradientBrush br = new LinearGradientBrush(rect, ColorPrimary, ColorPrimaryDark, LinearGradientMode.Horizontal))
                    {
                        pe.Graphics.FillPath(br, path);
                    }
                    using (Pen p = new Pen(Color.FromArgb(80, 56, 189, 248), 1.5f))
                    {
                        pe.Graphics.DrawPath(p, path);
                    }
                }
            };

            PanelControl pnlBannerLeft = new PanelControl();
            pnlBannerLeft.Dock = DockStyle.Left;
            pnlBannerLeft.Width = 440;
            pnlBannerLeft.BorderStyle = BorderStyles.NoBorder;
            pnlBannerLeft.Appearance.BackColor = Color.Transparent;
            pnlBannerLeft.Appearance.Options.UseBackColor = true;

            lblBannerSubtitle = new LabelControl();
            lblBannerSubtitle.Text = "REKAPITULASI SENSUS & GRADING KUALITAS BUAH PANEN";
            lblBannerSubtitle.Appearance.Font = new Font("Segoe UI", 7.5F, FontStyle.Bold);
            lblBannerSubtitle.Appearance.ForeColor = Color.FromArgb(186, 230, 253);
            lblBannerSubtitle.Location = new Point(20, 14);
            pnlBannerLeft.Controls.Add(lblBannerSubtitle);

            lblBannerTitle = new LabelControl();
            lblBannerTitle.Text = "Laporan Panen TPH - SPECTRA (SPR1)";
            lblBannerTitle.Appearance.Font = new Font("Segoe UI", 15F, FontStyle.Bold);
            lblBannerTitle.Appearance.ForeColor = Color.White;
            lblBannerTitle.Location = new Point(20, 28);
            pnlBannerLeft.Controls.Add(lblBannerTitle);

            LabelControl lblBannerNote = new LabelControl();
            lblBannerNote.Text = "Unit: PT Kayan Plantation  |  Mode: Read-Only Report & Filter Kolom";
            lblBannerNote.Appearance.Font = new Font("Segoe UI", 7.5F);
            lblBannerNote.Appearance.ForeColor = Color.FromArgb(224, 242, 254);
            lblBannerNote.Location = new Point(20, 58);
            pnlBannerLeft.Controls.Add(lblBannerNote);

            pnlBannerHeader.Controls.Add(pnlBannerLeft);

            pnlBannerRightContainer = new PanelControl();
            pnlBannerRightContainer.Dock = DockStyle.Right;
            pnlBannerRightContainer.Width = 520;
            pnlBannerRightContainer.BorderStyle = BorderStyles.NoBorder;
            pnlBannerRightContainer.Appearance.BackColor = Color.Transparent;
            pnlBannerRightContainer.Appearance.Options.UseBackColor = true;

            PanelControl bPeriode = CreateBannerMiniCard("Periode Data", DateTime.Now.ToString("dd-MMM-yyyy"), ColorSlate800, 0, out lblBannerPeriodeVal);
            PanelControl bTotal = CreateBannerMiniCard("Total TPH", "0", ColorPrimary, 1, out lblBannerTphVal);
            PanelControl bMatang = CreateBannerMiniCard("Matang (Jjg)", "0", ColorMatang, 2, out lblBannerMatangVal);
            PanelControl bMentah = CreateBannerMiniCard("Mentah (Jjg)", "0", ColorMentah, 3, out lblBannerMentahVal);
            PanelControl bAfkir = CreateBannerMiniCard("Afkir/Busuk", "0", ColorAfkir, 4, out lblBannerAfkirVal);

            pnlBannerRightContainer.Controls.Add(bPeriode);
            pnlBannerRightContainer.Controls.Add(bTotal);
            pnlBannerRightContainer.Controls.Add(bMatang);
            pnlBannerRightContainer.Controls.Add(bMentah);
            pnlBannerRightContainer.Controls.Add(bAfkir);

            pnlBannerHeader.Controls.Add(pnlBannerRightContainer);
            pnlDashboardCanvas.Controls.Add(pnlBannerHeader);

            pnlSummaryContainer = new PanelControl();
            pnlSummaryContainer.BorderStyle = BorderStyles.NoBorder;
            pnlSummaryContainer.Appearance.BackColor = ColorCanvasBg;
            pnlSummaryContainer.Appearance.Options.UseBackColor = true;
            pnlSummaryContainer.Height = 114;

            LabelControl lblSummaryTitle = new LabelControl();
            lblSummaryTitle.Text = "Rekapitulasi Mutu Buah Panen (TBS)";
            lblSummaryTitle.Appearance.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            lblSummaryTitle.Appearance.ForeColor = ColorSlate800;
            lblSummaryTitle.Location = new Point(16, 6);
            pnlSummaryContainer.Controls.Add(lblSummaryTitle);

            pnlFilterBar = new PanelControl();
            pnlFilterBar.BorderStyle = BorderStyles.NoBorder;
            pnlFilterBar.Appearance.BackColor = ColorCanvasBg;
            pnlFilterBar.Appearance.Options.UseBackColor = true;
            pnlFilterBar.Height = 28;

            LabelControl lblFilterPrompt = new LabelControl();
            lblFilterPrompt.Text = "Periode:";
            lblFilterPrompt.Appearance.Font = new Font("Segoe UI", 8F, FontStyle.Bold);
            lblFilterPrompt.Appearance.ForeColor = ColorSlate600;
            lblFilterPrompt.Location = new Point(0, 6);
            pnlFilterBar.Controls.Add(lblFilterPrompt);

            filterButtons.Clear();
            var periodList = new (string key, string text, int w)[]
            {
                ("TODAY", "Hari Ini", 66),
                ("3DAYS", "3 Hari", 56),
                ("7DAYS", "7 Hari", 56),
                ("1MONTH", "Sebulan", 66),
                ("3MONTHS", "3 Bulan", 66),
                ("1YEAR", "Setahun", 66),
                ("ALL", "Semua", 56)
            };

            int curBtnX = 52;
            foreach (var item in periodList)
            {
                SimpleButton btn = CreateFilterPillButton(item.key, item.text, item.w);
                btn.Location = new Point(curBtnX, 2);
                pnlFilterBar.Controls.Add(btn);
                filterButtons[item.key] = btn;
                curBtnX += item.w + 4;
            }
            pnlFilterBar.Width = curBtnX;
            pnlSummaryContainer.Controls.Add(pnlFilterBar);

            UpdateFilterButtonsVisual();

            PanelControl cardTph = CreateKpiCard("TOTAL TPH DIPERIKSA", "0", "Titik Pengumpulan Hasil", ColorPrimary, 0, out lblKpiTotalTphVal);
            PanelControl cardJanjang = CreateKpiCard("TOTAL JANJANG TBS", "0", "Akumulasi Seluruh Buah", ColorAccent, 1, out lblKpiTotalJanjangVal);
            PanelControl cardMatang = CreateKpiCard("BUAH MATANG (IDEAL)", "0", "Standar Panen Optimal", ColorMatang, 2, out lblKpiMatangVal);
            PanelControl cardMentah = CreateKpiCard("BUAH MENTAH / MENGKAL", "0", "Kurang Fraksi Matang", ColorMentah, 3, out lblKpiMentahVal);
            PanelControl cardAfkir = CreateKpiCard("BUSUK / ABNORMAL / GAGANG", "0", "Perlu Evaluasi Mandor", ColorAfkir, 4, out lblKpiAfkirVal);

            pnlSummaryContainer.Controls.Add(cardTph);
            pnlSummaryContainer.Controls.Add(cardJanjang);
            pnlSummaryContainer.Controls.Add(cardMatang);
            pnlSummaryContainer.Controls.Add(cardMentah);
            pnlSummaryContainer.Controls.Add(cardAfkir);

            pnlDashboardCanvas.Controls.Add(pnlSummaryContainer);

            pnlSectionEstateHeader = CreateSectionHeader("Rekapitulasi per Kebun / Estate", "Klik kartu kebun untuk langsung membuka rincian laporan per kebun di Sheet Report");
            pnlDashboardCanvas.Controls.Add(pnlSectionEstateHeader);

            pnlRowEstates = new PanelControl();
            pnlRowEstates.BorderStyle = BorderStyles.NoBorder;
            pnlRowEstates.Appearance.BackColor = ColorCanvasBg;
            pnlRowEstates.Appearance.Options.UseBackColor = true;
            pnlDashboardCanvas.Controls.Add(pnlRowEstates);

            scrollDashboard.Controls.Add(pnlDashboardCanvas);
            _Rb1cn_xtraTabPage1.Controls.Add(scrollDashboard);

            BuildEstateCards();
            EnableDoubleBuffering(scrollDashboard);
            scrollDashboard.Resize += (s, e) => AdjustDashboardLayout();
        }

        private PanelControl CreateBannerMiniCard(string title, string initVal, Color valColor, int index, out LabelControl outVal)
        {
            PanelControl card = new PanelControl();
            card.BorderStyle = BorderStyles.NoBorder;
            card.Appearance.BackColor = Color.Transparent;
            card.Appearance.Options.UseBackColor = true;
            card.Tag = index;
            card.Size = new Size(96, 56);
            card.Location = new Point(index * 102 + 6, 14);

            card.Paint += (s, pe) =>
            {
                pe.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle r = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (GraphicsPath path = CreateRoundedRectangle(r, 8))
                {
                    using (SolidBrush br = new SolidBrush(Color.White))
                        pe.Graphics.FillPath(br, path);
                    using (Pen p = new Pen(Color.FromArgb(226, 232, 240), 1f))
                        pe.Graphics.DrawPath(p, path);
                }
            };

            LabelControl lblTitle = new LabelControl();
            lblTitle.Text = title;
            lblTitle.Appearance.Font = new Font("Segoe UI", 7F, FontStyle.Bold);
            lblTitle.Appearance.ForeColor = ColorSlate600;
            lblTitle.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            lblTitle.AutoSizeMode = LabelAutoSizeMode.None;
            lblTitle.Size = new Size(92, 14);
            lblTitle.Location = new Point(2, 7);
            card.Controls.Add(lblTitle);

            LabelControl lblVal = new LabelControl();
            lblVal.Text = initVal;
            lblVal.Appearance.Font = new Font("Segoe UI", index == 0 ? 8.5F : 12F, FontStyle.Bold);
            lblVal.Appearance.ForeColor = valColor;
            lblVal.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            lblVal.AutoSizeMode = LabelAutoSizeMode.None;
            lblVal.Size = new Size(92, 24);
            lblVal.Location = new Point(2, 24);
            card.Controls.Add(lblVal);

            outVal = lblVal;
            return card;
        }

        private PanelControl CreateSectionHeader(string title, string flowNote)
        {
            PanelControl pnl = new PanelControl();
            pnl.BorderStyle = BorderStyles.NoBorder;
            pnl.Appearance.BackColor = ColorCanvasBg;
            pnl.Appearance.Options.UseBackColor = true;
            pnl.Height = 26;

            LabelControl lblTitle = new LabelControl();
            lblTitle.Text = title;
            lblTitle.Appearance.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            lblTitle.Appearance.ForeColor = ColorSlate800;
            lblTitle.Location = new Point(16, 4);
            pnl.Controls.Add(lblTitle);

            LabelControl lblNote = new LabelControl();
            lblNote.Text = flowNote;
            lblNote.Appearance.Font = new Font("Segoe UI", 8F);
            lblNote.Appearance.ForeColor = ColorSlate600;
            lblNote.Location = new Point(250, 6);
            pnl.Controls.Add(lblNote);

            return pnl;
        }

        private PanelControl CreateKpiCard(string title, string initVal, string subtext, Color accentColor, int index, out LabelControl outVal)
        {
            PanelControl card = new PanelControl();
            card.BorderStyle = BorderStyles.NoBorder;
            card.Appearance.BackColor = ColorCanvasBg;
            card.Appearance.Options.UseBackColor = true;
            card.Tag = index;

            card.Paint += (s, pe) =>
            {
                pe.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle r = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (GraphicsPath path = CreateRoundedRectangle(r, 10))
                {
                    using (SolidBrush br = new SolidBrush(Color.White))
                        pe.Graphics.FillPath(br, path);
                    using (Pen p = new Pen(ColorBorder, 1.2f))
                        pe.Graphics.DrawPath(p, path);
                }

                Rectangle stripRect = new Rectangle(0, 8, 4, card.Height - 16);
                using (GraphicsPath stripPath = CreateRoundedRectangle(stripRect, 2))
                using (SolidBrush brStrip = new SolidBrush(accentColor))
                    pe.Graphics.FillPath(brStrip, stripPath);
            };

            LabelControl lblTitle = new LabelControl();
            lblTitle.Text = title;
            lblTitle.Appearance.Font = new Font("Segoe UI", 7.2F, FontStyle.Bold);
            lblTitle.Appearance.ForeColor = ColorSlate600;
            lblTitle.Location = new Point(14, 8);
            card.Controls.Add(lblTitle);

            LabelControl lblVal = new LabelControl();
            lblVal.Text = initVal;
            lblVal.Appearance.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            lblVal.Appearance.ForeColor = accentColor;
            lblVal.Location = new Point(14, 24);
            card.Controls.Add(lblVal);

            LabelControl lblSub = new LabelControl();
            lblSub.Text = subtext;
            lblSub.Appearance.Font = new Font("Segoe UI", 7.5F);
            lblSub.Appearance.ForeColor = ColorSlate400;
            lblSub.Location = new Point(14, 52);
            card.Controls.Add(lblSub);

            outVal = lblVal;
            return card;
        }

        private void BuildEstateCards()
        {
            pnlRowEstates.Controls.Clear();
            estateCardMap.Clear();

            foreach (var def in estateDefinitions)
            {
                EstateCardControl card = new EstateCardControl(def);
                card.CardClicked += (s, estCode) => OpenEstateInReport(estCode);
                pnlRowEstates.Controls.Add(card);
                estateCardMap[def.EstateCode] = card;
            }
        }

        private void BuildReportSheet()
        {
            _Rb1cn_xtraTabPage2.Controls.Clear();

            pnlReportPhotoPanel = new PanelControl();
            pnlReportPhotoPanel.Dock = DockStyle.Right;
            pnlReportPhotoPanel.Width = 380;
            pnlReportPhotoPanel.Visible = false;

            splitterReportPhoto = new SplitterControl();
            splitterReportPhoto.Dock = DockStyle.Right;
            splitterReportPhoto.Width = 5;
            splitterReportPhoto.Visible = false;

            tabRightPhoto = new XtraTabControl();
            tabRightPhoto.Dock = DockStyle.Fill;
            tabRightPhoto.HeaderLocation = TabHeaderLocation.Top;

            pagePhotoPreview = new XtraTabPage();
            pagePhotoPreview.Text = "Foto Kualitas Buah TPH";

            btnClosePhotoPanel = new SimpleButton();
            btnClosePhotoPanel.Text = "Tutup";
            btnClosePhotoPanel.Size = new Size(70, 24);
            btnClosePhotoPanel.Location = new Point(290, 10);
            btnClosePhotoPanel.Click += (s, ev) =>
            {
                pnlReportPhotoPanel.Visible = false;
                splitterReportPhoto.Visible = false;
            };
            pagePhotoPreview.Controls.Add(btnClosePhotoPanel);

            lblTransactionTitle = new LabelControl();
            lblTransactionTitle.Text = "-";
            lblTransactionTitle.Appearance.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            lblTransactionTitle.Appearance.ForeColor = ColorPrimary;
            lblTransactionTitle.Location = new Point(16, 12);
            lblTransactionTitle.MaximumSize = new Size(260, 25);
            pagePhotoPreview.Controls.Add(lblTransactionTitle);

            lblEstateAndBlockInfo = new LabelControl();
            lblEstateAndBlockInfo.Text = "Kebun: - | Afdeling: - | Blok: -";
            lblEstateAndBlockInfo.Appearance.Font = new Font("Segoe UI", 8.25F);
            lblEstateAndBlockInfo.Appearance.ForeColor = ColorSlate600;
            lblEstateAndBlockInfo.Location = new Point(16, 35);
            pagePhotoPreview.Controls.Add(lblEstateAndBlockInfo);

            picConditionViewer = new PictureEdit();
            picConditionViewer.Location = new Point(16, 56);
            picConditionViewer.Size = new Size(345, 235);
            picConditionViewer.Properties.SizeMode = PictureSizeMode.Zoom;
            picConditionViewer.Properties.NullText = "Tidak ada lampiran foto TBS";
            picConditionViewer.Properties.ShowMenu = true;
            picConditionViewer.Properties.Appearance.BackColor = Color.FromArgb(241, 245, 249);
            picConditionViewer.Properties.Appearance.Options.UseBackColor = true;
            picConditionViewer.BorderStyle = BorderStyles.Simple;
            pagePhotoPreview.Controls.Add(picConditionViewer);

            lblPhotoStatus = new LabelControl();
            lblPhotoStatus.Text = "Status: -";
            lblPhotoStatus.Appearance.Font = new Font("Segoe UI", 7.5F, FontStyle.Italic);
            lblPhotoStatus.Appearance.ForeColor = ColorSlate600;
            lblPhotoStatus.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            lblPhotoStatus.AutoSizeMode = LabelAutoSizeMode.Vertical;
            lblPhotoStatus.Location = new Point(16, 294);
            lblPhotoStatus.Size = new Size(345, 24);
            pagePhotoPreview.Controls.Add(lblPhotoStatus);

            lblQualitySummary = new LabelControl();
            lblQualitySummary.Text = "Total Janjang: -";
            lblQualitySummary.Appearance.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            lblQualitySummary.Appearance.ForeColor = ColorSlate800;
            lblQualitySummary.Location = new Point(16, 322);
            pagePhotoPreview.Controls.Add(lblQualitySummary);

            lblDetailsCount = new LabelControl();
            lblDetailsCount.Text = "Matang: - | Mentah: - | Busuk: -\nGagang Panjang: - | Abnormal: -";
            lblDetailsCount.Appearance.Font = new Font("Segoe UI", 8.5F);
            lblDetailsCount.Appearance.ForeColor = ColorSlate800;
            lblDetailsCount.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            lblDetailsCount.AutoSizeMode = LabelAutoSizeMode.Vertical;
            lblDetailsCount.Location = new Point(16, 344);
            lblDetailsCount.Size = new Size(345, 34);
            pagePhotoPreview.Controls.Add(lblDetailsCount);

            lblKraniInfo = new LabelControl();
            lblKraniInfo.Text = "Petugas Kerani: -";
            lblKraniInfo.Appearance.Font = new Font("Segoe UI", 8.25F);
            lblKraniInfo.Appearance.ForeColor = ColorPrimary;
            lblKraniInfo.Location = new Point(16, 384);
            pagePhotoPreview.Controls.Add(lblKraniInfo);

            lblSyncStatusBadge = new LabelControl();
            lblSyncStatusBadge.Text = "Status: SYNCED";
            lblSyncStatusBadge.Appearance.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            lblSyncStatusBadge.Appearance.ForeColor = ColorMatang;
            lblSyncStatusBadge.Location = new Point(16, 416);
            pagePhotoPreview.Controls.Add(lblSyncStatusBadge);

            lblTimestampText = new LabelControl();
            lblTimestampText.Text = "Waktu Upload: -";
            lblTimestampText.Appearance.Font = new Font("Segoe UI", 8F);
            lblTimestampText.Appearance.ForeColor = ColorSlate400;
            lblTimestampText.Location = new Point(16, 436);
            pagePhotoPreview.Controls.Add(lblTimestampText);

            tabRightPhoto.TabPages.Add(pagePhotoPreview);
            pnlReportPhotoPanel.Controls.Add(tabRightPhoto);

            _Rb1cn_xtraTabPage2.Controls.Add(pnlReportPhotoPanel);
            _Rb1cn_xtraTabPage2.Controls.Add(splitterReportPhoto);
            _Rb1cn_xtraTabPage2.Controls.Add(_gridControlSR1);

            pnlReportPhotoPanel.BringToFront();
            splitterReportPhoto.BringToFront();
            _gridControlSR1.Dock = DockStyle.Fill;
            _gridControlSR1.SendToBack();

            ConfigureGridColumnsReadOnlyWithFilter();

            _gridControlSR1.MouseDown -= _gridControlSR1_MouseDown;
            _gridControlSR1.MouseDown += _gridControlSR1_MouseDown;

            _gridViewSR1.RowCellClick -= _gridViewSR1_RowCellClick;
            _gridViewSR1.RowCellClick += _gridViewSR1_RowCellClick;

            _gridViewSR1.FocusedRowChanged -= _gridViewSR1_FocusedRowChanged;
            _gridViewSR1.FocusedRowChanged += _gridViewSR1_FocusedRowChanged;
        }

        private void _gridControlSR1_MouseDown(object sender, MouseEventArgs e)
        {
            BandedGridView view = _gridViewSR1 as BandedGridView;
            if (view == null) return;

            GridHitInfo hit = view.CalcHitInfo(e.Location);
            if (hit.InRowCell && hit.Column != null && hit.Column.FieldName == "PHOTO_PATH")
            {
                OpenPhotoConditionTab(hit.RowHandle);
            }
        }

        private void _gridViewSR1_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (e.Column != null && e.Column.FieldName == "PHOTO_PATH")
            {
                OpenPhotoConditionTab(e.RowHandle);
            }
        }

        private void _gridViewSR1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (pnlReportPhotoPanel != null && pnlReportPhotoPanel.Visible && e.FocusedRowHandle >= 0)
            {
                OpenPhotoConditionTab(e.FocusedRowHandle);
            }
        }

        private void OpenPhotoConditionTab(int rowHandle)
        {
            if (rowHandle < 0 || _gridViewSR1.RowCount == 0) return;

            string idTrx = _gridViewSR1.GetRowCellValue(rowHandle, "ID")?.ToString() ?? "-";
            string estate = _gridViewSR1.GetRowCellValue(rowHandle, "ESTATE_CODE")?.ToString() ?? "-";
            string afdeling = _gridViewSR1.GetRowCellValue(rowHandle, "AFDELING_CODE")?.ToString() ?? "-";
            string block = _gridViewSR1.GetRowCellValue(rowHandle, "BLOCK_ID")?.ToString() ?? "-";
            string tph = _gridViewSR1.GetRowCellValue(rowHandle, "TPH_NUMBER")?.ToString() ?? "-";
            string nik = _gridViewSR1.GetRowCellValue(rowHandle, "KRANI_NIK")?.ToString() ?? "-";
            string kraniName = _gridViewSR1.GetRowCellValue(rowHandle, "KRANI_NAME")?.ToString() ?? "-";
            string harvestDate = _gridViewSR1.GetRowCellValue(rowHandle, "HARVEST_DATE")?.ToString() ?? "-";
            string matang = _gridViewSR1.GetRowCellValue(rowHandle, "MATANG_COUNT")?.ToString() ?? "0";
            string mentah = _gridViewSR1.GetRowCellValue(rowHandle, "MENTAH_COUNT")?.ToString() ?? "0";
            string busuk = _gridViewSR1.GetRowCellValue(rowHandle, "BUAH_BUSUK_COUNT")?.ToString() ?? "0";
            string tangkai = _gridViewSR1.GetRowCellValue(rowHandle, "TANGKAI_PANJANG_COUNT")?.ToString() ?? "0";
            string abnormal = _gridViewSR1.GetRowCellValue(rowHandle, "ABNORMAL_COUNT")?.ToString() ?? "0";
            string total = _gridViewSR1.GetRowCellValue(rowHandle, "TOTAL_BUAH")?.ToString() ?? "0";
            string photoPath = _gridViewSR1.GetRowCellValue(rowHandle, "PHOTO_PATH")?.ToString() ?? "";
            string syncStat = _gridViewSR1.GetRowCellValue(rowHandle, "SYNC_STATUS")?.ToString() ?? "SYNCED";
            string createdAt = _gridViewSR1.GetRowCellValue(rowHandle, "CREATED_AT")?.ToString() ?? "-";

            lblTransactionTitle.Text = $"ID #{idTrx} - TPH {tph}";
            lblEstateAndBlockInfo.Text = $"Kebun: {estate} | Afdeling: {afdeling} | Blok: {block}";
            lblQualitySummary.Text = $"Total Janjang: {total} TBS";
            lblDetailsCount.Text = $"Matang: {matang} | Mentah: {mentah} | Busuk: {busuk}\nGagang Panjang: {tangkai} | Abnormal: {abnormal}";
            lblKraniInfo.Text = $"Petugas Kerani: {kraniName} ({nik})\nTgl Panen: {harvestDate}";
            lblSyncStatusBadge.Text = $"Status: {syncStat.ToUpper()}";
            lblTimestampText.Text = $"Upload: {createdAt}";

            lblSyncStatusBadge.Appearance.ForeColor = syncStat.Equals("SYNCED", StringComparison.OrdinalIgnoreCase)
                ? ColorMatang
                : ColorMentah;

            picConditionViewer.Image = LoadPhotoBitmap(photoPath, out string sourceInfo);
            lblPhotoStatus.Text = sourceInfo;
            lblPhotoStatus.Appearance.ForeColor = picConditionViewer.Image != null ? ColorMatang : ColorAfkir;

            pnlReportPhotoPanel.Visible = true;
            splitterReportPhoto.Visible = true;
        }

        private void ConfigureGridColumnsReadOnlyWithFilter()
        {
            BandedGridView view = _gridViewSR1 as BandedGridView;
            if (view == null) return;

            view.BeginUpdate();
            view.Bands.Clear();
            view.Columns.Clear();

            view.OptionsView.ShowBands = true;
            view.OptionsView.ShowGroupPanel = false;
            view.OptionsView.ShowIndicator = true;
            view.OptionsView.ShowFooter = true;
            view.IndicatorWidth = 45;
            view.RowHeight = 26;
            view.ColumnPanelRowHeight = 38;
            view.BandPanelRowHeight = 24;

            // 1. KUNCI READ-ONLY PENUH (USER HANYA BISA READ & TIDAK BISA EDIT)
            view.OptionsBehavior.Editable = false;
            view.OptionsBehavior.ReadOnly = true;
            view.OptionsSelection.MultiSelect = true;
            view.ShowingEditor -= _gridViewSR1_ShowingEditor;
            view.ShowingEditor += _gridViewSR1_ShowingEditor;

            // 2. AKTIFKAN FILTER EXCEL + AUTO FILTER ROW DI SETIAP KOLOM
            view.OptionsView.ShowAutoFilterRow = true;
            view.OptionsFilter.AllowFilterEditor = true;
            view.OptionsFilter.DefaultFilterEditorView = FilterEditorViewMode.VisualAndText;
            view.OptionsCustomization.AllowFilter = true;
            view.OptionsCustomization.AllowSort = true;

            GridBand bandA = view.Bands.AddBand("A");
            AddReportColumn(view, bandA, "ID", "ID Trx", 65, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandB = view.Bands.AddBand("B");
            AddReportColumn(view, bandB, "ESTATE_CODE", "Kebun", 75, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandC = view.Bands.AddBand("C");
            AddReportColumn(view, bandC, "AFDELING_CODE", "Afdeling", 85, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandD = view.Bands.AddBand("D");
            AddReportColumn(view, bandD, "BLOCK_ID", "Blok", 80, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandE = view.Bands.AddBand("E");
            AddReportColumn(view, bandE, "TPH_NUMBER", "No. TPH", 85, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandF = view.Bands.AddBand("F");
            AddReportColumn(view, bandF, "KRANI_NIK", "NIK Kerani", 100, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandG = view.Bands.AddBand("G");
            AddReportColumn(view, bandG, "KRANI_NAME", "Nama Kerani Panen", 165, DevExpress.Utils.HorzAlignment.Near);

            GridBand bandH = view.Bands.AddBand("H");
            BandedGridColumn colDate = AddReportColumn(view, bandH, "HARVEST_DATE", "Tgl Panen", 100, DevExpress.Utils.HorzAlignment.Center);
            colDate.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            colDate.DisplayFormat.FormatString = "dd-MMM-yyyy";

            GridBand bandI = view.Bands.AddBand("I");
            BandedGridColumn colMatang = AddReportColumn(view, bandI, "MATANG_COUNT", "Buah Matang", 85, DevExpress.Utils.HorzAlignment.Center);
            colMatang.SummaryItem.SummaryType = SummaryItemType.Sum;
            colMatang.SummaryItem.DisplayFormat = "{0:#,##0}";

            GridBand bandJ = view.Bands.AddBand("J");
            BandedGridColumn colMentah = AddReportColumn(view, bandJ, "MENTAH_COUNT", "Buah Mentah", 85, DevExpress.Utils.HorzAlignment.Center);
            colMentah.SummaryItem.SummaryType = SummaryItemType.Sum;
            colMentah.SummaryItem.DisplayFormat = "{0:#,##0}";

            GridBand bandK = view.Bands.AddBand("K");
            BandedGridColumn colBusuk = AddReportColumn(view, bandK, "BUAH_BUSUK_COUNT", "Buah Busuk", 85, DevExpress.Utils.HorzAlignment.Center);
            colBusuk.SummaryItem.SummaryType = SummaryItemType.Sum;
            colBusuk.SummaryItem.DisplayFormat = "{0:#,##0}";

            GridBand bandL = view.Bands.AddBand("L");
            BandedGridColumn colTangkai = AddReportColumn(view, bandL, "TANGKAI_PANJANG_COUNT", "Tangkai Panjang", 95, DevExpress.Utils.HorzAlignment.Center);
            colTangkai.SummaryItem.SummaryType = SummaryItemType.Sum;
            colTangkai.SummaryItem.DisplayFormat = "{0:#,##0}";

            GridBand bandM = view.Bands.AddBand("M");
            BandedGridColumn colAbnormal = AddReportColumn(view, bandM, "ABNORMAL_COUNT", "Buah Abnormal", 90, DevExpress.Utils.HorzAlignment.Center);
            colAbnormal.SummaryItem.SummaryType = SummaryItemType.Sum;
            colAbnormal.SummaryItem.DisplayFormat = "{0:#,##0}";

            GridBand bandN = view.Bands.AddBand("N");
            BandedGridColumn colTotal = AddReportColumn(view, bandN, "TOTAL_BUAH", "Total Janjang", 95, DevExpress.Utils.HorzAlignment.Center);
            colTotal.AppearanceCell.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            colTotal.SummaryItem.SummaryType = SummaryItemType.Sum;
            colTotal.SummaryItem.DisplayFormat = "{0:#,##0}";

            GridBand bandO = view.Bands.AddBand("O");
            BandedGridColumn colPhoto = AddReportColumn(view, bandO, "PHOTO_PATH", "Foto TBS\r\n(Klik Preview)", 160, DevExpress.Utils.HorzAlignment.Center);
            colPhoto.AppearanceCell.ForeColor = Color.FromArgb(41, 128, 185);
            colPhoto.AppearanceCell.Font = new Font("Segoe UI", 8.25F, FontStyle.Underline);

            GridBand bandP = view.Bands.AddBand("P");
            AddReportColumn(view, bandP, "SYNC_STATUS", "Status Sinkron", 95, DevExpress.Utils.HorzAlignment.Center);

            GridBand bandQ = view.Bands.AddBand("Q");
            BandedGridColumn colCreated = AddReportColumn(view, bandQ, "CREATED_AT", "Waktu Upload", 135, DevExpress.Utils.HorzAlignment.Center);
            colCreated.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            colCreated.DisplayFormat.FormatString = "dd-MMM-yyyy HH:mm";

            view.RowCellStyle += (s, e) =>
            {
                if (e.Column.FieldName == "SYNC_STATUS")
                {
                    string val = e.CellValue?.ToString() ?? "";
                    if (val.Equals("SYNCED", StringComparison.OrdinalIgnoreCase))
                    {
                        e.Appearance.BackColor = ColorMatangBg;
                        e.Appearance.ForeColor = ColorMatang;
                        e.Appearance.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
                    }
                    else
                    {
                        e.Appearance.BackColor = ColorMentahBg;
                        e.Appearance.ForeColor = ColorMentah;
                        e.Appearance.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
                    }
                }
                else if (e.Column.FieldName == "TOTAL_BUAH")
                {
                    e.Appearance.ForeColor = ColorPrimary;
                    e.Appearance.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
                }
            };

            view.EndUpdate();
        }

        private BandedGridColumn AddReportColumn(BandedGridView view, GridBand band, string fieldName, string caption, int width, DevExpress.Utils.HorzAlignment align)
        {
            BandedGridColumn col = view.Columns.AddField(fieldName);
            col.Caption = caption;
            col.Width = width;
            col.Visible = true;

            col.OptionsColumn.AllowEdit = false;
            col.OptionsColumn.ReadOnly = true;

            col.OptionsFilter.AllowFilter = true;
            col.OptionsFilter.AllowAutoFilter = true;
            col.OptionsFilter.FilterPopupMode = FilterPopupMode.Excel;
            col.OptionsFilter.AutoFilterCondition = AutoFilterCondition.Contains;

            col.AppearanceCell.TextOptions.HAlignment = align;
            col.AppearanceHeader.Options.UseTextOptions = true;
            col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            col.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            col.AppearanceHeader.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            band.Columns.Add(col);
            return col;
        }

        private void _gridViewSR1_ShowingEditor(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void AdjustDashboardLayout()
        {
            if (scrollDashboard == null || pnlDashboardCanvas == null || pnlRowEstates == null) return;

            int clientWidth = scrollDashboard.ClientSize.Width;
            if (clientWidth <= 100) return;

            Point currentScroll = new Point(-scrollDashboard.AutoScrollPosition.X, -scrollDashboard.AutoScrollPosition.Y);
            scrollDashboard.AutoScrollPosition = new Point(0, 0);

            int leftMargin = 16;
            int spacing = 14;
            int columns = 4;

            int usableWidth = Math.Max(clientWidth - (leftMargin * 2), 960);
            int totalGaps = (columns - 1) * spacing;
            int cardWidth = (usableWidth - totalGaps) / columns;
            if (cardWidth < 210) cardWidth = 210;

            int rowWidth = (columns * cardWidth) + totalGaps;
            int cardHeight = 250;

            pnlBannerHeader.Location = new Point(leftMargin, 12);
            pnlBannerHeader.Size = new Size(rowWidth, 84);

            pnlSummaryContainer.Location = new Point(leftMargin, 104);
            pnlSummaryContainer.Size = new Size(rowWidth, 114);

            if (pnlFilterBar != null)
            {
                pnlFilterBar.Location = new Point(Math.Max(rowWidth - pnlFilterBar.Width, 230), 2);
            }

            int kpiCardWidth = (rowWidth - (4 * 12)) / 5;
            for (int i = 0; i < 5; i++)
            {
                foreach (Control c in pnlSummaryContainer.Controls)
                {
                    if (c is PanelControl pc && pc.Tag is int tag && tag == i)
                    {
                        pc.Location = new Point(i * (kpiCardWidth + 12), 36);
                        pc.Size = new Size(kpiCardWidth, 72);
                    }
                }
            }

            pnlSectionEstateHeader.Location = new Point(leftMargin, 228);
            pnlSectionEstateHeader.Size = new Size(rowWidth, 26);

            pnlRowEstates.Location = new Point(leftMargin, 256);
            pnlRowEstates.Size = new Size(rowWidth, cardHeight);

            foreach (var def in estateDefinitions)
            {
                if (estateCardMap.TryGetValue(def.EstateCode, out EstateCardControl card))
                {
                    int posX = def.Index * (cardWidth + spacing);
                    card.Location = new Point(posX, 0);
                    card.UpdateLayout(cardWidth, cardHeight);
                }
            }

            int canvasWidth = Math.Max(rowWidth + (leftMargin * 2), clientWidth);
            int canvasHeight = 256 + cardHeight + 28;

            pnlDashboardCanvas.Location = new Point(0, 0);
            pnlDashboardCanvas.Size = new Size(canvasWidth, canvasHeight);
            scrollDashboard.AutoScrollMinSize = new Size(canvasWidth, canvasHeight);

            if (currentScroll.Y > 0 || currentScroll.X > 0)
            {
                scrollDashboard.AutoScrollPosition = currentScroll;
            }
        }

        private SimpleButton CreateFilterPillButton(string key, string text, int width)
        {
            SimpleButton btn = new SimpleButton();
            btn.Text = text;
            btn.Tag = key;
            btn.Size = new Size(width, 24);
            btn.Cursor = Cursors.Hand;
            btn.LookAndFeel.UseDefaultLookAndFeel = false;
            btn.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
            btn.Appearance.Options.UseBackColor = true;
            btn.Appearance.Options.UseForeColor = true;
            btn.Appearance.Options.UseBorderColor = true;
            btn.Appearance.Options.UseFont = true;
            btn.Click += (s, e) => SetPeriodFilter(key);
            return btn;
        }

        private void SetPeriodFilter(string periodKey)
        {
            activePeriodFilter = periodKey;
            UpdateFilterButtonsVisual();
            UpdateDashboardSummary();
            ApplyFilter();
        }

        private void UpdateFilterButtonsVisual()
        {
            foreach (var kvp in filterButtons)
            {
                SimpleButton btn = kvp.Value;
                bool isSelected = (kvp.Key == activePeriodFilter);

                if (isSelected)
                {
                    btn.Appearance.BackColor = ColorPrimary;
                    btn.Appearance.ForeColor = Color.White;
                    btn.Appearance.BorderColor = ColorPrimaryDark;
                    btn.Appearance.Font = new Font("Segoe UI", 7.5F, FontStyle.Bold);
                }
                else
                {
                    btn.Appearance.BackColor = Color.White;
                    btn.Appearance.ForeColor = ColorSlate800;
                    btn.Appearance.BorderColor = ColorBorder;
                    btn.Appearance.Font = new Font("Segoe UI", 7.5F, FontStyle.Regular);
                }
                btn.Invalidate();
            }
        }

        private void GetActiveDateRange(out DateTime? startDate, out DateTime? endDate, out string periodLabel)
        {
            DateTime today = DateTime.Today;

            switch (activePeriodFilter)
            {
                case "TODAY":
                    startDate = today;
                    endDate = today.AddDays(1).AddTicks(-1);
                    periodLabel = today.ToString("dd-MMM-yyyy");
                    break;
                case "3DAYS":
                    startDate = today.AddDays(-2);
                    endDate = today.AddDays(1).AddTicks(-1);
                    periodLabel = $"{startDate.Value:dd/MM} - {today:dd/MM}";
                    break;
                case "7DAYS":
                    startDate = today.AddDays(-6);
                    endDate = today.AddDays(1).AddTicks(-1);
                    periodLabel = $"{startDate.Value:dd/MM} - {today:dd/MM}";
                    break;
                case "1MONTH":
                    startDate = today.AddMonths(-1);
                    endDate = today.AddDays(1).AddTicks(-1);
                    periodLabel = $"{startDate.Value:dd/MM} - {today:dd/MM}";
                    break;
                case "3MONTHS":
                    startDate = today.AddMonths(-3);
                    endDate = today.AddDays(1).AddTicks(-1);
                    periodLabel = $"{startDate.Value:dd/MM} - {today:dd/MM}";
                    break;
                case "1YEAR":
                    startDate = today.AddYears(-1);
                    endDate = today.AddDays(1).AddTicks(-1);
                    periodLabel = $"{startDate.Value:dd/MM/yy} - {today:dd/MM/yy}";
                    break;
                case "RIBBON_PERIOD":
                    if (int.TryParse(_5Connect_CBoxPrdY?.EditValue?.ToString(), out int yVal))
                    {
                        string mStr = _5Connect_CBoxPrdM?.EditValue?.ToString() ?? "";
                        if (string.IsNullOrEmpty(mStr) || mStr.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                        {
                            startDate = new DateTime(yVal, 1, 1, 0, 0, 0);
                            endDate = new DateTime(yVal, 12, 31, 23, 59, 59);
                            periodLabel = $"Tahun {yVal}";
                        }
                        else if (int.TryParse(mStr, out int mVal) && mVal >= 1 && mVal <= 12)
                        {
                            startDate = new DateTime(yVal, mVal, 1, 0, 0, 0);
                            endDate = startDate.Value.AddMonths(1).AddTicks(-1);
                            periodLabel = $"{startDate.Value:MMM-yyyy}";
                        }
                        else
                        {
                            startDate = null; endDate = null; periodLabel = "Semua";
                        }
                    }
                    else
                    {
                        startDate = null; endDate = null; periodLabel = "Semua";
                    }
                    break;
                case "ALL":
                default:
                    startDate = null;
                    endDate = null;
                    periodLabel = "Semua";
                    break;
            }
        }

        private void InitRibbonActionButtons()
        {
            RibbonControl rb = this.Ribbon;
            if (rb == null)
            {
                foreach (Control c in this.Controls)
                {
                    if (c is RibbonControl rc) { rb = rc; break; }
                }
            }

            if (rb == null) return;

            foreach (RibbonPage p in rb.Pages)
            {
                p.Visible = true;
                foreach (RibbonPageGroup grp in p.Groups)
                {
                    grp.Visible = true;
                }
            }

            if (_6Change_btnUnRgw != null) _6Change_btnUnRgw.Enabled = false;
            if (_6Change_btnRgw != null) _6Change_btnRgw.Enabled = false;

            if (_5Connect_BUnitDB != null)
            {
                RepositoryItemComboBox repoUnit = new RepositoryItemComboBox();
                repoUnit.Items.Add("ESTATE");
                repoUnit.TextEditStyle = TextEditStyles.DisableTextEditor;
                _5Connect_BUnitDB.Edit = repoUnit;
                _5Connect_BUnitDB.EditValue = "ESTATE";
                _5Connect_BUnitDB.Enabled = true;
            }

            if (_5Connect_btnRefreshAll != null)
            {
                _5Connect_btnRefreshAll.ItemClick -= _5Connect_btnRefresh_ItemClick;
                _5Connect_btnRefreshAll.ItemClick += _5Connect_btnRefresh_ItemClick;
            }
            if (_5Connect_btnRefresh != null)
            {
                _5Connect_btnRefresh.ItemClick -= _5Connect_btnRefresh_ItemClick;
                _5Connect_btnRefresh.ItemClick += _5Connect_btnRefresh_ItemClick;
            }

            InitPeriodDropdowns("dbo.TPH_TRANSACTION");

            btnRibbonResetFilter = new BarButtonItem(rb.Manager, "Semua\nKebun");
            btnRibbonResetFilter.RibbonStyle = RibbonItemStyles.Large | RibbonItemStyles.SmallWithText;
            btnRibbonResetFilter.LargeGlyph = CreateResetFilterRibbonIcon(32);
            btnRibbonResetFilter.Glyph = CreateResetFilterRibbonIcon(16);
            btnRibbonResetFilter.ItemClick += (s, ev) => FilterGridByEstate("");

            btnRibbonClearColFilter = new BarButtonItem(rb.Manager, "Reset Filter\nKolom");
            btnRibbonClearColFilter.RibbonStyle = RibbonItemStyles.Large | RibbonItemStyles.SmallWithText;
            btnRibbonClearColFilter.LargeGlyph = CreateClearFilterRibbonIcon(32);
            btnRibbonClearColFilter.Glyph = CreateClearFilterRibbonIcon(16);
            btnRibbonClearColFilter.ItemClick += (s, ev) =>
            {
                _gridViewSR1.ClearColumnsFilter();
                FilterGridByEstate("");
            };

            btnRibbonRefreshData = new BarButtonItem(rb.Manager, "Refresh\nLaporan");
            btnRibbonRefreshData.RibbonStyle = RibbonItemStyles.Large | RibbonItemStyles.SmallWithText;
            btnRibbonRefreshData.LargeGlyph = CreateRefreshRibbonIcon(32);
            btnRibbonRefreshData.Glyph = CreateRefreshRibbonIcon(16);
            btnRibbonRefreshData.ItemClick += (s, ev) => Syb3DBload(s, "SPECTRA_DB", _tblConsolBA);

            RibbonPage homePage = rb.Pages.Cast<RibbonPage>()
                .FirstOrDefault(p => p.Text.IndexOf("Home", StringComparison.OrdinalIgnoreCase) >= 0
                                  || p.Text.IndexOf("Beranda", StringComparison.OrdinalIgnoreCase) >= 0)
                ?? (rb.Pages.Count > 0 ? rb.Pages[0] : null);

            if (homePage != null)
            {
                RibbonPageGroup grpHome = new RibbonPageGroup("SPECTRA Report");
                grpHome.CaptionButtonVisible = DevExpress.Utils.DefaultBoolean.False;
                grpHome.ItemLinks.Add(btnRibbonResetFilter);
                grpHome.ItemLinks.Add(btnRibbonClearColFilter);
                grpHome.ItemLinks.Add(btnRibbonRefreshData);
                homePage.Groups.Add(grpHome);
            }
        }

        private Image CreateResetFilterRibbonIcon(int size)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                int pad = size >= 32 ? 4 : 2;
                int usable = size - (pad * 2);
                int cell = (usable - 4) / 3;
                using (SolidBrush br = new SolidBrush(ColorAccent))
                {
                    for (int r = 0; r < 3; r++)
                        for (int c = 0; c < 3; c++)
                            g.FillRectangle(br, pad + c * (cell + 2), pad + r * (cell + 2), cell, cell);
                }
            }
            return bmp;
        }

        private Image CreateClearFilterRibbonIcon(int size)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                int pad = size >= 32 ? 5 : 2;
                using (Pen pen = new Pen(ColorAfkir, size >= 32 ? 2.5f : 1.5f))
                {
                    g.DrawLine(pen, pad, pad, size - pad, size - pad);
                    g.DrawLine(pen, size - pad, pad, pad, size - pad);
                }
            }
            return bmp;
        }

        private Image CreateRefreshRibbonIcon(int size)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                int pad = size >= 32 ? 4 : 2;
                using (Pen pen = new Pen(ColorPrimary, size >= 32 ? 2.5f : 1.5f))
                {
                    g.DrawArc(pen, pad, pad, size - (pad * 2), size - (pad * 2), 45, 270);
                }
            }
            return bmp;
        }

        private void OpenEstateInReport(string estateCode)
        {
            FilterGridByEstate(estateCode);

            if (_Rb1cn_xtraTabControl1 != null && _Rb1cn_xtraTabPage2 != null)
            {
                _Rb1cn_xtraTabPage2.PageVisible = true;
                _Rb1cn_xtraTabControl1.SelectedTabPage = _Rb1cn_xtraTabPage2;
            }
        }

        private void FilterGridByEstate(string estateCode)
        {
            activeEstateFilter = estateCode;
            foreach (var kvp in estateCardMap)
            {
                kvp.Value.IsSelected = (kvp.Key == estateCode);
            }
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_dvSr1 == null) return;

            DateTime? startDate, endDate;
            string periodLabel;
            GetActiveDateRange(out startDate, out endDate, out periodLabel);

            List<string> filterList = new List<string>();

            if (!string.IsNullOrEmpty(activeEstateFilter))
            {
                filterList.Add($"ESTATE_CODE = '{activeEstateFilter}'");
            }

            if (startDate.HasValue)
            {
                filterList.Add($"HARVEST_DATE >= #{startDate.Value.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)}#");
            }

            if (endDate.HasValue)
            {
                filterList.Add($"HARVEST_DATE <= #{endDate.Value.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture)}#");
            }

            _dvSr1.RowFilter = filterList.Count > 0 ? string.Join(" AND ", filterList) : "";

            _gridControlSR1.RefreshDataSource();
            _gridViewSR1.RefreshData();
        }

        private void Syb3DBload(object sender, string vdbName, DataTable tblConsol1)
        {
            _Syb3InitPub1._mySyb3Layout.Syb3marqBarStart(sender, _AppParam3, _gridViewSR1, _Rb1cn_barProgres1, _Rb1cn_barZoom1, _Rb1cn_barStatSheet, _Rb1cn_barProgres1_repoMarque, this, _Syb3InitPub1, _Syb3InitLoc1);

            string targetDatabase = "SPECTRA_DB";
            _Syb3InitLoc1[_Syb3InitPub1.vSwapDB1] = targetDatabase;

            try
            {
                var dsTemp = _Syb3InitPub1._mySyb3Db.Syb3IDDict1(targetDatabase, false, _Syb3InitPub1, _Syb3InitLoc1);
                var propConn = dsTemp?.GetType().GetProperty("ConnectionString");
                activeConnString = propConn != null ? propConn.GetValue(dsTemp) as string : string.Empty;
            }
            catch { }

            if (string.IsNullOrEmpty(activeConnString))
            {
                try
                {
                    var dsActive = _Syb3InitPub1._mySyb3Db.Syb3IDDict1("", false, _Syb3InitPub1, _Syb3InitLoc1);
                    var propActive = dsActive?.GetType().GetProperty("ConnectionString");
                    string activeConn = propActive != null ? propActive.GetValue(dsActive) as string : string.Empty;

                    if (!string.IsNullOrEmpty(activeConn))
                    {
                        SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder(activeConn);
                        csb.InitialCatalog = targetDatabase;
                        activeConnString = csb.ConnectionString;
                    }
                }
                catch { }
            }

            if (string.IsNullOrEmpty(activeConnString))
            {
                activeConnString = "Data Source=localhost;Initial Catalog=SPECTRA_DB;Integrated Security=True;TrustServerCertificate=True;";
            }

            string sqlQuery = @"
                SELECT 
                    T.ID, 
                    ISNULL(T.ESTATE_CODE, '') AS ESTATE_CODE, 
                    ISNULL(T.AFDELING_CODE, '') AS AFDELING_CODE, 
                    ISNULL(T.BLOCK_ID, '') AS BLOCK_ID, 
                    ISNULL(T.TPH_NUMBER, '') AS TPH_NUMBER, 
                    ISNULL(T.KRANI_NIK, '') AS KRANI_NIK, 
                    ISNULL(E.FULL_NAME, '-') AS KRANI_NAME,
                    T.HARVEST_DATE, 
                    ISNULL(T.MATANG_COUNT, 0) AS MATANG_COUNT, 
                    ISNULL(T.MENTAH_COUNT, 0) AS MENTAH_COUNT, 
                    ISNULL(T.BUAH_BUSUK_COUNT, 0) AS BUAH_BUSUK_COUNT, 
                    ISNULL(T.TANGKAI_PANJANG_COUNT, 0) AS TANGKAI_PANJANG_COUNT, 
                    ISNULL(T.ABNORMAL_COUNT, 0) AS ABNORMAL_COUNT, 
                    ISNULL(T.TOTAL_BUAH, 0) AS TOTAL_BUAH, 
                    ISNULL(T.PHOTO_PATH, '') AS PHOTO_PATH, 
                    ISNULL(T.SYNC_STATUS, 'SYNCED') AS SYNC_STATUS, 
                    T.CREATED_AT 
                FROM dbo.TPH_TRANSACTION T WITH (NOLOCK)
                LEFT JOIN dbo.MD_EMPLOYEE E WITH (NOLOCK) ON LTRIM(RTRIM(T.KRANI_NIK)) = LTRIM(RTRIM(E.NIK))
                ORDER BY T.ID DESC";

            try
            {
                dtReport = new DataTable();
                using (SqlConnection conn = new SqlConnection(activeConnString))
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(sqlQuery, conn))
                    {
                        da.Fill(dtReport);
                    }
                }

                _dvSr1 = dtReport.DefaultView;
                _gridControlSR1.DataSource = _dvSr1;

                UpdateDashboardSummary();
                ApplyFilter();

                _Syb3InitPub1._mySyb3Layout.Syb3RefreshStatusBar2(sender, _Rb1cn_barStatSheet, _Rb1cn_barStatDB, _Rb1cn_barStatRec,
                    _dvSr1.Count, _Rb1cn_barStatPt, _Syb3InitLoc1, _Syb3InitPub1);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show("Gagal memuat laporan dari database SPECTRA_DB: " + ex.Message, "mySAP365 - SPR1", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            _Syb3InitPub1._mySyb3Layout.Syb3marqBarStop(sender, _gridViewSR1, _Rb1cn_barProgres1, _Rb1cn_barZoom1, _Rb1cn_barStatSheet, _Rb1cn_barProgres1_repo, this, _Syb3InitPub1, _Syb3InitLoc1);
        }

        private void UpdateDashboardSummary()
        {
            if (dtReport == null) return;

            DateTime? startDate, endDate;
            string periodLabel;
            GetActiveDateRange(out startDate, out endDate, out periodLabel);

            IEnumerable<DataRow> filteredRows = dtReport.AsEnumerable();
            if (startDate.HasValue)
            {
                filteredRows = filteredRows.Where(r => r["HARVEST_DATE"] != DBNull.Value && Convert.ToDateTime(r["HARVEST_DATE"]).Date >= startDate.Value.Date);
            }
            if (endDate.HasValue)
            {
                filteredRows = filteredRows.Where(r => r["HARVEST_DATE"] != DBNull.Value && Convert.ToDateTime(r["HARVEST_DATE"]).Date <= endDate.Value.Date);
            }

            var rowsList = filteredRows.ToList();

            int totalTph = rowsList.Count;
            int totalJanjang = rowsList.Sum(r => Convert.ToInt32(r["TOTAL_BUAH"] ?? 0));
            int totalMatang = rowsList.Sum(r => Convert.ToInt32(r["MATANG_COUNT"] ?? 0));
            int totalMentah = rowsList.Sum(r => Convert.ToInt32(r["MENTAH_COUNT"] ?? 0));
            int totalAfkir = rowsList.Sum(r =>
                Convert.ToInt32(r["BUAH_BUSUK_COUNT"] ?? 0) +
                Convert.ToInt32(r["TANGKAI_PANJANG_COUNT"] ?? 0) +
                Convert.ToInt32(r["ABNORMAL_COUNT"] ?? 0));

            foreach (var kvp in estateCardMap)
            {
                string estCode = kvp.Key;
                EstateCardControl card = kvp.Value;

                var estRows = rowsList
                    .Where(r => string.Equals(r["ESTATE_CODE"]?.ToString(), estCode, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                int eTph = estRows.Count;
                int eTotal = estRows.Sum(r => Convert.ToInt32(r["TOTAL_BUAH"] ?? 0));
                int eMatang = estRows.Sum(r => Convert.ToInt32(r["MATANG_COUNT"] ?? 0));
                int eMentah = estRows.Sum(r => Convert.ToInt32(r["MENTAH_COUNT"] ?? 0));
                int eAfkir = estRows.Sum(r =>
                    Convert.ToInt32(r["BUAH_BUSUK_COUNT"] ?? 0) +
                    Convert.ToInt32(r["TANGKAI_PANJANG_COUNT"] ?? 0) +
                    Convert.ToInt32(r["ABNORMAL_COUNT"] ?? 0));

                card.UpdateSummary(eTph, eTotal, eMatang, eMentah, eAfkir);
            }

            if (lblBannerPeriodeVal != null) lblBannerPeriodeVal.Text = periodLabel;
            if (lblBannerTphVal != null) lblBannerTphVal.Text = totalTph.ToString("#,##0");
            if (lblBannerMatangVal != null) lblBannerMatangVal.Text = totalMatang.ToString("#,##0");
            if (lblBannerMentahVal != null) lblBannerMentahVal.Text = totalMentah.ToString("#,##0");
            if (lblBannerAfkirVal != null) lblBannerAfkirVal.Text = totalAfkir.ToString("#,##0");

            if (lblKpiTotalTphVal != null) lblKpiTotalTphVal.Text = totalTph.ToString("#,##0");
            if (lblKpiTotalJanjangVal != null) lblKpiTotalJanjangVal.Text = totalJanjang.ToString("#,##0");
            if (lblKpiMatangVal != null) lblKpiMatangVal.Text = totalMatang.ToString("#,##0");
            if (lblKpiMentahVal != null) lblKpiMentahVal.Text = totalMentah.ToString("#,##0");
            if (lblKpiAfkirVal != null) lblKpiAfkirVal.Text = totalAfkir.ToString("#,##0");
        }

        private Image LoadPhotoBitmap(string photoRef, out string sourceInfo)
        {
            sourceInfo = "Path DB kosong";
            if (string.IsNullOrWhiteSpace(photoRef)) return null;

            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            }
            catch { }

            string cleanRef = photoRef.Trim().Replace('/', '\\');
            string fileName = Path.GetFileName(cleanRef);

            string[] baseSearchFolders = new string[]
            {
                @"C:\Users\SKYMIT59\Aplikasi Spectra\SPECTRA_BACKEND\uploads",
                @"C:\Users\SKYMIT59\Aplikasi Spectra\SPECTRA_BACKEND",
                @"C:\Users\SKYMIT60\SPECTRA-backend\uploads",
                Path.Combine(Application.StartupPath, "uploads"),
                Application.StartupPath
            };

            if (File.Exists(photoRef))
            {
                try
                {
                    using (FileStream fs = new FileStream(photoRef, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        sourceInfo = $"Ditemukan di path lokal: {photoRef}";
                        return Image.FromStream(fs);
                    }
                }
                catch { }
            }

            foreach (string folder in baseSearchFolders)
            {
                if (Directory.Exists(folder))
                {
                    string directFile = Path.Combine(folder, fileName);
                    if (File.Exists(directFile))
                    {
                        try
                        {
                            using (FileStream fs = new FileStream(directFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                sourceInfo = $"Ditemukan di folder: {directFile}";
                                return Image.FromStream(fs);
                            }
                        }
                        catch { }
                    }
                }
            }

            string tunnelBase = "https://manually-entitled-qui-encourages.trycloudflare.com";
            List<string> candidateUrls = new List<string>
            {
                $"http://localhost:5005/uploads/{fileName}",
                $"http://127.0.0.1:5005/uploads/{fileName}",
                $"{tunnelBase}/uploads/{fileName}"
            };

            using (WebClient wc = new WebClient())
            {
                wc.Headers.Add("User-Agent", "Mozilla/5.0 mySAP365-SPECTRA");
                foreach (string url in candidateUrls)
                {
                    try
                    {
                        byte[] imgBytes = wc.DownloadData(url);
                        if (imgBytes != null && imgBytes.Length > 0)
                        {
                            using (MemoryStream ms = new MemoryStream(imgBytes))
                            {
                                sourceInfo = $"Berhasil diunduh dari: {url}";
                                return Image.FromStream(ms);
                            }
                        }
                    }
                    catch { }
                }
            }

            sourceInfo = $"File '{fileName}' tidak ditemukan.";
            return null;
        }

        private void InitPeriodDropdowns(string tableName)
        {
            if (_5Connect_CBoxPrdY == null && _5Connect_CBoxPrdM == null) return;
            if (string.IsNullOrEmpty(activeConnString)) return;

            List<string> years = new List<string>();
            string latestYear = DateTime.Now.Year.ToString();

            try
            {
                using (SqlConnection conn = new SqlConnection(activeConnString))
                {
                    conn.Open();
                    string sqlYears = $"SELECT DISTINCT YEAR(HARVEST_DATE) AS prd_year FROM {tableName} WHERE HARVEST_DATE IS NOT NULL ORDER BY prd_year DESC";
                    using (SqlCommand cmd = new SqlCommand(sqlYears, conn))
                    using (SqlDataReader rd = cmd.ExecuteReader())
                    {
                        while (rd.Read())
                        {
                            if (!rd.IsDBNull(0)) years.Add(rd[0].ToString());
                        }
                    }
                }
            }
            catch { }

            if (!years.Contains(DateTime.Now.Year.ToString()))
            {
                years.Insert(0, DateTime.Now.Year.ToString());
            }

            if (_5Connect_CBoxPrdY != null)
            {
                RepositoryItemComboBox repoY = new RepositoryItemComboBox();
                repoY.TextEditStyle = TextEditStyles.DisableTextEditor;
                foreach (var y in years) repoY.Items.Add(y);
                _5Connect_CBoxPrdY.Edit = repoY;
                _5Connect_CBoxPrdY.EditValue = latestYear;
                _5Connect_CBoxPrdY.Visibility = BarItemVisibility.Always;
                _5Connect_CBoxPrdY.Enabled = true;
                _5Connect_CBoxPrdY.EditValueChanged -= OnPeriodFilterChanged;
                _5Connect_CBoxPrdY.EditValueChanged += OnPeriodFilterChanged;
            }

            if (_5Connect_CBoxPrdM != null)
            {
                RepositoryItemComboBox repoM = new RepositoryItemComboBox();
                repoM.TextEditStyle = TextEditStyles.DisableTextEditor;
                repoM.Items.Add("ALL");
                for (int i = 1; i <= 12; i++) repoM.Items.Add(i.ToString("D2"));
                _5Connect_CBoxPrdM.Edit = repoM;
                _5Connect_CBoxPrdM.EditValue = "ALL";
                _5Connect_CBoxPrdM.Visibility = BarItemVisibility.Always;
                _5Connect_CBoxPrdM.Enabled = true;
                _5Connect_CBoxPrdM.EditValueChanged -= OnPeriodFilterChanged;
                _5Connect_CBoxPrdM.EditValueChanged += OnPeriodFilterChanged;
            }
        }

        private void OnPeriodFilterChanged(object sender, EventArgs e)
        {
            activePeriodFilter = "RIBBON_PERIOD";
            UpdateFilterButtonsVisual();
            UpdateDashboardSummary();
            ApplyFilter();
        }

        private void _5Connect_btnRefresh_ItemClick(object sender, ItemClickEventArgs e)
        {
            Syb3DBload(sender, "SPECTRA_DB", _tblConsolBA);
        }

        private void frmSPR1_Idle(object sender, EventArgs e)
        {
            if (this.IsDisposed) return;
        }

        private void frmSPR1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose(true);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private class EstateCardControl : PanelControl
        {
            public EstateDef Definition { get; private set; }
            public event EventHandler<string> CardClicked;

            private LabelControl lblCodeBadge;
            private LabelControl lblEstateName;
            private LabelControl lblDescription;
            private LabelControl lblTotalJanjangVal;
            private LabelControl lblBreakdownText;
            private LabelControl lblMatangPctBadge;
            private SimpleButton btnViewReport;
            private bool isSelected = false;

            public bool IsSelected
            {
                get { return isSelected; }
                set { isSelected = value; this.Invalidate(); }
            }

            public EstateCardControl(EstateDef def)
            {
                this.SetStyle(ControlStyles.AllPaintingInWmPaint |
                              ControlStyles.UserPaint |
                              ControlStyles.OptimizedDoubleBuffer |
                              ControlStyles.ResizeRedraw, true);
                this.DoubleBuffered = true;
                this.Definition = def;
                InitializeCard();
            }

            private void InitializeCard()
            {
                this.BorderStyle = BorderStyles.NoBorder;
                this.Appearance.BackColor = ColorCanvasBg;
                this.Appearance.Options.UseBackColor = true;
                this.Cursor = Cursors.Hand;

                this.Paint += (s, pe) =>
                {
                    pe.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    Rectangle r = new Rectangle(0, 0, this.Width - 1, this.Height - 1);
                    using (GraphicsPath path = CreateRoundedRectangle(r, 10))
                    {
                        using (SolidBrush br = new SolidBrush(Color.White))
                            pe.Graphics.FillPath(br, path);

                        Color borderCol = isSelected ? ColorAccent : ColorBorder;
                        float borderW = isSelected ? 2f : 1.2f;
                        using (Pen p = new Pen(borderCol, borderW))
                            pe.Graphics.DrawPath(p, path);
                    }
                };

                lblCodeBadge = new LabelControl();
                lblCodeBadge.Text = Definition.EstateCode;
                lblCodeBadge.Appearance.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
                lblCodeBadge.Appearance.ForeColor = Color.White;
                lblCodeBadge.AutoSizeMode = LabelAutoSizeMode.None;
                lblCodeBadge.Size = new Size(44, 24);
                lblCodeBadge.Location = new Point(12, 12);
                lblCodeBadge.Paint += (s, pe) =>
                {
                    pe.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    Rectangle r = new Rectangle(0, 0, lblCodeBadge.Width - 1, lblCodeBadge.Height - 1);
                    using (GraphicsPath path = CreateRoundedRectangle(r, 6))
                    using (SolidBrush br = new SolidBrush(ColorPrimary))
                        pe.Graphics.FillPath(br, path);
                    TextRenderer.DrawText(pe.Graphics, lblCodeBadge.Text, lblCodeBadge.Font, r, Color.White,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                };
                this.Controls.Add(lblCodeBadge);

                lblEstateName = new LabelControl();
                lblEstateName.Text = Definition.EstateName;
                lblEstateName.Appearance.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
                lblEstateName.Appearance.ForeColor = ColorSlate900;
                lblEstateName.Location = new Point(64, 11);
                this.Controls.Add(lblEstateName);

                lblDescription = new LabelControl();
                lblDescription.Text = Definition.Description;
                lblDescription.Appearance.Font = new Font("Segoe UI", 7.5F);
                lblDescription.Appearance.ForeColor = ColorSlate600;
                lblDescription.Location = new Point(64, 28);
                this.Controls.Add(lblDescription);

                lblTotalJanjangVal = new LabelControl();
                lblTotalJanjangVal.Text = "0 Janjang (0 TPH)";
                lblTotalJanjangVal.Appearance.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
                lblTotalJanjangVal.Appearance.ForeColor = ColorPrimary;
                lblTotalJanjangVal.Location = new Point(12, 54);
                this.Controls.Add(lblTotalJanjangVal);

                lblBreakdownText = new LabelControl();
                lblBreakdownText.Text = "Matang: 0  |  Mentah: 0  |  Afkir: 0";
                lblBreakdownText.Appearance.Font = new Font("Segoe UI", 8F);
                lblBreakdownText.Appearance.ForeColor = ColorSlate800;
                lblBreakdownText.Location = new Point(12, 82);
                this.Controls.Add(lblBreakdownText);

                lblMatangPctBadge = new LabelControl();
                lblMatangPctBadge.Text = "RASIO MATANG: 0%";
                lblMatangPctBadge.Appearance.Font = new Font("Segoe UI", 8F, FontStyle.Bold);
                lblMatangPctBadge.Appearance.ForeColor = ColorMatang;
                lblMatangPctBadge.Appearance.BackColor = ColorMatangBg;
                lblMatangPctBadge.Appearance.BorderColor = ColorMatangBorder;
                lblMatangPctBadge.AutoSizeMode = LabelAutoSizeMode.None;
                lblMatangPctBadge.Paint += (s, pe) =>
                {
                    pe.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    Rectangle r = new Rectangle(0, 0, lblMatangPctBadge.Width - 1, lblMatangPctBadge.Height - 1);
                    using (GraphicsPath path = CreateRoundedRectangle(r, r.Height / 2))
                    {
                        using (SolidBrush br = new SolidBrush(lblMatangPctBadge.Appearance.BackColor))
                            pe.Graphics.FillPath(br, path);
                        using (Pen p = new Pen(lblMatangPctBadge.Appearance.BorderColor, 1f))
                            pe.Graphics.DrawPath(p, path);
                    }
                    TextRenderer.DrawText(pe.Graphics, lblMatangPctBadge.Text, lblMatangPctBadge.Font, r, lblMatangPctBadge.Appearance.ForeColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                };
                this.Controls.Add(lblMatangPctBadge);

                btnViewReport = new SimpleButton();
                btnViewReport.Text = "Buka Rincian Kebun";
                btnViewReport.Appearance.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
                btnViewReport.Appearance.ForeColor = ColorPrimary;
                btnViewReport.Appearance.Options.UseForeColor = true;
                btnViewReport.Click += (s, e) => CardClicked?.Invoke(this, Definition.EstateCode);
                this.Controls.Add(btnViewReport);

                this.Click += (s, e) => CardClicked?.Invoke(this, Definition.EstateCode);
                lblEstateName.Click += (s, e) => CardClicked?.Invoke(this, Definition.EstateCode);
            }

            public void UpdateLayout(int width, int height)
            {
                this.Size = new Size(width, height);
                int pad = 12;
                int innerW = width - (pad * 2);

                lblMatangPctBadge.Location = new Point(pad, 114);
                lblMatangPctBadge.Size = new Size(innerW, 26);

                btnViewReport.Location = new Point(pad, height - 42);
                btnViewReport.Size = new Size(innerW, 30);
            }

            public void UpdateSummary(int tphCount, int totalJanjang, int matang, int mentah, int afkir)
            {
                lblTotalJanjangVal.Text = $"{totalJanjang:#,##0} Janjang ({tphCount} TPH)";
                lblBreakdownText.Text = $"Matang: {matang:#,##0}  |  Mentah: {mentah:#,##0}  |  Afkir: {afkir:#,##0}";

                double pct = totalJanjang > 0 ? (matang * 100.0) / totalJanjang : 0.0;
                lblMatangPctBadge.Text = $"RASIO MATANG: {pct:0.0}%";

                if (totalJanjang == 0)
                {
                    lblMatangPctBadge.Appearance.BackColor = ColorBrandLight;
                    lblMatangPctBadge.Appearance.ForeColor = ColorSlate600;
                    lblMatangPctBadge.Appearance.BorderColor = ColorBorder;
                }
                else if (pct >= 80.0)
                {
                    lblMatangPctBadge.Appearance.BackColor = ColorMatangBg;
                    lblMatangPctBadge.Appearance.ForeColor = ColorMatang;
                    lblMatangPctBadge.Appearance.BorderColor = ColorMatangBorder;
                }
                else
                {
                    lblMatangPctBadge.Appearance.BackColor = ColorMentahBg;
                    lblMatangPctBadge.Appearance.ForeColor = ColorMentah;
                    lblMatangPctBadge.Appearance.BorderColor = ColorMentahBorder;
                }
                lblMatangPctBadge.Invalidate();
            }
        }
    }
}
