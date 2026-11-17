using DevExpress.Data.Filtering;
using DevExpress.Xpo;
using DevExpress.Xpo.DB;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraTab;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows.Forms;

namespace SPD1
{
    public partial class frmSPD1 : Syb3.Form.RB1.Syb3FormRB_1cn_TD
    {
        private DataTable dtMaster = null;
        private SqlDataAdapter daMaster = null;

        // Komponen Panel Kanan Preview Foto Janjang & Mutu TBS Lapangan
        private XtraTabControl tabRightPanel;
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
        private SimpleButton btnCloseTab;
        private SimpleButton btnBrowsePhoto;

        public frmSPD1()
        {
            InitializeComponent();
        }

        public frmSPD1(Syb3init1.mySyb3Init param_AppInit1, string[] param_App2, Object[] param_App3)
            : base(param_AppInit1)
        {
            InitializeComponent();
            _Syb3InitPub1 = param_AppInit1;
            _AppParam3 = param_App3;
        }

        private void frmSPD1_Load(object sender, EventArgs e)
        {
            if (this.DesignMode) return;
            Application.Idle += new EventHandler(frmSPD1_Idle);

            Syb3FormRB_1cn_reLoad1(sender, e, "EO", null, false, false, _Syb3InitPub1, _Syb3InitLoc1);

            _Syb3InitLoc1[_Syb3InitPub1.vLockModulwithPCR1GA] = "false";
            _Syb3InitLoc1[_Syb3InitPub1.vCountKeyClosing] = "1";
            _Syb3InitLoc1[_Syb3InitPub1.vfirstField] = "ID";

            _gridViewTD1.IndicatorWidth = 45;
            _gridViewTD1.RowHeight = 0;
            _gridViewTD1.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDown;
            _gridViewTD1.OptionsDetail.EnableMasterViewMode = false;
            _gridViewTD1.OptionsDetail.ShowDetailTabs = false;
            _Syb3InitLoc1[_Syb3InitPub1.vNoUserWorkArea] = "false";

            // Data transaksi murni dari lapangan, input baris manual di ERP dinonaktifkan
            _gridViewTD1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.None;
            _gridViewTD1.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.False;
            if (_1Cells_btnInsert != null) _1Cells_btnInsert.Enabled = false;

            // Membangun panel samping foto TBS
            BuildRightPhotoTab();

            _gridControlTD1.MouseDown += _gridControlTD1_MouseDown;
            _gridViewTD1.RowCellClick += _gridViewTD1_RowCellClick;
            _gridViewTD1.FocusedRowChanged += _gridViewTD1_FocusedRowChanged;

            if (_5Connect_BUnitDB != null)
            {
                RepositoryItemComboBox repoUnit = new RepositoryItemComboBox();
                repoUnit.Items.Add("ESTATE");
                repoUnit.TextEditStyle = TextEditStyles.DisableTextEditor;
                _5Connect_BUnitDB.Edit = repoUnit;
                _5Connect_BUnitDB.EditValue = "ESTATE";
                _5Connect_BUnitDB.Enabled = false;
            }

            if (_5Connect_CBoxPrdY != null)
            {
                _5Connect_CBoxPrdY.EditValue = 2026;
                _5Connect_CBoxPrdY.Enabled = false;
            }

            if (_5Connect_CBoxPrdM != null)
            {
                _5Connect_CBoxPrdM.Visibility = BarItemVisibility.Never;
            }

            if (_4Funct_btnHRM1 != null) _4Funct_btnHRM1.Visibility = BarItemVisibility.Never;
            if (_4Funct_btnTgl != null) _4Funct_btnTgl.Visibility = BarItemVisibility.Never;
            if (_4Funct_btnHRM9 != null) _4Funct_btnHRM9.Visibility = BarItemVisibility.Never;

            RibbonControl rb = this.Ribbon;
            if (rb == null)
            {
                foreach (Control c in this.Controls)
                {
                    if (c is RibbonControl rc) { rb = rc; break; }
                }
            }

            if (rb != null)
            {
                foreach (RibbonPage page in rb.Pages)
                {
                    if (page.Text.IndexOf("Formula", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        page.Name.IndexOf("Formula", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        page.Visible = false;
                    }
                }
            }

            foreach (GridBand band in _gridViewTD1.Bands)
            {
                if (band.Caption.IndexOf("Dibuat", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    band.Caption.IndexOf("Created", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    band.Visible = false;
                }
            }
        }

        private void BuildRightPhotoTab()
        {
            if (_panelCtrl1 == null) return;

            _panelCtrl1.Controls.Clear();
            _panelCtrl1.Dock = DockStyle.Right;
            _panelCtrl1.Width = 390;
            _panelCtrl1.Visible = false;

            if (_splitter1 != null)
            {
                _splitter1.Dock = DockStyle.Right;
                _splitter1.Visible = false;
            }

            if (_panelCtrl2 != null)
            {
                _panelCtrl2.Dock = DockStyle.Fill;
            }

            tabRightPanel = new XtraTabControl();
            tabRightPanel.Dock = DockStyle.Fill;
            tabRightPanel.HeaderLocation = TabHeaderLocation.Top;

            pagePhotoPreview = new XtraTabPage();
            pagePhotoPreview.Text = "Foto Kualitas Buah TPH";

            btnCloseTab = new SimpleButton();
            btnCloseTab.Text = "✕ Tutup";
            btnCloseTab.Size = new Size(68, 24);
            btnCloseTab.Location = new Point(298, 8);
            btnCloseTab.Click += (s, ev) =>
            {
                _panelCtrl1.Visible = false;
                if (_splitter1 != null) _splitter1.Visible = false;
            };
            pagePhotoPreview.Controls.Add(btnCloseTab);

            btnBrowsePhoto = new SimpleButton();
            btnBrowsePhoto.Text = "📁 Cari File...";
            btnBrowsePhoto.Size = new Size(80, 24);
            btnBrowsePhoto.Location = new Point(212, 8);
            btnBrowsePhoto.Click += (s, ev) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Pilih Foto TBS";
                    ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*";
                    string defDir = @"C:\Users\SKYMIT59\Aplikasi Spectra\SPECTRA_BACKEND\uploads";
                    if (Directory.Exists(defDir)) ofd.InitialDirectory = defDir;

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            using (FileStream fs = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read))
                            {
                                picConditionViewer.Image = Image.FromStream(fs);
                            }
                            lblPhotoStatus.Text = "Manual: " + Path.GetFileName(ofd.FileName);
                            lblPhotoStatus.Appearance.ForeColor = Color.FromArgb(39, 174, 96);
                        }
                        catch (Exception ex)
                        {
                            XtraMessageBox.Show("Gagal memuat gambar: " + ex.Message);
                        }
                    }
                }
            };
            pagePhotoPreview.Controls.Add(btnBrowsePhoto);

            lblTransactionTitle = new LabelControl();
            lblTransactionTitle.Text = "-";
            lblTransactionTitle.Appearance.Font = new Font("Tahoma", 9.5F, FontStyle.Bold);
            lblTransactionTitle.Appearance.ForeColor = Color.FromArgb(41, 128, 185);
            lblTransactionTitle.Location = new Point(14, 10);
            lblTransactionTitle.MaximumSize = new Size(190, 22);
            pagePhotoPreview.Controls.Add(lblTransactionTitle);

            lblEstateAndBlockInfo = new LabelControl();
            lblEstateAndBlockInfo.Text = "Lokasi: -";
            lblEstateAndBlockInfo.Appearance.Font = new Font("Tahoma", 8.25F, FontStyle.Regular);
            lblEstateAndBlockInfo.Appearance.ForeColor = Color.DarkSlateGray;
            lblEstateAndBlockInfo.Location = new Point(14, 34);
            pagePhotoPreview.Controls.Add(lblEstateAndBlockInfo);

            picConditionViewer = new PictureEdit();
            picConditionViewer.Location = new Point(14, 54);
            picConditionViewer.Size = new Size(352, 235);
            picConditionViewer.Properties.SizeMode = PictureSizeMode.Zoom;
            picConditionViewer.Properties.NullText = "Tidak ada lampiran foto TBS";
            picConditionViewer.Properties.ShowMenu = true;
            picConditionViewer.Properties.PictureInterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            picConditionViewer.BorderStyle = BorderStyles.Simple;
            pagePhotoPreview.Controls.Add(picConditionViewer);

            lblPhotoStatus = new LabelControl();
            lblPhotoStatus.Text = "Status: -";
            lblPhotoStatus.Appearance.Font = new Font("Tahoma", 7.5F, FontStyle.Italic);
            lblPhotoStatus.Appearance.ForeColor = Color.Gray;
            lblPhotoStatus.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            lblPhotoStatus.AutoSizeMode = LabelAutoSizeMode.Vertical;
            lblPhotoStatus.Location = new Point(14, 292);
            lblPhotoStatus.Size = new Size(352, 24);
            pagePhotoPreview.Controls.Add(lblPhotoStatus);

            lblQualitySummary = new LabelControl();
            lblQualitySummary.Text = "Total Janjang: -";
            lblQualitySummary.Appearance.Font = new Font("Tahoma", 9.5F, FontStyle.Bold);
            lblQualitySummary.Appearance.ForeColor = Color.FromArgb(44, 62, 80);
            lblQualitySummary.Location = new Point(14, 318);
            pagePhotoPreview.Controls.Add(lblQualitySummary);

            lblDetailsCount = new LabelControl();
            lblDetailsCount.Text = "Matang: - | Mentah: - | Busuk: -\nGagang Panjang: - | Abnormal: -";
            lblDetailsCount.Appearance.Font = new Font("Tahoma", 8.25F, FontStyle.Regular);
            lblDetailsCount.Appearance.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
            lblDetailsCount.AutoSizeMode = LabelAutoSizeMode.Vertical;
            lblDetailsCount.Location = new Point(14, 338);
            lblDetailsCount.Size = new Size(352, 32);
            pagePhotoPreview.Controls.Add(lblDetailsCount);

            lblKraniInfo = new LabelControl();
            lblKraniInfo.Text = "Petugas Kerani: -";
            lblKraniInfo.Appearance.Font = new Font("Tahoma", 8F);
            lblKraniInfo.Appearance.ForeColor = Color.DarkCyan;
            lblKraniInfo.Location = new Point(14, 374);
            pagePhotoPreview.Controls.Add(lblKraniInfo);

            lblSyncStatusBadge = new LabelControl();
            lblSyncStatusBadge.Text = "Status: SYNCED";
            lblSyncStatusBadge.Appearance.Font = new Font("Tahoma", 8.5F, FontStyle.Bold);
            lblSyncStatusBadge.Appearance.ForeColor = Color.FromArgb(39, 174, 96);
            lblSyncStatusBadge.Location = new Point(14, 394);
            pagePhotoPreview.Controls.Add(lblSyncStatusBadge);

            lblTimestampText = new LabelControl();
            lblTimestampText.Text = "Waktu: -";
            lblTimestampText.Appearance.Font = new Font("Tahoma", 7.5F);
            lblTimestampText.Appearance.ForeColor = Color.Gray;
            lblTimestampText.Location = new Point(14, 414);
            pagePhotoPreview.Controls.Add(lblTimestampText);

            tabRightPanel.TabPages.Add(pagePhotoPreview);
            _panelCtrl1.Controls.Add(tabRightPanel);
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

            // 1. PENCARIAN DI FOLDER FISIK SERVER
            string[] baseSearchFolders = new string[]
            {
                @"C:\Users\SKYMIT59\Aplikasi Spectra\SPECTRA_BACKEND\uploads",
                @"C:\Users\SKYMIT59\Aplikasi Spectra\SPECTRA_BACKEND",
                @"C:\Users\SKYMIT59\Aplikasi Spectra\SPECTRA_BACKEND\static",
                @"C:\Users\SKYMIT60\SPECTRA-backend\uploads",
                @"C:\Users\SKYMIT60\SPECTRA-backend",
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

                    string relativeFile = Path.Combine(folder, cleanRef.TrimStart('\\'));
                    if (File.Exists(relativeFile))
                    {
                        try
                        {
                            using (FileStream fs = new FileStream(relativeFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                sourceInfo = $"Ditemukan di: {relativeFile}";
                                return Image.FromStream(fs);
                            }
                        }
                        catch { }
                    }

                    try
                    {
                        string[] matches = Directory.GetFiles(folder, fileName, SearchOption.AllDirectories);
                        if (matches != null && matches.Length > 0)
                        {
                            using (FileStream fs = new FileStream(matches[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                sourceInfo = $"Ditemukan rekursif di: {matches[0]}";
                                return Image.FromStream(fs);
                            }
                        }
                    }
                    catch { }
                }
            }

            // 2. PENGAMBILAN MELALUI HTTP FASTAPI / CLOUDFLARE TUNNEL
            string tunnelBase = "https://sociology-complications-officials-insert.trycloudflare.com/";
            List<string> candidateUrls = new List<string>();

            if (photoRef.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                photoRef.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                candidateUrls.Add(photoRef);
            }
            else
            {
                candidateUrls.Add($"http://localhost:5005/uploads/{fileName}");
                candidateUrls.Add($"http://localhost:5005/static/{fileName}");
                candidateUrls.Add($"http://localhost:5005/{fileName}");
                candidateUrls.Add($"http://127.0.0.1:5005/uploads/{fileName}");
                candidateUrls.Add($"{tunnelBase}/uploads/{fileName}");
                candidateUrls.Add($"{tunnelBase}/static/{fileName}");
                candidateUrls.Add($"{tunnelBase}/{fileName}");
            }

            using (WebClient wc = new WebClient())
            {
                wc.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) mySAP365-SPECTRA");
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

            sourceInfo = $"File '{fileName}' tidak ditemukan di folder backend maupun via URL FastAPI (Port 5000).";
            return null;
        }

        private void _gridControlTD1_MouseDown(object sender, MouseEventArgs e)
        {
            GridHitInfo hit = _gridViewTD1.CalcHitInfo(e.Location);
            if (hit.InRowCell && hit.Column != null && hit.Column.FieldName == "PHOTO_PATH")
            {
                OpenPhotoConditionTab(hit.RowHandle);
            }
        }

        private void _gridViewTD1_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            if (e.Column != null && e.Column.FieldName == "PHOTO_PATH")
            {
                OpenPhotoConditionTab(e.RowHandle);
            }
        }

        private void _gridViewTD1_FocusedRowChanged(object sender, FocusedRowChangedEventArgs e)
        {
            if (_panelCtrl1 != null && _panelCtrl1.Visible && e.FocusedRowHandle >= 0)
            {
                OpenPhotoConditionTab(e.FocusedRowHandle);
            }
        }

        private void OpenPhotoConditionTab(int rowHandle)
        {
            if (rowHandle < 0) return;

            string idTrx = _gridViewTD1.GetRowCellValue(rowHandle, "ID")?.ToString() ?? "-";
            string estate = _gridViewTD1.GetRowCellValue(rowHandle, "ESTATE_CODE")?.ToString() ?? "-";
            string afdeling = _gridViewTD1.GetRowCellValue(rowHandle, "AFDELING_CODE")?.ToString() ?? "-";
            string block = _gridViewTD1.GetRowCellValue(rowHandle, "BLOCK_ID")?.ToString() ?? "-";
            string tph = _gridViewTD1.GetRowCellValue(rowHandle, "TPH_NUMBER")?.ToString() ?? "-";
            string nik = _gridViewTD1.GetRowCellValue(rowHandle, "KRANI_NIK")?.ToString() ?? "-";
            string kraniName = _gridViewTD1.GetRowCellValue(rowHandle, "KRANI_NAME")?.ToString() ?? "-";
            string harvestDate = _gridViewTD1.GetRowCellValue(rowHandle, "HARVEST_DATE")?.ToString() ?? "-";
            string matang = _gridViewTD1.GetRowCellValue(rowHandle, "MATANG_COUNT")?.ToString() ?? "0";
            string mentah = _gridViewTD1.GetRowCellValue(rowHandle, "MENTAH_COUNT")?.ToString() ?? "0";
            string busuk = _gridViewTD1.GetRowCellValue(rowHandle, "BUAH_BUSUK_COUNT")?.ToString() ?? "0";
            string tangkai = _gridViewTD1.GetRowCellValue(rowHandle, "TANGKAI_PANJANG_COUNT")?.ToString() ?? "0";
            string abnormal = _gridViewTD1.GetRowCellValue(rowHandle, "ABNORMAL_COUNT")?.ToString() ?? "0";
            string total = _gridViewTD1.GetRowCellValue(rowHandle, "TOTAL_BUAH")?.ToString() ?? "0";
            string photoPath = _gridViewTD1.GetRowCellValue(rowHandle, "PHOTO_PATH")?.ToString() ?? "";
            string syncStat = _gridViewTD1.GetRowCellValue(rowHandle, "SYNC_STATUS")?.ToString() ?? "SYNCED";
            string createdAt = _gridViewTD1.GetRowCellValue(rowHandle, "CREATED_AT")?.ToString() ?? "-";

            lblTransactionTitle.Text = $"ID #{idTrx} - TPH {tph}";
            lblEstateAndBlockInfo.Text = $"Kebun: {estate} | Afdeling: {afdeling} | Blok: {block}";
            lblQualitySummary.Text = $"Total Janjang: {total} TBS";
            lblDetailsCount.Text = $"Matang: {matang} | Mentah: {mentah} | Busuk: {busuk}\nGagang Panjang: {tangkai} | Abnormal: {abnormal}";
            lblKraniInfo.Text = $"Petugas Kerani: {kraniName} ({nik})\nTgl Panen: {harvestDate}";
            lblSyncStatusBadge.Text = $"Status: {syncStat.ToUpper()}";
            lblTimestampText.Text = $"Upload: {createdAt}";

            if (syncStat.Equals("SYNCED", StringComparison.OrdinalIgnoreCase))
            {
                lblSyncStatusBadge.Appearance.ForeColor = Color.FromArgb(39, 174, 96);
            }
            else
            {
                lblSyncStatusBadge.Appearance.ForeColor = Color.FromArgb(211, 84, 0);
            }

            picConditionViewer.Image = LoadPhotoBitmap(photoPath, out string sourceInfo);
            lblPhotoStatus.Text = sourceInfo;
            if (picConditionViewer.Image != null)
            {
                lblPhotoStatus.Appearance.ForeColor = Color.FromArgb(39, 174, 96);
            }
            else
            {
                lblPhotoStatus.Appearance.ForeColor = Color.FromArgb(192, 57, 43);
            }

            _panelCtrl1.Visible = true;
            if (_splitter1 != null) _splitter1.Visible = true;
            _panelCtrl1.BringToFront();
        }

        private void frmSPD1_Shown(object sender, EventArgs e)
        {
            Syb3dbLoadD(sender);
            _Syb3InitPub1._mySyb3Db.Syb3FocusLastPosition(sender, _gridViewTD1, _Syb3InitPub1, _Syb3InitLoc1);
            _Syb3InitPub1.loadWait1 = false;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void Syb3DBload(object sender, string dbName1)
        {
            _Syb3InitPub1._mySyb3Layout.Syb3marqBarStart(sender, _AppParam3, _gridViewTD1, _Rb1cn_barProgres1, _Rb1cn_barZoom1, _Rb1cn_barStatSheet, _Rb1cn_barProgres1_repoMarque, this, _Syb3InitPub1, _Syb3InitLoc1);

            string targetDatabase = "SPECTRA_DB";
            _Syb3InitLoc1[_Syb3InitPub1.vSwapDB1] = targetDatabase;

            string connString = string.Empty;

            try
            {
                var dsTemp = _Syb3InitPub1._mySyb3Db.Syb3IDDict1(targetDatabase, false, _Syb3InitPub1, _Syb3InitLoc1);
                var propConn = dsTemp?.GetType().GetProperty("ConnectionString");
                connString = propConn != null ? propConn.GetValue(dsTemp) as string : string.Empty;
            }
            catch { }

            if (string.IsNullOrEmpty(connString))
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
                        connString = csb.ConnectionString;
                    }
                }
                catch { }
            }

            if (string.IsNullOrEmpty(connString))
            {
                connString = "Data Source=localhost;Initial Catalog=SPECTRA_DB;Integrated Security=True;TrustServerCertificate=True;";
            }

            string query = @"
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

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                daMaster = new SqlDataAdapter(query, sqlConn);
                SqlCommandBuilder cmdBuilder = new SqlCommandBuilder(daMaster);

                if (dtMaster != null)
                {
                    dtMaster.Clear();
                    dtMaster.Dispose();
                }

                dtMaster = new DataTable();
                daMaster.Fill(dtMaster);
            }

            _gridControlTD1.DataSource = null;
            _gridControlTD1.DataSource = dtMaster;

            RepositoryItemComboBox repoEstate = new RepositoryItemComboBox();
            repoEstate.Items.AddRange(new object[] { "MRE", "LBE", "BKE", "AGE" });
            repoEstate.TextEditStyle = TextEditStyles.DisableTextEditor;
            _gridControlTD1.RepositoryItems.Add(repoEstate);

            RepositoryItemComboBox repoSync = new RepositoryItemComboBox();
            repoSync.Items.AddRange(new object[] { "SYNCED", "PENDING" });
            repoSync.TextEditStyle = TextEditStyles.DisableTextEditor;
            _gridControlTD1.RepositoryItems.Add(repoSync);

            RepositoryItemSpinEdit repoCount = new RepositoryItemSpinEdit();
            repoCount.IsFloatValue = false;
            repoCount.MinValue = 0;
            repoCount.MaxValue = 9999;
            repoCount.Mask.EditMask = "N0";
            _gridControlTD1.RepositoryItems.Add(repoCount);

            foreach (GridColumn col in _gridViewTD1.Columns)
            {
                string fName = (col.FieldName ?? "").ToUpper();

                if (fName == "ID")
                {
                    col.Caption = "ID";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.OptionsColumn.AllowEdit = false;
                    col.OptionsColumn.ReadOnly = true;
                    col.AppearanceCell.BackColor = Color.FromArgb(240, 240, 240);
                    col.Width = 65;
                }
                else if (fName == "ESTATE_CODE")
                {
                    col.Caption = "Kebun";
                    col.ColumnEdit = repoEstate;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 75;
                }
                else if (fName == "AFDELING_CODE")
                {
                    col.Caption = "Afdeling";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 85;
                }
                else if (fName == "BLOCK_ID")
                {
                    col.Caption = "Blok";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 75;
                }
                else if (fName == "TPH_NUMBER")
                {
                    col.Caption = "No. TPH";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 85;
                }
                else if (fName == "KRANI_NIK")
                {
                    col.Caption = "NIK Kerani";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 95;
                }
                else if (fName == "KRANI_NAME")
                {
                    col.Caption = "Nama Kerani Panen";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.OptionsColumn.AllowEdit = false;
                    col.OptionsColumn.ReadOnly = true;
                    col.Width = 160;
                }
                else if (fName == "HARVEST_DATE")
                {
                    col.Caption = "Tanggal Panen";
                    col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    col.DisplayFormat.FormatString = "dd-MMM-yyyy";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 100;
                }
                else if (fName == "MATANG_COUNT")
                {
                    col.Caption = "Buah Matang";
                    col.ColumnEdit = repoCount;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 75;
                }
                else if (fName == "MENTAH_COUNT")
                {
                    col.Caption = "Buah Mentah";
                    col.ColumnEdit = repoCount;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 75;
                }
                else if (fName == "BUAH_BUSUK_COUNT")
                {
                    col.Caption = "Buah Busuk";
                    col.ColumnEdit = repoCount;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 75;
                }
                else if (fName == "TANGKAI_PANJANG_COUNT")
                {
                    col.Caption = "Tangkai Panjang";
                    col.ColumnEdit = repoCount;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 85;
                }
                else if (fName == "ABNORMAL_COUNT")
                {
                    col.Caption = "Buah Abnormal";
                    col.ColumnEdit = repoCount;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 80;
                }
                else if (fName == "TOTAL_BUAH")
                {
                    col.Caption = "Total Janjang";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceCell.Font = new Font("Tahoma", 8.25F, FontStyle.Bold);
                    col.Width = 95;
                }
                else if (fName == "PHOTO_PATH")
                {
                    col.Caption = "Foto TBS \n(Klik Preview)";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceCell.ForeColor = Color.FromArgb(41, 128, 185);
                    col.AppearanceCell.Font = new Font("Tahoma", 8.25F, FontStyle.Underline);
                    col.OptionsColumn.AllowEdit = false;
                    col.OptionsColumn.ReadOnly = true;
                    col.Width = 170;
                }
                else if (fName == "SYNC_STATUS")
                {
                    col.Caption = "Status Sinkron";
                    col.ColumnEdit = repoSync;
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 100;
                }
                else if (fName == "CREATED_AT")
                {
                    col.Caption = "Waktu Upload";
                    col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
                    col.DisplayFormat.FormatString = "dd-MMM-yyyy HH:mm";
                    col.OptionsColumn.AllowEdit = false;
                    col.OptionsColumn.ReadOnly = true;
                    col.AppearanceCell.BackColor = Color.FromArgb(240, 240, 240);
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 140;
                }
            }

            _gridViewTD1.RowHeight = 0;
            _gridViewTD1.OptionsView.ShowFooter = true;
            _gridViewTD1.IndicatorWidth = 45;
            _gridControlTD1.RefreshDataSource();
            _gridViewTD1.RefreshData();

            _Syb3InitPub1._mySyb3Layout.Syb3marqBarStop(sender, _gridViewTD1, _Rb1cn_barProgres1, _Rb1cn_barZoom1, _Rb1cn_barStatSheet, _Rb1cn_barProgres1_repo, this, _Syb3InitPub1, _Syb3InitLoc1);
        }

        private void SaveDataToDatabase()
        {
            _gridViewTD1.PostEditor();
            _gridViewTD1.UpdateCurrentRow();

            if (daMaster != null && dtMaster != null)
            {
                try
                {
                    SqlCommandBuilder cb = new SqlCommandBuilder(daMaster);
                    daMaster.Update(dtMaster);
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show("Gagal menyimpan ke database SPECTRA_DB: " + ex.Message, "mySAP365 - SPD1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void _btnQSave_ItemClick(object sender, ItemClickEventArgs e)
        {
            SaveDataToDatabase();
            XtraMessageBox.Show("Data Transaksi TPH_TRANSACTION berhasil disimpan ke database SPECTRA_DB.", "mySAP365 - SPD1", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _gridControlTD1.RefreshDataSource();
            _gridViewTD1.UpdateTotalSummary();
        }

        private void frmSPD1_Idle(object sender, EventArgs e)
        {
            if (this.IsDisposed) return;
        }

        private void _btnQOpen_ItemClick(object sender, ItemClickEventArgs e)
        {
            _gridControlTD1.DataSource = null;
            Syb3dbLoadD(sender);
            _gridViewTD1.RefreshData();
            _gridControlTD1.Refresh();
        }

        private void Syb3dbLoadD(object sender)
        {
            Syb3DBload(sender, "SPECTRA_DB");
        }

        private void _btnQPreview_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (_Syb3InitLoc1[_Syb3InitPub1.vSecPass] == "true" && _Syb3InitLoc1[_Syb3InitPub1.vXtraPrintCustom] == "true")
            {
                _Syb3InitLoc1[_Syb3InitPub1.vSecPass] = "false";
                XtraMessageBox.Show("Hi, " + _Syb3InitPub1.UserID + ", No Customize Format, Please using the default each sheet", "mySAP365 - " + _Syb3InitLoc1[_Syb3InitPub1.vModulCode],
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void _Realtimer1_Tick(object sender, EventArgs e)
        {
        }

        private void _5Connect_btnRefresh_ItemClick(object sender, ItemClickEventArgs e)
        {
            _gridControlTD1.DataSource = null;
            Syb3DBload(sender, "SPECTRA_DB");
            _gridViewTD1.RefreshData();
            _gridControlTD1.Refresh();
        }

        private void _gridViewTD1_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
        }

        private void frmSPD1_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void _btnAppExport_SAP3Encryp_ItemClick(object sender, ItemClickEventArgs e)
        {
        }

        private void _5GetExt_btnGetDataFromDB_Encryp_ItemClick(object sender, ItemClickEventArgs e)
        {
        }

        private void _btnQUndo_ItemClick(object sender, ItemClickEventArgs e)
        {
            _gridControlTD1.DataSource = null;
            Syb3dbLoadD(sender);
        }

        private void _gridViewTD1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            _iCtrl = 2;

            string field = e.Column.FieldName.ToUpper();
            if (field == "MATANG_COUNT" || field == "MENTAH_COUNT" || field == "BUAH_BUSUK_COUNT" || field == "TANGKAI_PANJANG_COUNT" || field == "ABNORMAL_COUNT")
            {
                int matang = Convert.ToInt32(_gridViewTD1.GetRowCellValue(e.RowHandle, "MATANG_COUNT") ?? 0);
                int mentah = Convert.ToInt32(_gridViewTD1.GetRowCellValue(e.RowHandle, "MENTAH_COUNT") ?? 0);
                int busuk = Convert.ToInt32(_gridViewTD1.GetRowCellValue(e.RowHandle, "BUAH_BUSUK_COUNT") ?? 0);
                int tangkai = Convert.ToInt32(_gridViewTD1.GetRowCellValue(e.RowHandle, "TANGKAI_PANJANG_COUNT") ?? 0);
                int abnormal = Convert.ToInt32(_gridViewTD1.GetRowCellValue(e.RowHandle, "ABNORMAL_COUNT") ?? 0);
                int total = matang + mentah + busuk + tangkai + abnormal;

                _gridViewTD1.SetRowCellValue(e.RowHandle, "TOTAL_BUAH", total);
            }

            SaveDataToDatabase();

            if (_panelCtrl1 != null && _panelCtrl1.Visible && e.RowHandle == _gridViewTD1.FocusedRowHandle)
            {
                OpenPhotoConditionTab(e.RowHandle);
            }
        }

        private void _gridViewTD1D_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
        }

        private void _gridViewTD1D_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
        }

        private void _gridViewTD1_InitNewRow(object sender, InitNewRowEventArgs e)
        {
        }

        private void _gridViewTD1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.Column.FieldName == "SYNC_STATUS")
            {
                string val = e.CellValue?.ToString() ?? "";
                if (val.Equals("SYNCED", StringComparison.OrdinalIgnoreCase))
                {
                    e.Appearance.ForeColor = Color.FromArgb(30, 132, 73);
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
                else
                {
                    e.Appearance.ForeColor = Color.FromArgb(211, 84, 0);
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
            }
            else if (e.Column.FieldName == "TOTAL_BUAH")
            {
                e.Appearance.ForeColor = Color.FromArgb(41, 128, 185);
                e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
            }
        }

        private void _4Funct_btnHRM9_EditValueChanged(object sender, EventArgs e)
        {
        }

        private void _4Funct_btnHRM1_EditValueChanged(object sender, EventArgs e)
        {
        }

        private void _4Funct_btnTgl_EditValueChanged(object sender, EventArgs e)
        {
        }

        private void _gridViewTD1_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
        }

        private void _gridViewTD1D_ShownEditor(object sender, EventArgs e)
        {
        }

        private void _gridViewTD2D_ShownEditor(object sender, EventArgs e)
        {
        }

        private void _gridViewTD2D_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
        }

        private void _gridViewTD3D_ShownEditor(object sender, EventArgs e)
        {
        }

        private void _gridViewTD3D_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
        }

        private void _gridViewTD3D_ValidateRow(object sender, ValidateRowEventArgs e)
        {
        }

        private void _btnInsertAllData_ItemClick(object sender, ItemClickEventArgs e)
        {
            _gridControlTD1.DataSource = null;
            Syb3DBload(sender, "SPECTRA_DB");
        }

        private void _btnInsertIoTIntegrasi_ItemClick(object sender, ItemClickEventArgs e)
        {
            _gridControlTD1.DataSource = null;
            Syb3DBload(sender, "SPECTRA_DB");
        }

        private void _gridControlTD1_Click(object sender, EventArgs e)
        {
        }
    }
}