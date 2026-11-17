using DevExpress.Data.Filtering;
using DevExpress.Xpo;
using DevExpress.Xpo.DB;
//-- DEVEXPRESS
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SPM2
{
    public partial class frmSPM2 : Syb3.Form.RB1.Syb3FormRB_1cn_TD
    {
        private DataTable dtMaster = null;
        private SqlDataAdapter daMaster = null;

        public frmSPM2()
        {
            InitializeComponent();
        }

        public frmSPM2(Syb3init1.mySyb3Init param_AppInit1, string[] param_App2, Object[] param_App3)
            : base(param_AppInit1)
        {
            InitializeComponent();
            _Syb3InitPub1 = param_AppInit1;
            _AppParam3 = param_App3;
        }

        private void frmSPM2_Load(object sender, EventArgs e)
        {
            if (this.DesignMode) return;
            Application.Idle += new EventHandler(frmSPM2_Idle);

            Syb3FormRB_1cn_reLoad1(sender, e, "EO", null, false, false, _Syb3InitPub1, _Syb3InitLoc1);

            _Syb3InitLoc1[_Syb3InitPub1.vLockModulwithPCR1GA] = "false";
            _Syb3InitLoc1[_Syb3InitPub1.vCountKeyClosing] = "1";
            _Syb3InitLoc1[_Syb3InitPub1.vfirstField] = "ESTATE_CODE";

            _gridViewTD1.IndicatorWidth = 45;
            _gridViewTD1.RowHeight = 0;
            _gridViewTD1.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.MouseDown;
            _gridViewTD1.OptionsDetail.EnableMasterViewMode = false;
            _gridViewTD1.OptionsDetail.ShowDetailTabs = false;
            _Syb3InitLoc1[_Syb3InitPub1.vNoUserWorkArea] = "false";

            // Event penyimpanan baris selesai diisi/diedit
            _gridViewTD1.RowUpdated += _gridViewTD1_RowUpdated;

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

        private void frmSPM2_Shown(object sender, EventArgs e)
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

            // 1. Coba ambil lewat kamus mySAP365
            try
            {
                var dsTemp = _Syb3InitPub1._mySyb3Db.Syb3IDDict1(targetDatabase, false, _Syb3InitPub1, _Syb3InitLoc1);
                var propConn = dsTemp?.GetType().GetProperty("ConnectionString");
                connString = propConn != null ? propConn.GetValue(dsTemp) as string : string.Empty;
            }
            catch { }

            // 2. Jika belum terdaftar di kamus, ambil koneksi aktif mySAP365 lalu arahkan katalognya ke SPECTRA_DB
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

            // 3. Fallback jika sistem mySAP365 mengunci properti connection string
            if (string.IsNullOrEmpty(connString))
            {
                connString = "Data Source=localhost;Initial Catalog=SPECTRA_DB;Integrated Security=True;TrustServerCertificate=True;";
            }

            string query = @"
                SELECT 
                    ESTATE_CODE, 
                    ESTATE_NAME, 
                    COMPANY_NAME, 
                    IS_ACTIVE 
                FROM dbo.MD_ESTATE 
                ORDER BY ESTATE_CODE ASC";

            SqlConnection sqlConn = new SqlConnection(connString);
            daMaster = new SqlDataAdapter(query, sqlConn);
            SqlCommandBuilder cmdBuilder = new SqlCommandBuilder(daMaster);

            dtMaster = new DataTable();
            daMaster.Fill(dtMaster);

            if (dtMaster.Columns.Contains("ESTATE_CODE"))
            {
                dtMaster.PrimaryKey = new DataColumn[] { dtMaster.Columns["ESTATE_CODE"] };
            }

            _gridControlTD1.DataSource = dtMaster;

            // Repository Item agar menampilkan nilai asli 1 dan 0 (bukan True / False)
            RepositoryItemTextEdit repoIsActive = new RepositoryItemTextEdit();
            repoIsActive.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            _gridControlTD1.RepositoryItems.Add(repoIsActive);

            foreach (GridColumn col in _gridViewTD1.Columns)
            {
                string fName = (col.FieldName ?? "").ToUpper();

                if (fName == "ESTATE_CODE")
                {
                    col.Caption = "Kode Kebun";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 100;
                }
                else if (fName == "ESTATE_NAME")
                {
                    col.Caption = "Nama Kebun";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 180;
                }
                else if (fName == "COMPANY_NAME")
                {
                    col.Caption = "Nama Perusahaan";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 220;
                }
                else if (fName == "IS_ACTIVE")
                {
                    col.Caption = "Status Aktif (1=Ya, 0=Tidak)";
                    col.ColumnEdit = repoIsActive;
                    col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom;
                    col.DisplayFormat.FormatString = "{0:1;0;0}";
                    col.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    col.Width = 120;
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
            if (daMaster != null && dtMaster != null)
            {
                try
                {
                    SqlCommandBuilder cb = new SqlCommandBuilder(daMaster);
                    daMaster.Update(dtMaster);
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show("Gagal menyimpan ke database SPECTRA_DB: " + ex.Message, "mySAP365 - SPM2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void _btnQSave_ItemClick(object sender, ItemClickEventArgs e)
        {
            _gridViewTD1.PostEditor();
            _gridViewTD1.UpdateCurrentRow();
            SaveDataToDatabase();
            XtraMessageBox.Show("Data Master Kebun MD_ESTATE berhasil disimpan ke database SPECTRA_DB.", "mySAP365 - SPM2", MessageBoxButtons.OK, MessageBoxIcon.Information);
            _gridControlTD1.RefreshDataSource();
            _gridViewTD1.UpdateTotalSummary();
        }

        private void frmSPM2_Idle(object sender, EventArgs e)
        {
            if (this.IsDisposed) return;
        }

        private void _btnQOpen_ItemClick(object sender, ItemClickEventArgs e)
        {
            Syb3dbLoadD(sender);
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
            Syb3InitMRUEditTD1(sender, 50, 50);
        }

        private void _gridViewTD1_ValidatingEditor(object sender, BaseContainerValidateEditorEventArgs e)
        {
        }

        private void frmSPM2_FormClosing(object sender, FormClosingEventArgs e)
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
        }

        private void _gridViewTD1_RowUpdated(object sender, RowObjectEventArgs e)
        {
            SaveDataToDatabase();
        }

        private void _gridViewTD1D_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
        }

        private void _gridViewTD1D_CustomUnboundColumnData(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDataEventArgs e)
        {
        }

        private void _gridViewTD1_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            GridView view1 = (GridView)sender;
            view1.SetRowCellValue(e.RowHandle, "ESTATE_CODE", "KBN");
            view1.SetRowCellValue(e.RowHandle, "ESTATE_NAME", "Kebun Baru");
            view1.SetRowCellValue(e.RowHandle, "COMPANY_NAME", "PT Kayan Plantation");
            view1.SetRowCellValue(e.RowHandle, "IS_ACTIVE", 1);

            if (view1.Columns["ESTATE_CODE"] != null)
            {
                view1.FocusedColumn = view1.Columns["ESTATE_CODE"];
            }
        }

        private void _gridViewTD1_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            if (e.Column.FieldName == "IS_ACTIVE")
            {
                string val = e.CellValue?.ToString() ?? "";
                if (val == "1" || val.Equals("True", StringComparison.OrdinalIgnoreCase))
                {
                    e.Appearance.ForeColor = Color.FromArgb(30, 132, 73);
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
                else
                {
                    e.Appearance.ForeColor = Color.FromArgb(192, 57, 43);
                    e.Appearance.Font = new Font(e.Appearance.Font, FontStyle.Bold);
                }
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
