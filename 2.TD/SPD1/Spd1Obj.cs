using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text;
using DevExpress.Xpo;
using DevExpress.Xpo.Metadata;
using DevExpress.Data.Filtering;

namespace Syb3DictHVTD
{
    //--------------------------------------------------------------
    //-- ENTITAS TRANSAKSI KUALITAS BUAH TPH (dbo.TPH_TRANSACTION)
    //-- DATABASE: SPECTRA_DB | T-CODE: SPD1
    //--------------------------------------------------------------
    [Persistent("TPH_TRANSACTION")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class TPH_TRANSACTION : XPLiteObject
    {
        public TPH_TRANSACTION(Session session) : base(session) { }
        public TPH_TRANSACTION() : base(Session.DefaultSession) { }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            fESTATE_CODE = "MRE";
            fAFDELING_CODE = "AFD-01";
            fBLOCK_ID = "M101";
            fTPH_NUMBER = "TPH-01";
            fKRANI_NIK = "202315010";
            fHARVEST_DATE = DateTime.Now.Date;
            fMATANG_COUNT = 0;
            fMENTAH_COUNT = 0;
            fBUAH_BUSUK_COUNT = 0;
            fTANGKAI_PANJANG_COUNT = 0;
            fABNORMAL_COUNT = 0;
            fTOTAL_BUAH = 0;
            fSYNC_STATUS = "SYNCED";
            fCREATED_AT = DateTime.Now;
        }

        int fID;
        [Key(true), Persistent("ID"), DbType("int")]
        public int ID
        {
            get { return fID; }
            set { SetPropertyValue<int>("ID", ref fID, value); }
        }

        string fESTATE_CODE;
        [Persistent("ESTATE_CODE"), Size(20), DbType("varchar(20)")]
        public string ESTATE_CODE
        {
            get { return fESTATE_CODE; }
            set { SetPropertyValue<string>("ESTATE_CODE", ref fESTATE_CODE, value); }
        }

        string fAFDELING_CODE;
        [Persistent("AFDELING_CODE"), Size(20), DbType("varchar(20)")]
        public string AFDELING_CODE
        {
            get { return fAFDELING_CODE; }
            set { SetPropertyValue<string>("AFDELING_CODE", ref fAFDELING_CODE, value); }
        }

        string fBLOCK_ID;
        [Persistent("BLOCK_ID"), Size(20), DbType("varchar(20)")]
        public string BLOCK_ID
        {
            get { return fBLOCK_ID; }
            set { SetPropertyValue<string>("BLOCK_ID", ref fBLOCK_ID, value); }
        }

        string fTPH_NUMBER;
        [Persistent("TPH_NUMBER"), Size(50), DbType("varchar(50)")]
        public string TPH_NUMBER
        {
            get { return fTPH_NUMBER; }
            set { SetPropertyValue<string>("TPH_NUMBER", ref fTPH_NUMBER, value); }
        }

        string fKRANI_NIK;
        [Persistent("KRANI_NIK"), Size(50), DbType("varchar(50)")]
        public string KRANI_NIK
        {
            get { return fKRANI_NIK; }
            set { SetPropertyValue<string>("KRANI_NIK", ref fKRANI_NIK, value); }
        }

        DateTime fHARVEST_DATE;
        [Persistent("HARVEST_DATE"), DbType("date")]
        public DateTime HARVEST_DATE
        {
            get { return fHARVEST_DATE; }
            set { SetPropertyValue<DateTime>("HARVEST_DATE", ref fHARVEST_DATE, value); }
        }

        int fMATANG_COUNT;
        [Persistent("MATANG_COUNT"), DbType("int")]
        public int MATANG_COUNT
        {
            get { return fMATANG_COUNT; }
            set { SetPropertyValue<int>("MATANG_COUNT", ref fMATANG_COUNT, value); }
        }

        int fMENTAH_COUNT;
        [Persistent("MENTAH_COUNT"), DbType("int")]
        public int MENTAH_COUNT
        {
            get { return fMENTAH_COUNT; }
            set { SetPropertyValue<int>("MENTAH_COUNT", ref fMENTAH_COUNT, value); }
        }

        int fBUAH_BUSUK_COUNT;
        [Persistent("BUAH_BUSUK_COUNT"), DbType("int")]
        public int BUAH_BUSUK_COUNT
        {
            get { return fBUAH_BUSUK_COUNT; }
            set { SetPropertyValue<int>("BUAH_BUSUK_COUNT", ref fBUAH_BUSUK_COUNT, value); }
        }

        int fTANGKAI_PANJANG_COUNT;
        [Persistent("TANGKAI_PANJANG_COUNT"), DbType("int")]
        public int TANGKAI_PANJANG_COUNT
        {
            get { return fTANGKAI_PANJANG_COUNT; }
            set { SetPropertyValue<int>("TANGKAI_PANJANG_COUNT", ref fTANGKAI_PANJANG_COUNT, value); }
        }

        int fABNORMAL_COUNT;
        [Persistent("ABNORMAL_COUNT"), DbType("int")]
        public int ABNORMAL_COUNT
        {
            get { return fABNORMAL_COUNT; }
            set { SetPropertyValue<int>("ABNORMAL_COUNT", ref fABNORMAL_COUNT, value); }
        }

        int fTOTAL_BUAH;
        [Persistent("TOTAL_BUAH"), DbType("int")]
        public int TOTAL_BUAH
        {
            get { return fTOTAL_BUAH; }
            set { SetPropertyValue<int>("TOTAL_BUAH", ref fTOTAL_BUAH, value); }
        }

        string fPHOTO_PATH;
        [Persistent("PHOTO_PATH"), Size(500), DbType("varchar(500)")]
        public string PHOTO_PATH
        {
            get { return fPHOTO_PATH; }
            set { SetPropertyValue<string>("PHOTO_PATH", ref fPHOTO_PATH, value); }
        }

        string fSYNC_STATUS;
        [Persistent("SYNC_STATUS"), Size(50), DbType("varchar(50)")]
        public string SYNC_STATUS
        {
            get { return fSYNC_STATUS; }
            set { SetPropertyValue<string>("SYNC_STATUS", ref fSYNC_STATUS, value); }
        }

        DateTime fCREATED_AT;
        [Persistent("CREATED_AT"), DbType("datetime")]
        public DateTime CREATED_AT
        {
            get { return fCREATED_AT; }
            set { SetPropertyValue<DateTime>("CREATED_AT", ref fCREATED_AT, value); }
        }
    }

    [Persistent("TPH_TRANSACTION")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class spd1 : TPH_TRANSACTION
    {
        public spd1(Session session) : base(session) { }
        public spd1() : base(Session.DefaultSession) { }
    }
}
