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
    //-- ENTITAS MASTER DATA BLOK (dbo.MD_BLOCK)
    //-- DATABASE: SPECTRA_DB | T-CODE: SPM4
    //--------------------------------------------------------------
    [Persistent("MD_BLOCK")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class MD_BLOCK : XPLiteObject
    {
        public MD_BLOCK(Session session) : base(session) { }
        public MD_BLOCK() : base(Session.DefaultSession) { }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            fESTATE_CODE = "MRE";
            fAFDELING_CODE = "AFD-01";
            fBLOCK_NUMBER = "01";
            fIS_ACTIVE = 1;
        }

        string fBLOCK_ID;
        [Key(false), Persistent("BLOCK_ID"), Size(20), DbType("varchar(20)")]
        public string BLOCK_ID
        {
            get { return fBLOCK_ID; }
            set { SetPropertyValue<string>("BLOCK_ID", ref fBLOCK_ID, value); }
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

        string fBLOCK_NUMBER;
        [Persistent("BLOCK_NUMBER"), Size(20), DbType("varchar(20)")]
        public string BLOCK_NUMBER
        {
            get { return fBLOCK_NUMBER; }
            set { SetPropertyValue<string>("BLOCK_NUMBER", ref fBLOCK_NUMBER, value); }
        }

        int fIS_ACTIVE;
        [Persistent("IS_ACTIVE"), DbType("int")]
        public int IS_ACTIVE
        {
            get { return fIS_ACTIVE; }
            set { SetPropertyValue<int>("IS_ACTIVE", ref fIS_ACTIVE, value); }
        }
    }

    [Persistent("MD_BLOCK")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class spm4 : MD_BLOCK
    {
        public spm4(Session session) : base(session) { }
        public spm4() : base(Session.DefaultSession) { }
    }
}
