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
    //-- ENTITAS MASTER DATA PENGGUNA / KARYAWAN (dbo.MD_EMPLOYEE)
    //-- DATABASE: SPECTRA_DB | T-CODE: SPM1
    //--------------------------------------------------------------
    [Persistent("MD_EMPLOYEE")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class MD_EMPLOYEE : XPLiteObject
    {
        public MD_EMPLOYEE(Session session) : base(session) { }
        public MD_EMPLOYEE() : base(Session.DefaultSession) { }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            fPOSITION = "Kerani Panen";
            fESTATE_CODE = "MRE";
            fAFDELING_CODE = "AFD-01";
            fIS_ACTIVE = 1;
            fCREATED_AT = DateTime.Now;
        }

        string fNIK;
        [Key(false), Persistent("NIK"), Size(50), DbType("varchar(50)")]
        public string NIK
        {
            get { return fNIK; }
            set { SetPropertyValue<string>("NIK", ref fNIK, value); }
        }

        string fFULL_NAME;
        [Persistent("FULL_NAME"), Size(150), DbType("varchar(150)")]
        public string FULL_NAME
        {
            get { return fFULL_NAME; }
            set { SetPropertyValue<string>("FULL_NAME", ref fFULL_NAME, value); }
        }

        string fPOSITION;
        [Persistent("POSITION"), Size(50), DbType("varchar(50)")]
        public string POSITION
        {
            get { return fPOSITION; }
            set { SetPropertyValue<string>("POSITION", ref fPOSITION, value); }
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

        int fIS_ACTIVE;
        [Persistent("IS_ACTIVE"), DbType("int")]
        public int IS_ACTIVE
        {
            get { return fIS_ACTIVE; }
            set { SetPropertyValue<int>("IS_ACTIVE", ref fIS_ACTIVE, value); }
        }

        DateTime fCREATED_AT;
        [Persistent("CREATED_AT"), DbType("datetime")]
        public DateTime CREATED_AT
        {
            get { return fCREATED_AT; }
            set { SetPropertyValue<DateTime>("CREATED_AT", ref fCREATED_AT, value); }
        }
    }

    [Persistent("MD_EMPLOYEE")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class spm1 : MD_EMPLOYEE
    {
        public spm1(Session session) : base(session) { }
        public spm1() : base(Session.DefaultSession) { }
    }
}
