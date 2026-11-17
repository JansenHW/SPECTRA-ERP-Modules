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
    //-- ENTITAS MASTER DATA KEBUN / ESTATE (dbo.MD_ESTATE)
    //-- DATABASE: SPECTRA_DB | T-CODE: SPM2
    //--------------------------------------------------------------
    [Persistent("MD_ESTATE")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class MD_ESTATE : XPLiteObject
    {
        public MD_ESTATE(Session session) : base(session) { }
        public MD_ESTATE() : base(Session.DefaultSession) { }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            fCOMPANY_NAME = "PT Kayan Plantation";
            fIS_ACTIVE = 1;
        }

        string fESTATE_CODE;
        [Key(false), Persistent("ESTATE_CODE"), Size(20), DbType("varchar(20)")]
        public string ESTATE_CODE
        {
            get { return fESTATE_CODE; }
            set { SetPropertyValue<string>("ESTATE_CODE", ref fESTATE_CODE, value); }
        }

        string fESTATE_NAME;
        [Persistent("ESTATE_NAME"), Size(100), DbType("varchar(100)")]
        public string ESTATE_NAME
        {
            get { return fESTATE_NAME; }
            set { SetPropertyValue<string>("ESTATE_NAME", ref fESTATE_NAME, value); }
        }

        string fCOMPANY_NAME;
        [Persistent("COMPANY_NAME"), Size(150), DbType("varchar(150)")]
        public string COMPANY_NAME
        {
            get { return fCOMPANY_NAME; }
            set { SetPropertyValue<string>("COMPANY_NAME", ref fCOMPANY_NAME, value); }
        }

        int fIS_ACTIVE;
        [Persistent("IS_ACTIVE"), DbType("int")]
        public int IS_ACTIVE
        {
            get { return fIS_ACTIVE; }
            set { SetPropertyValue<int>("IS_ACTIVE", ref fIS_ACTIVE, value); }
        }
    }

    [Persistent("MD_ESTATE")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class spm2 : MD_ESTATE
    {
        public spm2(Session session) : base(session) { }
        public spm2() : base(Session.DefaultSession) { }
    }
}
