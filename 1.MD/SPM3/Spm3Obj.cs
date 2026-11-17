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
    //-- ENTITAS MASTER DATA AFDELING (dbo.MD_AFDELING)
    //-- DATABASE: SPECTRA_DB | T-CODE: SPM3
    //--------------------------------------------------------------
    [Persistent("MD_AFDELING")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class MD_AFDELING : XPLiteObject
    {
        public MD_AFDELING(Session session) : base(session) { }
        public MD_AFDELING() : base(Session.DefaultSession) { }

        public override void AfterConstruction()
        {
            base.AfterConstruction();
            fESTATE_CODE = "MRE";
        }

        string fAFDELING_CODE;
        [Key(false), Persistent("AFDELING_CODE"), Size(20), DbType("varchar(20)")]
        public string AFDELING_CODE
        {
            get { return fAFDELING_CODE; }
            set { SetPropertyValue<string>("AFDELING_CODE", ref fAFDELING_CODE, value); }
        }

        string fESTATE_CODE;
        [Persistent("ESTATE_CODE"), Size(20), DbType("varchar(20)")]
        public string ESTATE_CODE
        {
            get { return fESTATE_CODE; }
            set { SetPropertyValue<string>("ESTATE_CODE", ref fESTATE_CODE, value); }
        }

        string fAFDELING_NAME;
        [Persistent("AFDELING_NAME"), Size(100), DbType("varchar(100)")]
        public string AFDELING_NAME
        {
            get { return fAFDELING_NAME; }
            set { SetPropertyValue<string>("AFDELING_NAME", ref fAFDELING_NAME, value); }
        }
    }

    [Persistent("MD_AFDELING")]
    [OptimisticLocking(false)]
    [DeferredDeletion(false)]
    public class spm3 : MD_AFDELING
    {
        public spm3(Session session) : base(session) { }
        public spm3() : base(Session.DefaultSession) { }
    }
}
