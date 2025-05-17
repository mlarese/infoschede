using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using NextFramework;
using NextFramework.NextWeb;
using NextFramework.Messaggi;
using NextFramework.NextPassport;
using NextFramework.NextB2B;
using NextPdfTools;

public partial class Plugin_Email : NextFramework.NextControls.NextUserControl
{
    #region Settings...

    #endregion
    
    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        #endregion

        base.OnLoad(e);

        if (!string.IsNullOrEmpty(Request.QueryString["TESTO"]) && Request.QueryString["TESTO"].ToString() != "")
            EmailTesto.InnerHtml = NextString.HtmlEncode(Request.QueryString["TESTO"].ToString());
        else
            this.Layer.Visible = false;
    }
}
