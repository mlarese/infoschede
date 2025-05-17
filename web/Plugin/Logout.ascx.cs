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
using NextFramework.NextPassport;
using NextFramework.NextB2B;
using NextFramework.NextB2B.DSOrdiniTableAdapters;

/// <summary>
/// Logout.
/// </summary>
public partial class Plugin_AreaRiservata_Logout : NextFramework.NextControls.NextUserControl {

    #region Settings...
    /// <summary>
    /// ID della pagina di post.
    /// </summary>
    public int StnPaginaPostId; 
    #endregion

    /// <summary>
    /// Fa il logout e redirige alla pagina indicata.
    /// </summary>
    /// <param name="e">Parametro di default per questo evento.</param>
    /// <seealso cref="FormsAuthentication.SignOut()"/>
    protected override void OnLoad(EventArgs e) {
        base.OnLoad(e);

        #region Settings init...
        int.TryParse(_settings["paginaPostId"], out StnPaginaPostId);
        #endregion

        // effettua il logout e il reset session della shopping cart
        if (Page.User.Identity.IsAuthenticated)
        {
            FormsAuthentication.SignOut();

            if (StnPaginaPostId > 0)
                Response.Redirect(NextPage.GetPageSiteUrl(StnPaginaPostId));
            else
                Response.Redirect(NextPage.Url);
        }
    }
}
