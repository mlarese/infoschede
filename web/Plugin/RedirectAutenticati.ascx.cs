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
using NextFramework.NextWeb;
using NextFramework;
using NextFramework.NextControls;
using NextFramework.NextB2B;
using NextFramework.NextPassport;
using NextFramework.NextB2B.DSOrdiniTableAdapters;
using System.Security.Cryptography;
using System.IO;
using System.Text;

/// <summary>
/// Login dell'area riservata.
/// </summary>
public partial class Plugin_RedirectAutenticati : NextFramework.NextControls.NextUserControl
{
    #region Settings...

    /// <summary>
    /// Pagina di redirect per utente supervisore
    /// </summary>
    int StnSupervisorePaginaRedirectIdx;

    /// <summary>
    /// Pagina di redirect per utente costruttore
    /// </summary>
    int StnCostruttorePaginaRedirectIdx;

    /// <summary>
    /// Pagina di redirect per utente autenticato
    /// </summary>
    int StnPaginaRedirectIdx;

    #endregion

    /// <summary>
    /// Setta le variabili interne e la pagina di destinazione.
    /// </summary>
    /// <param name="e">Parametro di default per questo evento.</param>
    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        #region Settings init...

        StnSupervisorePaginaRedirectIdx = int.Parse(_settings["supervisorePaginaRedirectIdx"]);

        StnCostruttorePaginaRedirectIdx = int.Parse(_settings["costruttorePaginaRedirectIdx"]);

        StnPaginaRedirectIdx = int.Parse(_settings["paginaRedirectIdx"]);

        #endregion

        if (NextMembershipRivenditore.Current != null)
        {
            if (NextMembershipRivenditore.Current.Riv_profilo_id == 1)
            {
                //costruttore
                Response.Redirect(NextPage.Index.BLL.GetUrl(StnCostruttorePaginaRedirectIdx));
            }
            else if (NextMembershipRivenditore.Current.Riv_profilo_id == 5)
            {
                //supervisore
                Response.Redirect(NextPage.Index.BLL.GetUrl(StnSupervisorePaginaRedirectIdx));
            }
            else
            {
                Response.Redirect(NextPage.Index.BLL.GetUrl(StnPaginaRedirectIdx));
            }
        }
        else
        {
            Response.Redirect(NextPage.Index.BLL.GetUrl(StnPaginaRedirectIdx));
        }
    }
}
