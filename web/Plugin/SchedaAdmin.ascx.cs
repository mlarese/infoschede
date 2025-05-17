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
using System.Data.SqlClient;
using System.Globalization;
using NextFramework;
using NextFramework.NextWeb;
using NextFramework.NextControls;
using NextFramework.NextPassport;
using NextFramework.NextB2B;
using NextFramework.NextB2B.DSArticoliTableAdapters;

public partial class Plugin_SchedaAdmin : NextFramework.NextControls.NextUserControl
{
    #region Settings...

    /// <summary>
    /// Lista plugin da visualizzare (separati da ',').
    /// </summary>
    public string StnListaPlugin;

    /// <summary>
    /// CssClass da applicare ai plugin.
    /// </summary>
    public string StnCssClassPlugin;

    /// <summary>
    /// Titolo.
    /// </summary>
    public string StnTitolo;

    /// <summary>
    /// Descrizione.
    /// </summary>
    public string StnDescrizione;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        try { StnListaPlugin = _settings["listaPlugin"]; }
        catch { StnListaPlugin = "SchedaDettaglio"; };

        try { StnCssClassPlugin = _settings["cssClassPlugin"]; }
        catch { StnCssClassPlugin = ""; };

        try { StnTitolo = _settings["Titolo"]; } catch { };

        try { StnDescrizione = _settings["Descrizione"]; } catch { };

        #endregion

        base.OnLoad(e);

        int clienteId = 0, schedaId = 0, clienteSchedaId = 0, ddtId = 0;
        NextMembershipRivenditore cliente = new NextMembershipRivenditore();

        if (NextPage.User.Identity.IsAuthenticated)
            FormsAuthentication.SignOut();

        if (Request.QueryString["SCHEDAID"] != null && Request.QueryString["DDTID"] == null &&
            int.TryParse(Request.QueryString["SCHEDAID"].ToString(), out schedaId) && schedaId > 0)
        {
            DataTable dt = InfoschedeTools.GetSchedeDataTable(schedaId, 0, "", "", "");
            if (dt.Rows.Count > 0 && Request.QueryString["CLIENTEID"] != null &&
                int.TryParse(dt.Rows[0]["sc_cliente_id"].ToString(), out clienteSchedaId) && clienteSchedaId > 0 &&
                int.TryParse(Request.QueryString["CLIENTEID"].ToString(), out clienteId) && clienteId > 0 &&
                cliente.SetPropertiesById(clienteSchedaId) && clienteSchedaId == clienteId &&
                Request.QueryString["KEY"] != null && Request.QueryString["IDCNT"] != null &&
                Request.QueryString["IDCNT"].ToString() == cliente.Contatto.Id.ToString() &&
                Request.QueryString["KEY"].ToString() == cliente.Contatto.CodiceInserimento)
            {
                FormsAuthentication.SetAuthCookie(cliente.UserName, false);

                Panel[] panels = NextPage.LoadINextPluginList(StnListaPlugin, StnCssClassPlugin);
                INextPlugin plugin;
                foreach (Panel p in panels)
                {
                    plugin = (INextPlugin)p.Controls[0];
                    if (!String.IsNullOrEmpty(StnTitolo))
                        plugin.Settings["Titolo"] = StnTitolo;
                    if (!String.IsNullOrEmpty(StnDescrizione))
                        plugin.Settings["Descrizione"] = StnDescrizione;
                    BoxDiv.Controls.Add(p);
                }

                FormsAuthentication.SignOut();
            }
            else
            {
                Response.Redirect(NextPage.UrlHomePage);
            }
        }

        else if (Request.QueryString["DDTID"] != null &&
            int.TryParse(Request.QueryString["DDTID"].ToString(), out ddtId) && ddtId > 0)
        {
            DataTable dt = InfoschedeTools.GetDdtDataTable(ddtId, 0, 0, "", "");
            if (dt.Rows.Count > 0 && Request.QueryString["CLIENTEID"] != null &&
                //int.TryParse(dt.Rows[0]["ddt_trasportatore_id"].ToString(), out clienteSchedaId) && clienteSchedaId > 0 &&
                int.TryParse(Request.QueryString["CLIENTEID"].ToString(), out clienteId) && clienteId > 0 &&
                cliente.SetPropertiesById(clienteId) &&
                Request.QueryString["KEY"] != null && Request.QueryString["IDCNT"] != null &&
                Request.QueryString["IDCNT"].ToString() == cliente.Contatto.Id.ToString() &&
                Request.QueryString["KEY"].ToString() == cliente.Contatto.CodiceInserimento)
            {
                FormsAuthentication.SetAuthCookie(cliente.UserName, false);
                
                Panel[] panels = NextPage.LoadINextPluginList(StnListaPlugin, StnCssClassPlugin);
                INextPlugin plugin;
                foreach (Panel p in panels)
                {
                    plugin = (INextPlugin)p.Controls[0];
                    if (!String.IsNullOrEmpty(StnTitolo))
                        plugin.Settings["Titolo"] = StnTitolo;
                    if (!String.IsNullOrEmpty(StnDescrizione))
                        plugin.Settings["Descrizione"] = StnDescrizione;
                    BoxDiv.Controls.Add(p);
                }

                FormsAuthentication.SignOut();
            }
            else
            {
                Response.Redirect(NextPage.UrlHomePage);
            }
        }
        else
        {
            Response.Redirect(NextPage.UrlHomePage);
        }
    }
}