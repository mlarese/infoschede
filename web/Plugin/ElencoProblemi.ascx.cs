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
using NextFramework.NextControls;
using NextFramework.NextB2B;
using NextFramework.NextB2B.DSArticoliTableAdapters;
using NextFramework.NextWeb;
using NextFramework.NextPassport;
using NextFramework.NextTools;


public partial class Plugin_ElencoProblemi : NextFramework.NextControls.NextUserControl
{
    // articoli
    DataTable prob_lista = new DataTable();

    #region Settings...

    /// <summary>
    /// Pagina collegata dei guasti.
    /// </summary>
    public int StnPaginaRegistrazioneId;

    /// <summary>
    /// Pagina che visualizza un messaggio per il caso in cui il problema sia stato risolto.
    /// </summary>
    public int StnPaginaRisoltoId;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitolo;

    /// <summary>
    /// Testo domanda assistenza o risolto.
    /// </summary>
    public string StnDomanda;

    /// <summary>
    /// Profilo collegato. Se = 0 prende quello dell'utente corrente (per autenticati)
    /// </summary>
    public int StnProfiloId;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        #region Settings init...

        try { int.TryParse(_settings["paginaRegistrazioneId"], out StnPaginaRegistrazioneId); }catch { };

        try { int.TryParse(_settings["paginaRisoltoId"], out StnPaginaRisoltoId); } catch { };

        try { StnTitolo = _settings["titolo"]; } catch { };

        try { StnDomanda = _settings["domanda"]; } catch { };

        try { int.TryParse(_settings["profiloId"], out StnProfiloId); } catch { };

        #endregion

        NextControlsTools.SetCssClass(this.Layer, "elencoproblemi");

        TitoloSez.InnerText = StnTitolo;

        if (StnProfiloId == 0)
            StnProfiloId = NextMembershipRivenditore.Current.Riv_profilo_id.Value;
        
        // estrae i problemi
        string sql = "SELECT DISTINCT prb_nome_it" +
                                   ", prb_id" +
                                   ", prb_modalita_easy" +
                                   ", CAST(prb_avviso_per_conferma_it AS NVarchar(4000)) AS soluzione" +
                                   ", 0 AS ordinamento" +
                      " FROM sgtb_problemi" +
                 " LEFT JOIN srel_problemi_articoli ON prb_id = rpa_problema_id" +
                 " LEFT JOIN srel_problemi_mar_tip ON prb_id = rpm_problema_id" +
                 " LEFT JOIN srel_problemi_profili ON prb_id = rpp_problema_id" +
                 " LEFT JOIN grel_art_valori ON rel_id = rpa_articolo_rel_id" +
                     " WHERE prb_riscontrato = 0" +
                       " AND prb_visibile = 1" +
                       " AND rpp_problema_id = prb_id AND rpp_profilo_id = " + StnProfiloId.ToString() +
                       " AND ((rpa_problema_id = prb_id AND rel_art_id =" + NextNumeric.ToInt(Request.QueryString["ARTID"]) + ")" +
                         " OR (rpm_problema_id = prb_id" +
                       " AND (rpm_tipologia_id = " + NextNumeric.ToInt(Request.QueryString["CATID"]) + " OR rpm_tipologia_id = 0)" +
                       " AND (rpm_marchio_id = " + NextNumeric.ToInt(Request.QueryString["MARID"]) + " OR rpm_marchio_id = 0)))" +
                     // altro problema
                     " UNION" +
                    " SELECT 'altro problema' AS prb_nome_it" +
                          ", 0 AS prb_id" +
                          ", CAST(1 AS BIT) AS prb_modalita_easy" +
                          ", '' AS soluzione" +
                          ", 1 AS ordinamento" +
                  " ORDER BY ordinamento";
        
        prob_lista = NextPage.Connection.GetDataTable(sql);

        if (prob_lista.Rows.Count > 1)
        {
            Problemi.DataSource = prob_lista;
            Problemi.DataBind();
        }
        else
            Response.Redirect(NextString.QueryStringAdd(NextPage.GetPageSiteUrl(StnPaginaRegistrazioneId),
                                                        "PROBID=0&ARTID=" + Request.QueryString["ARTID"]));

        #region script espandi soluzione
        string jqueryScript = "\n" +
                              "$(\"div.faq_pnl\").click( \n" +
                              "    function() {\n" +
                              "        $(\"div#\" + this.id).toggleClass(\"faqexpanded\",true); \n" +
                              "        }\n" +
                              "); \n";
        NextPage.JQueryReadyManager.AddReadyScript(jqueryScript);
        #endregion
    }

    protected void Problemi_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            DataRow dr = ((DataRowView)e.Item.DataItem).Row;

            Button ok = (Button)e.Item.FindControl("Ok");
            Button assistenza = (Button)e.Item.FindControl("Assistenza");
            ok.Attributes.Add("prob", dr["prb_id"].ToString());
            assistenza.Attributes.Add("prob", dr["prb_id"].ToString());
            
            HtmlAnchor link = (HtmlAnchor)e.Item.FindControl("Link");
            HtmlGenericControl sol = (HtmlGenericControl)e.Item.FindControl("Soluzioni");
            HtmlGenericControl cont = (HtmlGenericControl)e.Item.FindControl("Container");
            HtmlGenericControl domanda = (HtmlGenericControl)e.Item.FindControl("Domanda");
            link.InnerText = dr["prb_nome_it"].ToString();
            domanda.InnerText = StnDomanda;

            if (bool.Parse(dr["prb_modalita_easy"].ToString()))
            {
                cont.Visible = false;
                NextControlsTools.SetCssClass(link, "form");
                if (dr["prb_nome_it"].ToString() == "altro problema")
                {
                    HtmlGenericControl avviso = new HtmlGenericControl("div");
                    avviso.InnerText = "Se non hai trovato il problema riscontrato, clicca su 'altro problema' :";
                    e.Item.Controls.AddAt(0, avviso);
                }
                link.HRef = NextString.QueryStringAdd(NextPage.GetPageSiteUrl(StnPaginaRegistrazioneId),
                                                      "PROBID=" + dr["prb_id"].ToString() +
                                                      "&ARTID=" + Request.QueryString["ARTID"]);
            }
            else
            {
                NextControlsTools.SetCssClass(link, "info");
                sol.InnerHtml = dr["soluzione"].ToString();
            }
        }
    }

    protected void Ok_Clicked(Object sender, EventArgs e)
    {
        Button button = (Button)sender;
        NextFrameworkLog.InsertLog("guasto", int.Parse(button.Attributes["prob"]), "problema_risolto",
                                   "problema risolto autonomamente con il sistema di assistenza guidata", 38);
        Response.Redirect(NextPage.GetPageSiteUrl(StnPaginaRisoltoId));
    }

    protected void Assistenza_Clicked(Object sender, EventArgs e)
    {
        Button button = (Button)sender;
        Response.Redirect(NextString.QueryStringAdd(NextPage.GetPageSiteUrl(StnPaginaRegistrazioneId),
                                                    "PROBID=" + button.Attributes["prob"] +
                                                    "&ARTID=" + Request.QueryString["ARTID"]));
    }

    protected void Indietro_Click(object sender, EventArgs e)
    {
        NextPage.IndexAux.BLL.GetIndicePadre(NextPage.Index.Id);
        NextPage.IndexAux.Update();
        Response.Redirect(NextString.QueryStringAdd(NextPage.IndexAux.Link, "MARID=" + Request.QueryString["MARID"] +
                                                                            "&CATID=" + Request.QueryString["CATID"]));
    }
}