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

public partial class Plugin_ElencoProdotti : NextFramework.NextControls.NextUserControl
{
    // articoli
    DataTable art_lista = new DataTable();

    #region Settings...

    /// <summary>
    /// Pagina collegata dei guasti.
    /// </summary>
    public int StnPaginaGuastiId;

    /// <summary>
    /// Pagina collegata dei guasti.
    /// </summary>
    public int StnPaginaAssistenzaId;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitolo;

    /// <summary>
    /// Profilo utente. Se =0 prende quello del profilo corrente
    /// </summary>
    public int StnProfiloId;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        int.TryParse(_settings["profiloId"], out StnProfiloId);

        int.TryParse(_settings["paginaGuastiId"], out StnPaginaGuastiId);

        int.TryParse(_settings["paginaAssistenzaId"], out StnPaginaAssistenzaId);

        StnTitolo = _settings["Titolo"];

        #endregion

        base.OnLoad(e);

        if (Session["filtroProdotti"] == null)
            Session["filtroProdotti"] = "";

        NextControlsTools.SetCssClass(this.Layer, "elencoprodotti");

        TitoloSez.InnerText = StnTitolo;

        if (StnProfiloId == 0)
            StnProfiloId = NextMembershipRivenditore.Current.Riv_profilo_id.Value;

        // estrae gli articoli
        string sql = "SELECT art_nome_it" +
                          ", art_id, 0 AS ordine " +
                      " FROM gtb_articoli" +
                     " WHERE art_marca_id = " + NextNumeric.ToInt(Request.QueryString["MARID"]) +
                       " AND art_disabilitato = 0 " +
                       " AND (art_tipologia_id = " + NextNumeric.ToInt(Request.QueryString["CATID"]) +
                         " OR art_tipologia_id IN (SELECT tip_id" +
                                                   " FROM gtb_tipologie " +
                                                  " WHERE tip_tipologie_padre_lista like '" + NextNumeric.ToInt(Request.QueryString["CATID"]) + "%')) " +
                     Session["filtroProdotti"].ToString() +
              " UNION SELECT art_nome_it" +
                          ", art_id, 1 as ordine " +
                      " FROM gtb_articoli" +
                     " WHERE art_cod_int = '" + InfoschedeTools.CodModelloDefault + "'" + 
                     " ORDER BY ordine, art_nome_it ";
    
        // collega alla Nextlist delle iniziali
        art_lista = NextPage.Connection.GetDataTable(sql);
        if (art_lista.Rows.Count > 0)
        {
            Prodotti.DataSource = art_lista;
            Prodotti.DataBind();
        }
    }

    protected void Prodotti_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            DataRow dr = ((DataRowView)e.Item.DataItem).Row;
            HtmlAnchor link = (HtmlAnchor)e.Item.FindControl("Link");
            link.InnerText = dr["art_nome_it"].ToString();

            if (dr["art_id"].ToString() == InfoschedeTools.GetModelloDefaultId().ToString())
            {
                HtmlGenericControl avviso = new HtmlGenericControl("div");
                avviso.InnerText = "Se non hai trovato il modello che cerchi, clicca su 'altro modello' :";
                e.Item.Controls.AddAt(0, avviso);
                link.HRef = NextString.QueryStringAdd(NextPage.GetPageSiteUrl(StnPaginaAssistenzaId),
                                                      "PROBID=0" +
                                                      "&ARTID=" + dr["ART_id"].ToString() +
                                                      "&MARID=" + Request.QueryString["MARID"]);
            }
            else
                link.HRef = NextString.QueryStringAdd(NextPage.GetPageSiteUrl(StnPaginaGuastiId),
                                                      "ARTID=" + dr["ART_id"].ToString() +
                                                      "&MARID=" + Request.QueryString["MARID"] +
                                                      "&CATID=" + Request.QueryString["CATID"]);
        }
    }
    
    protected void Indietro_Click(object sender, EventArgs e)
    {
        NextPage.IndexAux.BLL.GetIndicePadre(NextPage.Index.Id);
        NextPage.IndexAux.Update();
        Response.Redirect(NextString.QueryStringAdd(NextPage.IndexAux.Link, "MARID=" + Request.QueryString["MARID"]));
    }
}