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

public partial class Plugin_AlberoCategorie : NextFramework.NextControls.NextUserControl
{
    // iniziali
    DataTable padri_lista;

    #region Settings...

    /// <summary>
    /// Pagina collegata dei prodotti.
    /// </summary>
    public int StnPaginaProdottiId;

    /// <summary>
    /// Categoria base: Modelli.
    /// </summary>
    public int StnCategoriaBaseId;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitolo;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        #region Settings init...

        try { int.TryParse(_settings["paginaProdottiId"], out StnPaginaProdottiId); } catch { }

        try { int.TryParse(_settings["categoriaBaseId"], out StnCategoriaBaseId); } catch { }

        try { StnTitolo = _settings["Titolo"]; } catch { }

        #endregion

        NextPage.IndexAux.BLL.GetIndice(0, NextPage.Index.Id, 0, true, false, false, "", 0, "");
        NextPage.IndexAux.Update();  

        TitoloSez.InnerText = StnTitolo;

        // estrae le iniziali delle marche
        string sql = "SELECT DISTINCT tip_tipologie_padre_lista FROM gtb_articoli" +
                " INNER JOIN gtb_tipologie ON art_tipologia_id = tip_id" +
                " INNER JOIN gtb_marche ON art_marca_id = mar_id" +
                     " WHERE mar_id = " + NextNumeric.ToInt(Request.QueryString["MARID"]) + " AND tip_visibile = 1 AND tip_albero_visibile=1 AND IsNull(art_disabilitato,0)=0";

        // collega alla Nextlist delle iniziali
        padri_lista = new DataTable();
        padri_lista = NextPage.Connection.GetDataTable(sql);

        string padri = "";
        for (int i = 0; i < padri_lista.Rows.Count; i++)
            padri = padri + padri_lista.Rows[i]["tip_tipologie_padre_lista"] + ",";
        if (padri.Length > 0)
            padri.Remove(padri.Length - 1);

        sql = "SELECT DISTINCT tip_id" +
                            ", tip_nome_it" +
                            ", tip_livello" +
                            ", tip_tipologie_padre_lista" +
                            ", (CASE WHEN EXISTS(SELECT top 1 art_id FROM gtb_articoli WHERE art_tipologia_id = gtb_tipologie.tip_id AND art_marca_id = " + NextNumeric.ToInt(Request.QueryString["MARID"]) + " AND IsNull(art_disabilitato,0)=0) THEN 1 ELSE 0 END) AS has_art " +
               " FROM gtb_tipologie" +
              " WHERE '" + padri + "' like '%' + CAST(tip_id AS VARCHAR) + '%' " +
                " AND tip_tipologie_padre_lista like '" + StnCategoriaBaseId + ",%' "+
           " ORDER BY tip_tipologie_padre_lista" +
                   ", tip_nome_it";

        padri_lista = NextPage.Connection.GetDataTable(sql);
        if (padri_lista.Rows.Count > 0)
        {
            Albero.DataSource = padri_lista;
            Albero.DataBind();
        }
    }

    protected void Albero_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            DataRow cat = ((DataRowView)e.Item.DataItem).Row;
            HtmlAnchor link = (HtmlAnchor)e.Item.FindControl("Link");

            link.InnerText = cat["tip_nome_it"].ToString();
            if (NextBoolean.ToBoolean(cat["has_art"]))
                link.HRef = NextControlsTools.AddQueryString(NextPage.IndexAux.Link,
                    "MARID=" + Request.QueryString["MARID"] + "&CATID=" + cat["tip_id"].ToString());
            else
                NextControlsTools.SetCssClass(link, "disabled");
            NextControlsTools.SetCssClass(e.Item, "liv" + cat["tip_livello"].ToString());
        }
    }

    protected void Indietro_Click(object sender, EventArgs e)
    {
        NextPage.IndexAux.BLL.GetIndicePadre(NextPage.Index.Id);
        NextPage.IndexAux.Update();
        Response.Redirect(NextPage.IndexAux.Link);
    }
}