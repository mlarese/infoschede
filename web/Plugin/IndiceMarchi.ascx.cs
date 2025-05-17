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

public partial class Plugin_IndiceMarchi : NextFramework.NextControls.NextUserControl
{
    // iniziali
    DataTable inizialiDt = new DataTable();

    // marchi
    DataTable marchiDt = new DataTable();

    #region Settings...

    /// <summary>
    /// Pagina collegata delle categorie.
    /// </summary>
    public int StnPaginaCategorieId;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitolo;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        #region Settings init...

        int.TryParse(_settings["paginaCategorieId"], out StnPaginaCategorieId);

        StnTitolo = _settings["Titolo"];

        #endregion        

        NextPage.IndexAux.BLL.GetIndice(0, NextPage.Index.Id, 0, true, false, false, "", 0, "");
        NextPage.IndexAux.Update();        

        TitoloSez.InnerText = StnTitolo;

        // estrae le iniziali delle marche
        string sql = "SELECT DISTINCT SUBSTRING(gtb_marche.mar_nome_it, 1, 1) AS iniz" + 
                      " FROM gtb_marche" +
                     " WHERE EXISTS(SELECT art_id" +
                                    " FROM gtb_articoli" +
                                   " WHERE art_marca_id = gtb_marche.mar_id)" +
                  " ORDER BY iniz";

        // collega alla Nextlist delle iniziali
        inizialiDt = NextPage.Connection.GetDataTable(sql);
        if (inizialiDt.Rows.Count > 0)
        {
            ListaLettere.DataSource = inizialiDt;
            ListaLettere.DataBind();
        }
    }

    protected void ListaLettere_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            // recupera dati
            DataRow inizRow = ((DataRowView)e.Item.DataItem).Row;

            // lettera iniziale
            HtmlGenericControl inizMarchio = (HtmlGenericControl)e.Item.FindControl("iniziale");
            inizMarchio.InnerText = inizRow[0].ToString();

            // lista marchi
            NextList marchi = (NextList)e.Item.FindControl("ListaMarchi");

            // collega alla Nextlist dei singoli marchi
            string sql = "SELECT *" +
                              ", 0 AS tipo" +
                          " FROM gtb_marche" +
                         " WHERE SUBSTRING(mar_nome_it, 1, 1)='" + inizRow[0].ToString() + "'" +
                           " AND COALESCE(mar_logo, '') <> '' " +
                           " AND mar_generica = 0" +
                         " UNION ALL" +
                        " SELECT *" +
                              ", 1 AS tipo" +
                          " FROM gtb_marche" +
                         " WHERE SUBSTRING(mar_nome_it, 1, 1)='" + inizRow[0].ToString() + "'" +
                           " AND COALESCE(mar_logo, '') = '' " +
                           " AND mar_generica = 0" +
                      " ORDER BY tipo" +
                              ", mar_nome_it";

            marchiDt = NextPage.Connection.GetDataTable(sql);
            if (marchiDt.Rows.Count > 0)
            {
                marchi.DataSource = marchiDt;
                marchi.DataBind();
            }
            else
                e.Item.Visible = false;
        }
    }

    protected void ListaMarchi_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            // recupera dati
            DataRow marca = ((DataRowView)e.Item.DataItem).Row;
            
            // singolo marchio
            HtmlImage logo = (HtmlImage)e.Item.FindControl("logo");
            HtmlAnchor nomeMarchio = (HtmlAnchor)e.Item.FindControl("marchio");
            HtmlGenericControl titolo = (HtmlGenericControl)e.Item.FindControl("Titolo");
            HtmlGenericControl cont = (HtmlGenericControl)e.Item.FindControl("ImgCont");

            if(string.IsNullOrEmpty(marca["mar_logo"].ToString()))
            {
                logo.Visible = false;
                titolo.InnerText = marca["mar_nome_it"].ToString();
                nomeMarchio.HRef = NextControlsTools.AddQueryString(NextPage.IndexAux.Link, "MARID=" + marca["mar_id"].ToString());
            }
            else
            {
                titolo.Visible = false;
                logo.Src = NextPage.UrlImages + marca["mar_logo"].ToString().TrimStart('/');
                logo.Alt = marca["mar_nome_it"].ToString();
                nomeMarchio.Title = marca["mar_nome_it"].ToString();
                nomeMarchio.HRef = NextControlsTools.AddQueryString(NextPage.IndexAux.Link, "MARID=" + marca["mar_id"].ToString());
            }            
        }
    }
}