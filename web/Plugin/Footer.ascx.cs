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
using NextFramework.NextB2B;
using NextFramework.NextControls;

public partial class Plugin_Footer : NextFramework.NextControls.NextUserControl {

    #region Settings...

    /// <summary>
    /// Url del logo da inserire nel footer
    /// </summary>
    public string StnLogo;

    /// <summary>
    /// Nome aziendale ed indirizzo
    /// </summary>
    public string StnIndirizzo;

    /// <summary>
    /// Lista dei plugin da caricare
    /// </summary>
    public string StnListaPlugin;

    /// <summary>
    /// ID del raggruppamento padre a cui appartengono gli altri
    /// </summary>
    public int StnPadreRaggruppamentoID;

    #endregion

    /// <summary>
    /// </summary>
    /// <param name="e">Parametro di default.</param>
    protected override void OnLoad (EventArgs e) {

        #region Settings init...

        StnLogo = _settings["Logo"];
        StnIndirizzo = _settings["Indirizzo"];
        StnPadreRaggruppamentoID = int.Parse(_settings["PadreRaggruppamentoID"]);

        StnListaPlugin = _settings["ListaPlugin"];


        #endregion

        base.OnLoad(e);

        //imposta logo
        Logo.ImageUrl = NextPage.UrlImages + StnLogo.TrimStart('/');
        Home.NavigateUrl = NextPage.GetPageSiteUrl(NextApplication.PageDefault);
        //imposta indirizzo
        Indirizzo.InnerHtml = StnIndirizzo;
        
        //genera menu        
        DataTable dt = NextPage.Index.BLL.GetIndiceDataTable(0, 0, StnPadreRaggruppamentoID,NextPage.IndexAux.BLL.Contenuto.Tabella.GetIndiceTabellaId("", BLLIndiceTabella.TabellaNome.Raggruppamenti), true, false, false, "");
        ListaMenu.DataSource = dt;
        ListaMenu.DataBind();


        // Genero i plugin nel pannello DivMoreBox
        char[] splitter  = {','};

        string[] plugs = StnListaPlugin.Split(splitter);
        foreach (string identifObjects in plugs)
        {
            try
            {
                HtmlGenericControl box = new HtmlGenericControl("div");
                box.Attributes.Add("class","divmoreelement");
                Panel temp = NextPage.LoadINextPlugin(identifObjects);
                box.Controls.Add(temp);
                DivMoreBox.Controls.Add(box);
            }
            catch
            {
                
            }
        }

    }
    

    protected void ListaMenu_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        HtmlGenericControl title = (HtmlGenericControl)e.Item.FindControl("Titolo");
        NextPage.IndexAux.BLL.SetProperties(((DataRowView)e.Item.DataItem).Row);
        NextPage.IndexAux.Update();
        title.InnerText = NextPage.IndexAux.Titolo;

        NextList menu = (NextList)e.Item.FindControl("menu");        
        DataTable dt = NextPage.Index.BLL.GetIndiceDataTable(0, 0, NextPage.IndexAux.Id, NextPage.IndexAux.BLL.Contenuto.Tabella.GetIndiceTabellaId("", BLLIndiceTabella.TabellaNome.Pagine), true, false, false, "");
        menu.DataSource = dt;
        menu.DataBind();
    }


    protected void menu_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        HtmlAnchor link = (HtmlAnchor)e.Item.FindControl("Link");
        NextPage.IndexAux.BLL.SetProperties(((DataRowView)e.Item.DataItem).Row);
        NextPage.IndexAux.Update();
        link.InnerText = NextPage.IndexAux.Titolo;
        link.HRef = NextPage.IndexAux.Link;

    }
}
