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

public partial class Plugin_ProcessoAssistenza : NextFramework.NextControls.NextUserControl
{   
    #region Settings...

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloMarchi;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloCategorie;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloModelli;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloGuasti;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloRegistrazione;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloAssistenza;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnTitoloAssistenzaInviata;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescMarchi;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescCategorie;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescModelli;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescGuasti;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescRegistrazione;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescAssistenza;

    /// <summary>
    /// Titolo sezione.
    /// </summary>
    public string StnDescAssistenzaInviata;

    /// <summary>
    /// Id pagina assistenza.
    /// </summary>
    public int StnPaginaAssistenzaIdx;

    /// <summary>
    /// Id pagina assistenza inviata.
    /// </summary>
    public int StnPaginaAssistenzaInviataIdx;

    /// <summary>
    /// Id del capostipite.
    /// </summary>
    public int StnPadreIdx;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        StnTitoloMarchi = _settings["TitoloMarchi"];
        StnTitoloCategorie = _settings["TitoloCategorie"];
        StnTitoloModelli = _settings["TitoloModelli"];
        StnTitoloGuasti = _settings["TitoloGuasti"];
        StnTitoloRegistrazione = _settings["TitoloRegistrazione"];
        StnTitoloAssistenza = _settings["TitoloAssistenza"];
        StnTitoloAssistenzaInviata = _settings["TitoloAssistenzaInviata"];

        StnDescMarchi = _settings["DescMarchi"];
        StnDescCategorie = _settings["DescCategorie"];
        StnDescModelli = _settings["DescModelli"];
        StnDescGuasti = _settings["DescGuasti"];
        StnDescRegistrazione = _settings["DescRegistrazione"];
        StnDescAssistenza = _settings["DescAssistenza"];
        StnDescAssistenzaInviata = _settings["DescAssistenzaInviata"];

        int.TryParse(_settings["paginaAssistenzaIdx"], out StnPaginaAssistenzaIdx);
        int.TryParse(_settings["paginaAssistenzaInviataIdx"], out StnPaginaAssistenzaInviataIdx);        
        int.TryParse(_settings["padreIdx"], out StnPadreIdx);        

        #endregion

        base.OnLoad(e);

        NextControlsTools.SetCssClass(this.Layer, "processoassistenza");

        DataTable passi = new DataTable();
        DataTable tmp = new DataTable();

        passi = NextPage.Index.BLL.GetIndiceDataTable(0, 0, 0, 0, false, false, false, "",0,"idx_livello",false,false,false,0,
            false, false, " idx_tipologie_padre_lista like '%" + StnPadreIdx.ToString() + "%' ");
        tmp = NextPage.Index.BLL.GetIndiceDataTable(0, 0, 0, 0, false, false, false, "", 0, "idx_livello",false,false,false,0,
            false, false, " idx_tipologie_padre_lista like '%" + StnPaginaAssistenzaIdx.ToString() + "%' ");

        passi.Merge(tmp);        

        Processo.DataSource = passi;
        Processo.DataBind();
    }

    protected void Processo_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {      
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            DataRow passo = ((DataRowView)e.Item.DataItem).Row;

            int idx = int.Parse(passo["idx_id"].ToString());
            int livello = int.Parse(passo["idx_livello"].ToString());

            if (NextPage.Index.Id == idx)
                NextControlsTools.SetCssClass(e.Item, "selected");

            HtmlGenericControl pos = (HtmlGenericControl)e.Item.FindControl("PassoP");
            HtmlGenericControl titolo = (HtmlGenericControl)e.Item.FindControl("TitoloH1");
            HtmlGenericControl desc = (HtmlGenericControl)e.Item.FindControl("DescrizioneP");
            pos.InnerText = "passo " + (e.Item.ItemIndex + 1);

            if (idx == StnPaginaAssistenzaIdx)
            {
                titolo.InnerText = StnTitoloAssistenza;
                desc.InnerText = StnDescAssistenza;
            }
            else if (idx == StnPaginaAssistenzaInviataIdx)
            {
                titolo.InnerText = StnTitoloAssistenzaInviata;
                desc.InnerText = StnDescAssistenzaInviata;
            }
            else if (livello == 2)
            {
                titolo.InnerText = StnTitoloMarchi;
                desc.InnerText = StnDescMarchi;
            }
            else if (livello == 3)
            {
                titolo.InnerText = StnTitoloCategorie;
                desc.InnerText = StnDescCategorie;
            }
            else if (livello == 4)
            {
                titolo.InnerText = StnTitoloModelli;
                desc.InnerText = StnDescModelli;
            }
            else if (livello == 5)
            {
                titolo.InnerText = StnTitoloGuasti;
                desc.InnerText = StnDescGuasti;
            }
            else if (livello == 6)
            {
                titolo.InnerText = StnTitoloRegistrazione;
                desc.InnerText = StnDescRegistrazione;
            }
        }
    }
}