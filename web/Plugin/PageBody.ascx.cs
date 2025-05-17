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
using NextFramework.NextControls;

public partial class Plugin_PageBody : NextFramework.NextControls.NextUserControl
{
    #region Settings...

    /// <summary>
    /// Nome del plugin "colonna destra".
    /// </summary>
    public string StnNameColumnSxPlugin;

    /// <summary>
    /// Nome del plugin principale da caricare
    /// </summary>
    public string StnNameMainPlugin;

    /// <summary>
    /// Nome del plugin delle briciole da caricare
    /// </summary>
    public string StnNameBriciolePlugin;

    /// <summary>
    /// Testo del footer
    /// </summary>
    public string StnFooterText;

    /// <summary>
    /// Link del powered by
    /// </summary>
    public string StnPoweredByHref;

    /// <summary>
    /// Testo del powered by
    /// </summary>
    public string StnPoweredByText;

    /// <summary>
    /// Title del link powered by
    /// </summary>
    public string StnPoweredByTitle;

    /// <summary>
    /// Margine bottom del plugin briciole.
    /// </summary>
    public int StnBricioleMarginBottom;

    #endregion
    
    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        StnNameColumnSxPlugin = _settings["NameColumnSxPlugin"];
        StnNameMainPlugin = _settings["NameMainPlugin"];
        StnNameBriciolePlugin = _settings["NameBriciolePlugin"];

        try { int.TryParse(_settings["bricioleMarginBottom"], out StnBricioleMarginBottom); }
        catch { StnBricioleMarginBottom = 0; }

        StnFooterText = _nextLanguage.ChooseString(_settings, "FooterText");
        StnPoweredByHref = _nextLanguage.ChooseString(_settings, "PoweredByHref");
        StnPoweredByText = _nextLanguage.ChooseString(_settings, "PoweredByText");
        StnPoweredByTitle = _nextLanguage.ChooseString(_settings, "PoweredByTitle");

        #endregion
        
        NextControlsTools.SetCssClass(this.Layer, "footedbody pagebody");
        Panel pluginDiv;
        Panel[] plugins_list_Div;

        this.ZOrder = 0;

        // Plugin per le briciole
        if (!string.IsNullOrEmpty(StnNameBriciolePlugin))
        {
            pluginDiv = NextPage.LoadINextPlugin(StnNameBriciolePlugin);
            if (StnBricioleMarginBottom != 0)
            {
                pluginDiv.Style.Add(HtmlTextWriterStyle.MarginBottom, StnBricioleMarginBottom + "px");
            }
            BriciolePanel.Controls.Add(pluginDiv);
        }

        // Plugin per il pannello centrale
        if (!string.IsNullOrEmpty(StnNameMainPlugin))
        {
            plugins_list_Div = NextPage.LoadINextPluginList(StnNameMainPlugin, "");
            for (int i = 0; i < plugins_list_Div.Length; i++)
            {
                MainPanel.Controls.Add(plugins_list_Div[i]);
            }
            NextControlsTools.SetCssClass(MainPanel, "container_" + StnNameMainPlugin.ToLower());

            //imposta dimensione minima
            MainPanel.Style.Add("min-height", this.Height + "em");
        }
        else
        {
            //imposta dimensione all'oggetto.
            BackgroundInterno.Style.Add(HtmlTextWriterStyle.Height, this.Height + "em");
        }

        // Plugin per la colonna sinistra 
        if (!string.IsNullOrEmpty(StnNameColumnSxPlugin))
        {
            plugins_list_Div = NextPage.LoadINextPluginList(StnNameColumnSxPlugin, "");
            for (int i = 0; i < plugins_list_Div.Length; i++)
            {
                ColumnSxPanel.Controls.Add(plugins_list_Div[i]);
            }
            NextControlsTools.SetCssClass(ColumnSxPanel, "container_" + StnNameColumnSxPlugin.ToLower());

            //imposta dimensione minima
            ColumnSxPanel.Style.Add("min-height", this.Height + "em");
        }  


        FooterText.InnerText = StnFooterText;
        PoweredBy.InnerText = StnPoweredByText;
        PoweredBy.HRef = StnPoweredByHref;
        PoweredBy.Title = StnPoweredByTitle;
        PoweredBy.Target = StnPoweredByText;

    }
}
