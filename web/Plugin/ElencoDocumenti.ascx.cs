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
using NextFramework.NextControls;

/// <summary>
/// Magiclist che controlla la visibilità dei documenti del NextMemo.
/// </summary>
public partial class Plugin_B2B_ElencoDocumenti : NextFramework.NextControls.NextUserControl
{

    /// <summary>
    /// Controlla la visibilità dei documenti del NextMemo.
    /// </summary>
    /// <param name="sender">Magicbox sender.</param>
    /// <param name="e">Parametro di default.</param>
    protected void Catalogo_MagicboxDataBinding(object sender, EventArgs e) {
        Magicbox magicbox = (Magicbox)sender;        
        magicbox.DataSource.Link_it = NextPage.UrlBase + magicbox.DataSource.Link_it;
        magicbox.DataSource.LinkRW_it = NextPage.UrlBase + magicbox.DataSource.LinkRW_it;

        if (magicbox.DataSource.Contenuto.Tabella.Nome.Equals("mtb_documenti", StringComparison.InvariantCultureIgnoreCase))
            using (NextFramework.NextMemo2.BLLDocumento documento = new NextFramework.NextMemo2.BLLDocumento(_nextPage.Connection)) 
            {
                magicbox.Parent.Visible = documento.IsAvailableToUser(magicbox.DataSource.Contenuto.ChiaveEsterna);                
            }

    }

}
