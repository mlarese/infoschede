using System;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using NextFramework;
using NextFramework.NextWeb;

/// <summary>
/// La dynalay.
/// </summary>
public partial class _Default : NextFramework.NextWeb.NextPage {

    
    /// <summary>
    /// Inizializza la dynalay inserendo anche i layer associati.
    /// </summary>
    /// <param name="e">Argomento di default per questo metodo.</param>
    protected override void OnLoad(EventArgs e) {
        
        // richiama l'OnLoad della NextPage (non necessario se si usa Page_Load)
        base.OnLoad(e);

        // inizializzo i dati della pagina
        Initialize();

        // inizializzo il tag head
        SetHead();

        // inizializzo il tag body
        SetBody();

        // inizializzo il tag form inserendo i layers
        SetForm();
    }

}