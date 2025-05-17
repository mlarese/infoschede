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
using NextFramework.NextWeb;
using NextFramework;
using NextFramework.NextControls;
using NextFramework.NextB2B;
using NextFramework.NextPassport;
using NextFramework.NextB2B.DSOrdiniTableAdapters;
using System.Security.Cryptography;
using System.IO;
using System.Text;

/// <summary>
/// Login dell'area riservata.
/// </summary>
public partial class Plugin_Login : NextFramework.NextControls.NextUserControl
{
    #region Settings...

    /// <summary>
    /// Pagina a cui viene reindirizzato l'utente; se 0 o NULL torna all'home page dell'area riservata.
    /// </summary>
    int StnPaginaRedirectId;

    /// <summary>
    /// Testo di descrizione dell'intestazione
    /// </summary>
    string StnIntestazione;

    /// <summary>
    /// Posizione del permesso all'interno del passport (valore salvato in rel_utenti_sito
    /// </summary>
    int StnPermessoAssistenza;

    #endregion

    /// <summary>
    /// Setta le variabili interne e la pagina di destinazione.
    /// </summary>
    /// <param name="e">Parametro di default per questo evento.</param>
    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);

        #region Settings init...

        try { StnPaginaRedirectId = int.Parse(Settings["paginaRedirectId"]); }
        catch {
            if (NextPage.PageSiteId == NextApplication.PageReservedLogin)
                StnPaginaRedirectId = NextApplication.PageReservedDefault;
            else
                StnPaginaRedirectId = NextPage.PageSiteId;
        }

        StnPermessoAssistenza = int.Parse(_settings["permessoAssistenza"]);

        StnIntestazione = NextLanguage.ChooseString(_settings, "Intestazione");

        #endregion
    
        if (NextPage.User.Identity.IsAuthenticated)
        {            
            if (StnPaginaRedirectId == NextApplication.PageReservedDefault)
                Response.Redirect(NextPage.GetPageSiteUrlRedirect(NextApplication.PageReservedDefault, ""));
            else
            {
                this.Visible = false;
                Layer.Visible = false; 
            }
        }
        else
        {
            //imposta etichette
            Login.FailureText = _nextLanguage.ChooseString("Errore nei dati immessi. Accesso non riuscito.",
                                                           "Input data error. Login failed");
            Login.LoginButtonText = _nextLanguage.ChooseString("entra", "enter");
            
            //imposta direttamente il focus sull'input di login.
            if (StnPaginaRedirectId == 0)
                Login.Focus();
        }

        if (!string.IsNullOrEmpty(StnIntestazione))
        {
            Intestazione.InnerText = StnIntestazione;
        }
        else
        {
            Intestazione.Visible = false;
        }
        // Imposta il bottone di default se l'utente batte invio al posto di cliccare
//        if(NextPage.PageSiteId == NextApplication.PageReservedLogin)
//            NextControlsTools.SetDefaultButton(this, (IButtonControl)Login.FindControl("LoginButton"));
    }

    /// <summary>
    /// Esegue il redirect eseguito correttamente.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void Login_LoggedIn(object sender, EventArgs e)
    {
        BLLUser utente = new BLLUser();
        Shoppingcart sc = new Shoppingcart();
        DSOrdini.ShoppingcartDataTable dt = new DSOrdini.ShoppingcartDataTable();

        int ut_id = 0;
        ut_id = utente.GetUser(((TextBox)Login.FindControl("UserName")).Text).UtenteId;

        string sql = "select * from rel_utenti_sito left join tb_utenti on ut_ID=rel_ut_id where rel_permesso=1 " +
                     "and ut_id=" + ut_id;

        DataTable datat = NextPage.Connection.GetDataTable(sql);

        if (datat.Rows.Count > 0)
        {
            string value = ((TextBox)Login.FindControl("UserName")).Text + ";" + ((TextBox)Login.FindControl("Password")).Text;
            string cripted = "";
            Random random = new Random();

            foreach (char c in value)
            {
                cripted += NextString.FixLenght(System.Convert.ToUInt32(c).ToString(), 3, '0');
                cripted += random.Next(0, 9).ToString();
            }
            FormsAuthentication.SignOut();
            Response.Redirect(NextString.QueryStringAdd(NextPage.UrlAdmin, "DATAFROMNET=" + cripted + "&EXECUTE=NETACCESS"));
        }
        
        // recupera la shopping cart corrente (se non è vuota) oppure l'ultima dell'utente loggato
        using (ShoppingcartTableAdapter ta = new ShoppingcartTableAdapter())
        {
            if (String.IsNullOrEmpty(Shoppingcart.Current.Sc_session_id))
            {
                dt = ta.GetShoppingcartByUtenteId(ut_id, false, false, 0);
                if (dt.Rows.Count > 0)
                {
                    sc.SetProperties(dt[0]);
                    sc.Update();
                    ta.Recupera(Session.SessionID, Request.UserHostAddress, sc.Sc_id);
                }
            }
            else
            {
                sc = Shoppingcart.Current;
                if (sc.TotaleDettagli > 0)
                    ta.UpdateProprietario(ut_id, ut_id, sc.Sc_id);
                else
                {
                    dt = ta.GetShoppingcartByUtenteId((int?)ut_id, false, false, 0);
                    if (dt.Rows.Count > 0)
                    {
                        sc.SetProperties(dt[0]);
                        sc.Update();
                        ta.Recupera(Session.SessionID, Request.UserHostAddress, sc.Sc_id);
                    }
                }
            }
        }

        InfoschedeTools.Bacheca_PubblicaMessaggioInSessione(InfoschedeTools.Bacheca_TipoMessaggio.Ok, "Benvenuto", 
                                                           utente.Nome + " " + utente.Cognome + "<br>" + "Utente di: " + utente.Organizzazione, this);

        // redirect
        if (StnPaginaRedirectId == NextApplication.PageReservedDefault)
            Response.Redirect(NextPage.GetPageSiteUrlRedirect(NextApplication.PageReservedDefault, ""));
        else if (StnPaginaRedirectId == NextPage.PageSiteId)
            Response.Redirect(NextPage.Url);
        else
            Response.Redirect(NextPage.GetPageSiteUrlRedirect(StnPaginaRedirectId, ""));
    }

    protected void Login_PreRender(object sender, EventArgs e)
    {
        if (Request.QueryString.GetValues("ReturnUrl") != null)
        {
            InfoschedeTools.Bacheca_PubblicaMessaggio(InfoschedeTools.Bacheca_TipoMessaggio.Errore, "Solo per utenti registrati!", "Per poter utilizzare la funzionalità effettua il login.");
        }
    }
}