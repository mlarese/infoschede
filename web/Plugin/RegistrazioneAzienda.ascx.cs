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
using System.Globalization;
using NextFramework;
using NextFramework.NextWeb;
using NextFramework.NextCom;
using NextFramework.NextControls;
using NextFramework.NextPassport;
using NextFramework.NextB2B;
using NextFramework.NextB2B.DSClientiPrezziTableAdapters;

public partial class Plugin_RegistrazioneAzienda : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente
    /// </summary>
    protected BLLUser utente;

    #region Settings...

    /// <summary>
    /// Imposta o meno l'autenticazione automatica dopo la registrazione.
    /// </summary>
    public bool StnAutoAutenticazione;

    /// <summary>
    /// Profilo da associare.
    /// </summary>
    public int StnProfiloId;

    /// <summary>
    /// Permesso da abilitare.
    /// </summary>
    public string StnPermesso;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        int.TryParse(_settings["profiloId"], out StnProfiloId);

        bool.TryParse(_settings["autoAutenticazione"], out StnAutoAutenticazione);

        StnPermesso = _settings["permesso"];

        #endregion

        base.OnLoad(e);
        ZOrder = 1;
        Titolo.Visible = !NextPage.IsEmail;
        Titolo.InnerText = "Inserisci i tuoi dati";
        NextControlsTools.SetCssClass(this.Layer, "registrazioneprivato");
    }

    protected void Contattaci_DataBound(object sender, EventArgs e)
    {
        int id_contatto = 0;
        // visualizzazione dati
        if (Request.QueryString["IDCNT"] != null)
        {
            // id del contatto inserito
            int.TryParse(Request.QueryString["IDCNT"], out id_contatto);
            if (id_contatto > 0)
            {
                // recupera login dell'utente appena inserito
                utente = new BLLUser();
                utente.GetUserByContattoId(id_contatto);
                string login = utente.GetUtenteLoginByContattoId(id_contatto);
                ((HtmlGenericControl)Contattaci.Form.FindControl("Login")).InnerText = utente.Login;
                ((HtmlGenericControl)Contattaci.Form.FindControl("Password")).InnerText = NextString.String(utente.Password.Length, '*');
            }
        }
    }

    protected void Contattaci_FormSave(object sender, EventArgs e)
    {
        string login, pwd;
        login = ((TextBox)Contattaci.Form.FindControl("Login")).Text;
        pwd = ((TextBox)Contattaci.Form.FindControl("Password")).Text;

        // aggiunge i dati dell'utente
        utente = new BLLUser();
        int id_ut = utente.AddUtenteByContattoId(Contattaci.Contatto.Id, login, pwd, true);
        // abilita utente
        string permesso = ProfilesB2B.Rivenditore;
        utente.AbilitaUtente(id_ut, Contattaci.Contatto.Id, permesso);
        utente.AbilitaUtente(id_ut, Contattaci.Contatto.Id, StnPermesso);

        // crea il rivenditore dal contatto, e lo aggiunge alla rubrica clienti
        NextMembershipRivenditore rivend = new NextMembershipRivenditore();
        //rivend.SetPropertiesByContattoId(Contattaci.Contatto.Id);
        //rivend.setCodiceRiv(id_ut);
        //rivend.Riv_profilo_id = StnProfiloId;
        //rivend.AddRivenditore();

        rivend.Riv_profilo_id = 0;
        rivend.Riv_porto_default_id = 0;
        rivend.Id = id_ut;
        rivend.setCodiceRiv(id_ut);
        rivend.Contatto.SetContatto(Contattaci.Contatto.Id);
        rivend.AddRivenditore();
        rivend.SetPropertiesByContattoId(Contattaci.Contatto.Id);

        // associa l'utente alla rubrica corrispondente al profilo
        string sql = "select pro_rubrica_id from gtb_profili where pro_id = " + StnProfiloId.ToString();
        int rubId = NextPage.Connection.ExecuteInt(sql);
        utente.RubricaAdd(Contattaci.Contatto.Id, rubId);

        // autenticazione
        if (!Page.User.Identity.IsAuthenticated && StnAutoAutenticazione)
            FormsAuthentication.SetAuthCookie(login, false);

        Contattaci.ConfermaParams = "PROBID=" + Request.QueryString["PROBID"] + "&ARTID=" + Request.QueryString["ARTID"];
    }

    protected void CheckUser(object sender, ObjectDataSourceMethodEventArgs e)
    {
        ((HtmlGenericControl)(Contattaci.Form.FindControl("Avviso"))).Visible = true;
        //recupera datasource
        BLLContatto c = (BLLContatto)e.InputParameters[0];
        BLLUser user = new BLLUser();
        string errors = user.CheckUserData(((TextBox)Contattaci.Form.FindControl("Login")).Text,
                                          ((TextBox)Contattaci.Form.FindControl("Password")).Text,
                                          ((TextBox)Contattaci.Form.FindControl("ConfermaPassword")).Text, "",
                                          Contattaci.Contatto.Id, 0);

        if (!string.IsNullOrEmpty(errors))
        {
            c.SaveErrors = errors;
        }
    }
}
