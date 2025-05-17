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
using NextFramework.NextPassport;
using NextFramework.NextB2B;
using NextPdfTools;

public partial class Plugin_InviaEmail : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// utente autenticato
    /// </summary>
    protected NextMembershipRivenditore utente = new NextMembershipRivenditore();

    int schedaId = 0, utenteId = 0, adminId = 0, ddtId = 0;

    string pageUrl;

    DataTable dt = new DataTable();

    #region Settings...

    /// <summary>
    /// Nome del parametro del passport per il testo di default nell'email da inviare
    /// </summary>
    public string StnParametroNome;

    /// <summary>
    /// Codice Tipo di email da inviare (preventivo, consuntivo, ecc.)
    /// </summary>
    public string StnTipo;

    /// <summary>
    /// Nome del Tipo di email da inviare (preventivo, consuntivo, ecc.)
    /// </summary>
    public string StnTipoNome;

    /// <summary>
    /// Pagina utilizzata per la creazione del documento allegato in pdf.
    /// </summary>
    public int StnPaginaAllegatoId;

    /// <summary>
    /// Pagina inviata come email contenente l'allegato.
    /// </summary>
    public int StnPaginaEmailId;

    /// <summary>
    /// Se true visualizza anche i campi aggiuntivi (per gli invii di richiesta ritiro e lettera di vettura).
    /// </summary>
    public bool StnVisualizzaAltro;

    /// <summary>
    /// Se true invia in copia anche al cliente (oltre che al trasportatore).
    /// </summary>
    public bool StnInviaCopiaACliente;

    #endregion
    
    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        try { StnParametroNome = _settings["parametroNome"]; } catch { }

        try { StnTipo = _settings["tipo"]; } catch { }
        try { StnTipoNome = _settings["tipoNome"]; } catch { }

        try { StnPaginaAllegatoId = int.Parse(_settings["paginaAllegatoId"]); } catch { };

        try { StnPaginaEmailId = int.Parse(_settings["paginaEmailId"]); } catch { };

        try { StnVisualizzaAltro = bool.Parse(_settings["visualizzaAltro"]); } catch { StnVisualizzaAltro = false; };

        try { StnInviaCopiaACliente = bool.Parse(_settings["inviaCopiaACliente"]); }
        catch { StnInviaCopiaACliente = false; };

        #endregion

        base.OnLoad(e);

        if (StnInviaCopiaACliente)
            AvvisoInvioInCopia.Visible = true;

        // logout preventivo
        if (NextPage.User.Identity.IsAuthenticated)
            FormsAuthentication.SignOut();

        // se i parametri sono corretti imposta la pagina, altrimenti effettua logout e redirect alla home
        if (Request.QueryString["ID_ADMIN"] != null && Request.QueryString["ID_ADMIN"].ToString() != "" &&
            int.TryParse(Request.QueryString["ID_ADMIN"].ToString(), out adminId) && adminId > 0)
        {
            Titolo.InnerText = "Invia " + StnTipoNome;
            EmailAllegatoLabel.InnerText = "allegato da inviare:";
            EmailTestoLabel.InnerText = "testo da inviare:";
            Invia.Text = "Invia " + StnTipoNome;

            // scheda
            if (Request.QueryString["ID_SCHEDA"] != null && Request.QueryString["ID_SCHEDA"].ToString() != "" &&
                int.TryParse(Request.QueryString["ID_SCHEDA"].ToString(), out schedaId) && schedaId > 0)
            {
                dt = InfoschedeTools.GetSchedeDataTable(schedaId, 0, "", "", "");
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (int.TryParse(dt.Rows[0]["sc_cliente_id"].ToString(), out utenteId) && utenteId > 0)
                    {
                        if (utenteId > 0 && utente.SetPropertiesById(utenteId))
                        {
                            // autenticazione utente cliente necessaria per il pdf
                            FormsAuthentication.SetAuthCookie(utente.UserName, false);

                            // labels
                            NumeroSchedaLabel.InnerText = "scheda numero:";
                            DataRicevimentoLabel.InnerText = "data ricevimento:";
                            ClienteLabel.InnerText = "cliente:";
                            EmailClienteLabel.InnerText = "e-mail:";
                            
                            // dati scheda
                            NumeroSchedaValue.InnerText = dt.Rows[0]["sc_numero"].ToString();
                            DateTime dataRicevimento;
                            if (DateTime.TryParse(dt.Rows[0]["sc_data_ricevimento"].ToString(), out dataRicevimento))
                                DataRicevimentoValue.InnerText = dataRicevimento.ToString(NextDateTime.StringFormats.DateIta);
                            else
                                DataRicevimentoValue.InnerText = DateTime.Today.ToString(NextDateTime.StringFormats.DateIta);
                            ClienteValue.InnerText = utente.Contatto.GetName();
                            utente.Contatto.SetRecapiti();
                            EmailClienteValue.InnerText = utente.Contatto.Email;
                            
                            // anteprima allegato
                            pageUrl = NextPage.GetPageSiteUrlRedirect(StnPaginaAllegatoId, "SCHEDAID=" + schedaId +
                                                                                           "&CLIENTEID=" + utente.Id +
                                                                                           "&IDCNT=" + utente.Contatto.Id.ToString() +
                                                                                           "&KEY=" + utente.Contatto.CodiceInserimento);
                            EmailAllegato.HRef = pageUrl;
                            EmailAllegato.InnerText = "Visualizza " + StnTipoNome;

                            // testo email eventualmente da modificare
                            EmailTesto.InnerText = Parameters.GetParam<string>(StnParametroNome);
                        }
                        else
                        {
                            FormsAuthentication.SignOut();
                            Response.Redirect(NextPage.UrlHomePage);
                        }
                    }
                    else
                    {
                        FormsAuthentication.SignOut();
                        Response.Redirect(NextPage.UrlHomePage);
                    }
                }
                else
                {
                    FormsAuthentication.SignOut();
                    Response.Redirect(NextPage.UrlHomePage);
                }
            }

            // ddt
            else if (Request.QueryString["ID_DDT"] != null && Request.QueryString["ID_DDT"].ToString() != "" &&
                     int.TryParse(Request.QueryString["ID_DDT"].ToString(), out ddtId) && ddtId > 0)
            {
                dt = InfoschedeTools.GetDdtDataTable(ddtId, 0, 0, "", "");
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (int.TryParse(dt.Rows[0]["ddt_trasportatore_id"].ToString(), out utenteId) && utenteId > 0)
                    {
                        if (utenteId > 0 && utente.SetPropertiesById(utenteId))
                        {
                            // autenticazione utente cliente necessaria per il pdf
                            FormsAuthentication.SetAuthCookie(utente.UserName, false);

                            // labels
                            NumeroSchedaLabel.InnerText = "DDT numero:";
                            DataRicevimentoLabel.InnerText = "data richiesta:";
                            ClienteLabel.InnerText = "trasportatore:";
                            EmailClienteLabel.InnerText = "e-mail:";

                            // dati ddt
                            NumeroSchedaValue.InnerText = dt.Rows[0]["ddt_numero"].ToString();
                            DateTime dataRichiestaRitiro;
                            if (DateTime.TryParse(dt.Rows[0]["ddt_data"].ToString(), out dataRichiestaRitiro))
                                DataRicevimentoValue.InnerText = dataRichiestaRitiro.ToString(NextDateTime.StringFormats.DateIta);
                            else
                                DataRicevimentoValue.InnerText = DateTime.Today.ToString(NextDateTime.StringFormats.DateIta);
                            ClienteValue.InnerText = utente.Contatto.GetName();
                            utente.Contatto.SetRecapiti();
                            EmailClienteValue.InnerText = utente.Contatto.Email;

                            // anteprima allegato
                            pageUrl = NextPage.GetPageSiteUrlRedirect(StnPaginaAllegatoId, "DDTID=" + ddtId +
                                                                                           "&CLIENTEID=" + dt.Rows[0]["ddt_cliente_id"].ToString() +
                                                                                           "&IDCNT=" + dt.Rows[0]["IdElencoIndirizzi"].ToString() +
                                                                                           "&KEY=" + dt.Rows[0]["Codiceinserimento"].ToString());
                            EmailAllegato.HRef = pageUrl;
                            EmailAllegato.InnerText = "Visualizza " + StnTipoNome;

                            // testo email eventualmente da modificare
                            EmailTesto.InnerText = Parameters.GetParam<string>(StnParametroNome);
                        }
                        else
                        {
                            FormsAuthentication.SignOut();
                            Response.Redirect(NextPage.UrlHomePage);
                        }
                    }
                    else
                    {
                        FormsAuthentication.SignOut();
                        Response.Redirect(NextPage.UrlHomePage);
                    }
                }
                else
                {
                    FormsAuthentication.SignOut();
                    Response.Redirect(NextPage.UrlHomePage);
                }
            }
            else
            {
                FormsAuthentication.SignOut();
                Response.Redirect(NextPage.UrlHomePage);
            }
            NextControlsTools.SetCssClass(this.Layer, "inviaemail");
        }
        else
        {
            FormsAuthentication.SignOut();
            Response.Redirect(NextPage.UrlHomePage);
        }
    }
    
    protected void Invia_Click(object sender, EventArgs e)
    {
        // url
        string pdfPath;

        if (schedaId > 0)
            pdfPath = StnTipo + "_scheda_" + schedaId + "_" + DateTime.Now.Date.ToString(NextDateTime.StringFormats.DateTimeNoSeparatorISO) + ".pdf";
        else
            pdfPath = StnTipo + "_ddt_" + ddtId + "_" + DateTime.Now.Date.ToString(NextDateTime.StringFormats.DateTimeNoSeparatorISO) + ".pdf";
       
        string basePath = NextPage.PathUpload + "\\1\\pdf\\" + StnTipo + "\\" + DateTime.Now.ToString("yyyy-MM") + "\\",
              pageEmailUrl = NextPage.GetPageSiteUrl(StnPaginaEmailId, "SCHEDAID=" + schedaId + "&DDTID=" + ddtId +
                                                                        "&CLIENTEID=" + utenteId +
                                                                        "&IDCNT=" + utente.Contatto.Id.ToString() +
                                                                        "&KEY=" + utente.Contatto.CodiceInserimento +
                                                                        "&TESTO=" + EmailTesto.Value);
        // allegato
        string[] attachList = new string[1];
        attachList[0] = basePath + pdfPath;

        if (Session["inviato_" + StnTipo + "_" + (schedaId > 0 ? schedaId : ddtId)] == null)
            Session["inviato_" + StnTipo + "_" + (schedaId > 0 ? schedaId : ddtId)] = "";

        // pdf
        if (Session["inviato_" + StnTipo + "_" + (schedaId > 0 ? schedaId : ddtId)].ToString() == "ok" ||
            NextPdf.GetPdfFromPageUrl(pageUrl, basePath + pdfPath, true, true, 0, 25, 1, false, "", true))
        {
            try
            {
                string emailCcn = "";
                if (StnInviaCopiaACliente)
                {
                    if (Request.QueryString["ID_DDT"] != null && Request.QueryString["ID_DDT"].ToString() != "" &&
                        int.TryParse(Request.QueryString["ID_DDT"].ToString(), out ddtId) && ddtId > 0)
                    {
                        dt.Clear();
                        dt = InfoschedeTools.GetSchedeDataTable(0, 0, "", "", "", ddtId);
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            if (int.TryParse(dt.Rows[0]["sc_cliente_id"].ToString(), out utenteId) && utenteId > 0)
                            {
                                if (utenteId > 0)
                                {
                                    NextMembershipRivenditore utenteAgg = new NextMembershipRivenditore();
                                    utenteAgg.SetPropertiesById(utenteId);
                                    utenteAgg.Contatto.SetRecapiti();
                                    emailCcn = NextFramework.Messaggi.Messaggio.GetRecapito(utenteAgg.Contatto, NextFramework.Messaggi.Messaggio.TipoMessaggio.Email);
                                    if (string.IsNullOrEmpty(emailCcn))
                                        using (NextFramework.NextCom.BLLNumero bll = new NextFramework.NextCom.BLLNumero())
                                        {
                                            emailCcn = bll.GetNumero(utenteAgg.Id, NextFramework.Messaggi.Messaggio.ConvertFromTipoToTipoNumero(NextFramework.Messaggi.Messaggio.TipoMessaggio.Email), false, false);
                                        }
                                }
                            }
                        }
                    }
                }

                // email conferma
                NextPage.Alert.Email.SendPageFromAdminToContact(NextPage.Connection, "Invio " + StnTipoNome, "", pageEmailUrl,
                                                                adminId, utente.Contatto, null, emailCcn, true, true, attachList);
                Session["inviato_" + StnTipo + "_" + (schedaId > 0 ? schedaId : ddtId)] = "ok";

                // conferma invio
                DatiDiv.Visible = false;
                Titolo.InnerText = "Invio '" + StnTipoNome + "' correttamente riuscito";
                Invia.Visible = false;
                Chiudi.Visible = true;
            }
            catch
            {
                // problemi
                DatiDiv.Visible = false;
                Titolo.InnerText = "Invio '" + StnTipoNome + "' non riuscito; riprova.";
                Invia.Visible = false;
                Chiudi.Visible = true;
            }
            finally
            {
                FormsAuthentication.SignOut();
            }
        }
    }

    protected override void OnUnload(EventArgs e)
    {
        FormsAuthentication.SignOut();
    }
}
