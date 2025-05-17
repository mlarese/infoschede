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
using System.Data.SqlClient;
using System.Globalization;
using NextFramework;
using NextFramework.NextWeb;
using NextFramework.NextControls;
using NextFramework.NextPassport;
using NextFramework.NextB2B;
using NextFramework.NextB2B.DSArticoliTableAdapters;

public partial class Plugin_Scheda : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente autenticato
    /// </summary>
    protected NextMembershipRivenditore cliente = NextMembershipRivenditore.CurrentCliente;

    int artId = 0, probId = 0, schedaId = 0, clienteId = 0, centroAssId = 0;
    decimal costoPresa = 0, costoRiconsegna = 0, oreManodopera, costoManodopera = 0, costoRicambi = 0, costoTotale = 0;

    bool inGaranzia = false, reqGaranzia = false;

    string sql, queryString, jqueryScript, marchioInfo;

    #region Settings...

    /// <summary>
    /// Pagina conferma scheda inserita.
    /// </summary>
    public int StnPaginaSchedaId;

    /// <summary>
    /// Pagina elenco richieste.
    /// </summary>
    public int StnPaginaElencoId;

    /// <summary>
    /// Pagina email conferma.
    /// </summary>
    public int StnPaginaEmailId;

    /// <summary>
    /// Stato scheda.
    /// </summary>
    public int StnStatoSchedaId;

    /// <summary>
    /// Id admin.
    /// </summary>
    public int StnAdminId;

    /// <summary>
    /// Titolo.
    /// </summary>
    public string StnTitolo;

    /// <summary>
    /// Descrizione.
    /// </summary>
    public string StnDescrizione;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        try { StnPaginaSchedaId = int.Parse(_settings["paginaSchedaId"]); }
        catch { };

        try { StnPaginaElencoId = int.Parse(_settings["paginaElencoId"]); }
        catch { };

        try { StnPaginaEmailId = int.Parse(_settings["paginaEmailId"]); }
        catch { StnPaginaEmailId = NextPage.PageSiteId; };

        try { StnStatoSchedaId = int.Parse(_settings["StatoSchedaId"]); }
        catch { StnStatoSchedaId = 0; };

        try { StnAdminId = int.Parse(_settings["adminId"]); }
        catch { StnAdminId = 0; };

        try { StnTitolo = _settings["Titolo"]; }
        catch { };
        if (String.IsNullOrEmpty(StnTitolo))
            StnTitolo = "Richiesta di assistenza";

        try { StnDescrizione = _settings["Descrizione"]; }
        catch { };
        if (String.IsNullOrEmpty(StnDescrizione))
            StnDescrizione = "Compilando il form sottostante ci farete pervenire la vostra richiesta di assistenza a cui risponderemo al piu' presto.";

        #endregion

        base.OnLoad(e);

        // gestione cliente pubblico per generazione pdf
        if ((cliente == null || cliente.IsPublic) && Request.QueryString["CLIENTEID"] != null &&
            int.TryParse(Request.QueryString["CLIENTEID"].ToString(), out clienteId) && clienteId > 0 &&
            Request.QueryString["KEY"] != null && Request.QueryString["IDCNT"] != null)
            if (!cliente.SetPropertiesById(clienteId) ||
                Request.QueryString["IDCNT"].ToString() != cliente.Contatto.Id.ToString() &&
                Request.QueryString["KEY"].ToString() != cliente.Contatto.CodiceInserimento)
                Response.Redirect(NextPage.UrlHomePage);
        
        // verifica cliente
        if ((cliente != null && !cliente.IsPublic))
        {
            #region labels

            Titolo.InnerText = StnTitolo;
            Descrizione.InnerText = StnDescrizione;
            TitoloPrincipali.InnerText = "Dati principali";
            TitoloModello.InnerText = "Dati del modello";
            TitoloAcquisto.InnerText = "Dati dell'acquisto";
            TitoloRiparazione.InnerText = "Dati della riparazione";
            TitoloControlli.InnerText = "Controlli effettuati durante e dopo la riparazione";
            TitoloTrasporto.InnerText = "Dati del trasporto";
            TitoloRiepilogo.InnerText = "Riepilogo totale prezzi (iva esclusa)";
            StatoSchedaLabel.InnerText = "Stato scheda:";
            NumeroSchedaLabel.InnerText = "Numero:";
            DataRicevimentoLabel.InnerText = "Data ricevimento:";
            CentroAssistenzaLabel.InnerText = "Centro assistenza:";
            ClienteLabel.InnerText = "Cliente:";
            NoteClienteLabel.InnerText = "Note del cliente:";
            RiferimentoClienteLabel.InnerText = "Riferimento cliente:";
            ModelloLabel.InnerText = "Modello:";
            ModelloVariantiLabel.InnerText = "varianti:";
            MatricolaLabel.InnerText = "Matricola:";
            DataAcquistoLabel.InnerText = "Data di acquisto:";
            NegozioAcquistoLabel.InnerText = "Negozio di acquisto:";
            NumeroScontrinoLabel.InnerText = "Numero scontrino:";
            GaranziaLabel.InnerText = (Request.QueryString["SCHEDAID"] != null ? "In garanzia:" : "Richiedi garanzia:");
            AccessoriListaLabel.InnerText = "Accessori presenti:";
            AccessoriAltroLabel.InnerText = (Request.QueryString["SCHEDAID"] != null ? "Accessori presenti:" : "altro:");
            GuastoSegnalatoLabel.InnerText = "Guasto segnalato:";
            GuastoSegnalatoAltroLabel.InnerText = GuastoSegnalatoLabel.InnerText;
            GuastoRiscontratoLabel.InnerText = "Guasto riscontrato:";
            EsitoInterventoLabel.InnerText = "Esito dell'intervento:";
            DataFineLavoroLabel.InnerText = "Data fine lavoro:";
            OreManodoperaLabel.InnerText = "Ore manodopera intervento:";
            PrezzoManodoperaLabel.InnerText = "Prezzo totale manodopera:";
            RicambiUtilizzatiLabel.InnerText = "Ricambi utilizzati:";
            NoteChiusuraLabel.InnerText = "Note di chiusura:";
            CostoPresaLabel.InnerText = "Costo presa:";
            NumeroDdtCaricoLabel.InnerText = "Numero DDT di carico:";
            DataDdtCaricoLabel.InnerText = "Data DDT di carico:";
            CostoRiconsegnaLabel.InnerText = "Costo riconsegna:";
            NumeroDdtRiconsegnaLabel.InnerText = "Numero DDT di riconsegna:";
            DataDdtRiconsegnaLabel.InnerText = "Data DDT di riconsegna:";
            TrasportatoreLabel.InnerText = "Trasportatore:";
            CostoPresaRiconsegnaLabel.InnerText = "Costi di presa / riconsegna:";
            CostoManodoperaLabel.InnerText = "Costo di manodopera:";
            CostoRicambiLabel.InnerText = "Costo totale ricambi:";
            CostoTotaleLabel.InnerText = "Costo totale scheda:";
            Invia.Text = "Invia richiesta";
            Indietro.Text = "Torna all'elenco";

            #endregion

            // verifica modalità
            if (Request.QueryString["ARTID"] != null && Request.QueryString["PROBID"] != null)
            {
                if (int.TryParse(Request.QueryString["ARTID"].ToString(), out artId) && artId > 0 &&
                int.TryParse(Request.QueryString["PROBID"].ToString(), out probId) && probId >= 0)
                {
                    #region visibilità
                    StatoSchedaDiv.Visible = false;
                    NumeroSchedaDiv.Visible = false;
                    CentroAssistenzaDiv.Visible = false;
                    RiferimentoClienteDiv.Visible = cliente.Contatto.IsSocieta;
                    GuastoSegnalatoDiv.Visible = (probId > 0);
                    GuastoSegnalatoAltroDiv.Visible = (probId == 0);
                    GuastoRiscontratoDiv.Visible = false;
                    EsitoInterventoDiv.Visible = false;
                    DataFineLavoroDiv.Visible = false;
                    OreManodoperaDiv.Visible = false;
                    PrezzoManodoperaDiv.Visible = false;
                    RicambiUtilizzatiDiv.Visible = false;
                    NoteChiusuraDiv.Visible = false;
                    TitoloTrasportoDiv.Visible = false;
                    CostoPresaDiv.Visible = false;
                    NumeroDdtCaricoDiv.Visible = false;
                    DataDdtCaricoDiv.Visible = false;
                    CostoRiconsegnaDiv.Visible = false;
                    NumeroDdtRiconsegnaDiv.Visible = false;
                    DataDdtRiconsegnaDiv.Visible = false;
                    TrasportatoreDiv.Visible = false;
                    TitoloControlliDiv.Visible = false;
                    DescrittoriDiv.Visible = false;
                    TitoloRiepilogoDiv.Visible = false;
                    CostoPresaRiconsegnaDiv.Visible = false;
                    CostoManodoperaDiv.Visible = false;
                    CostoRicambiDiv.Visible = false;
                    CostoTotaleDiv.Visible = false;
                    RiferimentoClienteValue.Visible = false;
                    MatricolaValue.Visible = false;
                    DataAcquistoValue.Visible = false;
                    NegozioAcquistoValue.Visible = false;
                    NumeroScontrinoValue.Visible = false;
                    AccessoriListaValue.Visible = false;
                    AccessoriAltroValue.Visible = false;
                    GuastoSegnalatoAltroValue.Visible = false;
                    NoteClienteValue.Visible = false;
                    IndietroDiv.Visible = false;
                    #endregion

                    // data ricevimento
                    DataRicevimentoValue.InnerText = DateTime.Today.ToString(NextDateTime.StringFormats.DateIta);

                    // cliente
                    ClienteValue.InnerText = cliente.Contatto.GetName();

                    // modello
                    sql = "SELECT art_nome_it" +
                           " FROM gtb_articoli" +
                     " INNER JOIN grel_art_valori ON rel_art_id = art_id" +
                          " WHERE rel_art_id = " + artId;
                    ModelloValue.InnerText = (string)NextPage.Connection.ExecuteScalar(sql);
                    Articolo modello = new Articolo();
                    if (modello.SetArticolo(artId))
                    {
                        // articolo con più varianti
                        if (modello.Art_varianti)
                        {
                            string listaVarianti = modello.Variante.GetListaValoriText(true);
                            using (ArticoliVariantiTableAdapter ta = new ArticoliVariantiTableAdapter())
                            {
                                DSArticoli.ArticoliVariantiDataTable dtVar =
                                    ta.GetArticoliVariantiNonDisabilitatiByArticoloIdValoreId(artId, null);
                                if (dtVar.Count > 0)
                                {

                                }
                            }
                        }
                        else
                            ModelloVariantiDiv.Visible = false;

                        // modello generico di default
                        if (modello.Art_cod_int == InfoschedeTools.CodModelloDefault)
                        {
                            ModelloValue.Visible = false;
                            ModelloInput.Visible = true;

                            // logo ed immagine aggiuntiva costruttore
                            int marcaId = 0;
                            if (int.TryParse(Request.QueryString["MARID"], out marcaId) && marcaId > 0)
                            {
                                sql = "SELECT *" +
                                       " FROM gtb_marche" +
                                      " WHERE mar_id = " + marcaId;
                                using (DataTable dtMarca = NextPage.Connection.GetDataTable(sql))
                                {
                                    if (dtMarca.Rows.Count > 0)
                                    {
                                        if (!String.IsNullOrEmpty(dtMarca.Rows[0]["mar_logo"].ToString()))
                                            LogoMarcaImg.Src = NextPage.UrlImages + dtMarca.Rows[0]["mar_logo"].ToString().TrimStart('/');
                                        else
                                            LogoMarcaDiv.Visible = false;
                                        if (!String.IsNullOrEmpty(dtMarca.Rows[0]["mar_img"].ToString()))
                                            marchioInfo = dtMarca.Rows[0]["mar_img"].ToString();
                                    }
                                }
                            }
                        }

                        else
                        {
                            ModelloValue.Visible = true;
                            ModelloInput.Visible = false;

                            // logo ed immagine aggiuntiva costruttore
                            if (!String.IsNullOrEmpty(modello.Mar_logo))
                                LogoMarcaImg.Src = NextPage.UrlImages + modello.Mar_logo.TrimStart('/');
                            else
                                LogoMarcaDiv.Visible = false;
                            if (!String.IsNullOrEmpty(modello.Mar_img))
                                marchioInfo = modello.Mar_img;
                        }
                    }
                    else
                        ModelloVariantiDiv.Visible = false;

                    NumeroScontrinoInfoImg.Src = NextPage.UrlImages + "interfaccia/icona_informazioni_mini.jpg";
                    NumeroScontrinoEsempioImg.Src = NextPage.UrlImages + "interfaccia/scontrino_fiscale_mini.jpg";

                    if (!String.IsNullOrEmpty(marchioInfo))
                    {
                        MatricolaEsempioImg.Src = NextPage.UrlImages + marchioInfo.TrimStart('/');
                        ModelloInfoImg.Src = NextPage.UrlImages + "interfaccia/icona_informazioni_mini.jpg";
                        MatricolaInfoImg.Src = NumeroScontrinoInfoImg.Src;
                    }
                    else
                    {
                        ModelloInfoImg.Visible = false;
                        MatricolaInfoImg.Visible = false;
                    }
                    
                    #region script visualizza info modello e numero scontrino
                    jqueryScript = "\n" +
                                   "function visualizzaInfoModello(s) {\n" +
                                   "    $(\"img.nmatricola\").toggleClass(\"expanded\"); \n }\n" +
                                   "function visualizzaInfoScontrino(s) {\n" +
                                   "    $(\"img.nscontrino\").toggleClass(\"expanded\"); \n }\n" +
                                   "$(function() { \n" +
                                   "    $('#" + ModelloInfoImg.ClientID + "').mouseover(visualizzaInfoModello); \n }); \n" +
                                   "$(function() { \n" +
                                   "    $('#" + ModelloInfoImg.ClientID + "').mouseout(visualizzaInfoModello); \n }); \n" +
                                   "$(function() { \n" +
                                   "    $('#" + MatricolaInfoImg.ClientID + "').mouseover(visualizzaInfoModello); \n }); \n" +
                                   "$(function() { \n" +
                                   "    $('#" + MatricolaInfoImg.ClientID + "').mouseout(visualizzaInfoModello); \n }); \n" +
                                   "$(function() { \n" +
                                   "    $('#" + NumeroScontrinoInfoImg.ClientID + "').mouseover(visualizzaInfoScontrino); \n }); \n" +
                                   "$(function() { \n" +
                                   "    $('#" + NumeroScontrinoInfoImg.ClientID + "').mouseout(visualizzaInfoScontrino); \n }); \n";
                    NextPage.JQueryReadyManager.AddReadyScript(jqueryScript);
                    #endregion

                    // accessori
                    NextDropDownList.SetDropDownList(AccessoriListaDdl, "Scegli", "Nessun accessorio", false);
                    sql = "SELECT acc_id AS ID" +
                               ", acc_nome_it AS NOME" +
                           " FROM sgtb_accessori";
                    DataTable dt = NextPage.Connection.GetDataTable(sql);
                    if (dt.Rows.Count > 0)
                    {
                        AccessoriListaDdl.DataSource = dt;
                        AccessoriListaDdl.DataValueField = "ID";
                        AccessoriListaDdl.DataTextField = "NOME";
                        AccessoriListaDdl.HtmlEncode = false;
                        AccessoriListaDdl.DataBind();
                        ListItem altro = new ListItem("altro...", "-1");
                        AccessoriListaDdl.Items.Add(altro);
                    }
                    else
                    {
                        AccessoriListaDdl.Visible = false;
                        AccessoriAltroLabel.Visible = false;
                    }

                    #region script abilita input
                    jqueryScript = "\n" +
                                   "$(function() { \n" +
                                   "    $('select#" + AccessoriListaDdl.ClientID + "').change(altriAccessori); \n" +
                                   "}); \n" +
                                   "function altriAccessori() {\n" +
                                   "    if ($('select#" + AccessoriListaDdl.ClientID + "').val() < 0) {\n" +
                                   "        $(\"div.altri_accessori\").addClass(\"expanded\",true); \n" +
                                   "    }\n" +
                                   "    else {\n" +
                                   "        $(\"div.altri_accessori\").removeClass(\"expanded\",true); \n" +
                                   "    }\n" +
                                   "}\n";
                    NextPage.JQueryReadyManager.AddReadyScript(jqueryScript);
                    #endregion

                    #region script datepicker
                    string daysOfWeek = _nextLanguage.ChooseString("'do','lu','ma','me','gi','ve','sa'", "'su','mo','tu','we','th','fr','sa'"),
                           datepicker_options = " minDate: '-360', \n" +
                                                " maxDate: '0', \n" +
                                                " showOn: 'both', \n" +
                                                " dateFormat: 'dd/mm/yy', \n" +
                                                " prevText: '&#x3c;', \n" +
                                                " nextText: '&#x3e;', \n" +
                                                " dayNamesShort: [" + daysOfWeek + "], \n" +
                                                " dayNamesMin: [" + daysOfWeek + "], \n" +
                                                " buttonText: '', \n" +
                                                " buttonImage: '" + NextPage.UrlImages + "/interfaccia/calendar.gif', \n" +
                                                " selectOtherMonths: true, \n" +
                                                " showOtherMonths: true, \n" +
                                                " selectDefaultDate: true, \n" +
                                                " hideIfNoPrevNext: true, \n" +
                                                " defaultDate: '0' \n",
                           datepicker_prefix = " var dates = ";
                    NextPage.JQueryReadyManager.ApplyPluginDatePicker("#" + DataAcquistoInput.ClientID,
                                                                      datepicker_options, datepicker_prefix);
                    #endregion

                    // guasto segnalato
                    if (probId > 0)
                    {
                        sql = "SELECT prb_nome_it" +
                               " FROM sgtb_problemi" +
                              " WHERE prb_id = " + probId;
                        GuastoSegnalatoValue.InnerText = (string)NextPage.Connection.ExecuteScalar(sql);
                    }
                }
                else
                    Response.Redirect(NextPage.UrlHomePage);
            }

            else if (Request.QueryString["SCHEDAID"] != null)
            {
                if (int.TryParse(Request.QueryString["SCHEDAID"].ToString(), out schedaId) && schedaId > 0)
                {
                    #region visibilità
                    StatoSchedaDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    CentroAssistenzaDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    RiferimentoClienteDiv.Visible = (Request.QueryString["CONFERMA"] == null) || cliente.Contatto.IsSocieta;
                    ModelloVariantiDiv.Visible = false;
                    TitoloControlliDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    DescrittoriDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    TitoloRiepilogoDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    CostoPresaRiconsegnaDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    CostoManodoperaDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    CostoRicambiDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    CostoTotaleDiv.Visible = (Request.QueryString["CONFERMA"] == null);
                    RiferimentoClienteInput.Visible = false;
                    ModelloInput.Visible = false;
                    MatricolaInput.Visible = false;
                    DataAcquistoInput.Visible = false;
                    NegozioAcquistoInput.Visible = false;
                    NumeroScontrinoInput.Visible = false;
                    GaranziaCb.Visible = false;
                    AccessoriListaDdl.Visible = false;
                    AccessoriAltroInput.Visible = false;
                    GuastoSegnalatoAltroInput.Visible = false;
                    NoteClienteTxtarea.Visible = false;
                    ModelloInfoImg.Visible = false;
                    MatricolaInfoImg.Visible = false;
                    MatricolaEsempioImg.Visible = false;
                    NumeroScontrinoInfoImg.Visible = false;
                    NumeroScontrinoEsempioImg.Visible = false;
                    InviaDiv.Visible = false;
                    IndietroDiv.Visible = !NextPage.IsEmail && Request.QueryString["CLIENTEID"] == null;
                    #endregion
                    
                    // recupera dati
                    DataTable dt = InfoschedeTools.GetSchedeDataTable(schedaId, 0, "", "", "");
                    if (dt.Rows.Count > 0)
                    {
                        // data ricevimento
                        DateTime dataRicevimento;
                        if (DateTime.TryParse(dt.Rows[0]["sc_data_ricevimento"].ToString(), out dataRicevimento))
                            DataRicevimentoValue.InnerText = dataRicevimento.ToString(NextDateTime.StringFormats.DateIta);
                        else
                            DataRicevimentoValue.InnerText = DateTime.Today.ToString(NextDateTime.StringFormats.DateIta);

                        // numero scheda
                        NumeroSchedaValue.InnerText = dt.Rows[0]["sc_numero"].ToString();

                        // stato scheda
                        StatoSchedaValue.InnerText = dt.Rows[0]["stato"].ToString();

                        // cliente
                        if (clienteId > 0 || int.TryParse(dt.Rows[0]["sc_cliente_id"].ToString(), out clienteId) &&
                            (clienteId == cliente.Id) || cliente.Riv_profilo_id == 1 || cliente.Riv_profilo_id == 5)
                            ClienteValue.InnerText = dt.Rows[0]["nome_rivenditore"].ToString();//cliente.Contatto.GetName();
                        else
                            Response.Redirect(NextPage.UrlHomePage);

                        // note cliente
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_note_cliente"].ToString()))
                            NoteClienteValue.InnerText = dt.Rows[0]["sc_note_cliente"].ToString();
                        else
                            NoteClienteDiv.Visible = false;

                        // riferimento cliente
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_rif_cliente"].ToString()))
                            RiferimentoClienteValue.InnerText = dt.Rows[0]["sc_rif_cliente"].ToString();
                        else
                            RiferimentoClienteDiv.Visible = false;
                        
                        // modello
                        ModelloValue.InnerText = dt.Rows[0]["modello"].ToString();

                        // logo costruttore
                        string imageSrc = "";
                        if (Request.QueryString["MARID"] != null)
                        {
                            int marcaId = 0;
                            if (int.TryParse(Request.QueryString["MARID"].ToString(), out marcaId) && marcaId > 0)
                            {
                                sql = "SELECT *" +
                                       " FROM gtb_marche" +
                                      " WHERE mar_id = " + marcaId;
                                using (DataTable dtMarca = NextPage.Connection.GetDataTable(sql))
                                {
                                    if (dtMarca.Rows.Count > 0)
                                        if (!String.IsNullOrEmpty(dtMarca.Rows[0]["mar_logo"].ToString()))
                                            imageSrc = dtMarca.Rows[0]["mar_logo"].ToString().TrimStart('/');
                                }
                            }
                        }
                        else
                            imageSrc = dt.Rows[0]["mar_logo"].ToString().TrimStart('/');
                        if (!String.IsNullOrEmpty(imageSrc))
                            LogoMarcaImg.Src = NextPage.UrlImages + imageSrc.TrimStart('/');
                        else
                            LogoMarcaDiv.Visible = false;
                        
                        // matricola
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_matricola"].ToString()))
                            MatricolaValue.InnerText = dt.Rows[0]["sc_matricola"].ToString();
                        else
                            MatricolaDiv.Visible = false;

                        // data acquisto
                        DateTime dataAcquisto;
                        if (DateTime.TryParse(dt.Rows[0]["sc_data_acquisto"].ToString(), out dataAcquisto))
                            DataAcquistoValue.InnerText = dataAcquisto.ToString(NextDateTime.StringFormats.DateIta);
                        else
                            DataAcquistoDiv.Visible = false;

                        // negozio acquisto
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_negozio_acquisto"].ToString()))
                            NegozioAcquistoValue.InnerText = dt.Rows[0]["sc_negozio_acquisto"].ToString();
                        else
                            NegozioAcquistoDiv.Visible = false;

                        // numero scontrino
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_numero_scontrino"].ToString()))
                            NumeroScontrinoValue.InnerText = dt.Rows[0]["sc_numero_scontrino"].ToString();
                        else
                            NumeroScontrinoDiv.Visible = false;

                        // garanzia
                        if (bool.TryParse(dt.Rows[0]["sc_in_garanzia"].ToString(), out inGaranzia) &&
                            (bool.TryParse(dt.Rows[0]["sc_richiesta_garanzia"].ToString(), out reqGaranzia)) ||
                             String.IsNullOrEmpty(dt.Rows[0]["sc_richiesta_garanzia"].ToString()))
                            GaranziaValue.InnerText = (inGaranzia ? "Sì" : (reqGaranzia ? "Richiesta in attesa di conferma" : "No"));
                        else
                            GaranziaValue.InnerText = inGaranzia + "No";

                        // accessorio
                        if (!String.IsNullOrEmpty(dt.Rows[0]["accessorio"].ToString()))
                            AccessoriListaValue.InnerText = dt.Rows[0]["accessorio"].ToString();
                        else
                            AccessoriListaDiv.Visible = false;

                        // altro accessorio
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_accessori_presenti_altro"].ToString()))
                        {
                            AccessoriAltroValue.InnerText = dt.Rows[0]["sc_accessori_presenti_altro"].ToString();
                            AccessoriAltroLabel.Style.Clear();
                            NextControlsTools.SetCssClass(AccessoriAltroDiv, "expanded");
                        }
                        else
                            AccessoriAltroDiv.Visible = false;

                        // guasto segnalato
                        if (!String.IsNullOrEmpty(dt.Rows[0]["guasto_segnalato"].ToString()))
                            GuastoSegnalatoValue.InnerText = dt.Rows[0]["guasto_segnalato"].ToString();
                        else
                            GuastoSegnalatoDiv.Visible = false;

                        // altro guasto segnalato
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_guasto_segnalato_altro"].ToString()))
                            GuastoSegnalatoAltroValue.InnerText = dt.Rows[0]["sc_guasto_segnalato_altro"].ToString();
                        else
                            GuastoSegnalatoAltroDiv.Visible = false;

                        // guasto riscontrato
                        if (!String.IsNullOrEmpty(dt.Rows[0]["guasto_riscontrato"].ToString()))
                            GuastoRiscontratoValue.InnerText = dt.Rows[0]["guasto_riscontrato"].ToString();
                        else
                            GuastoRiscontratoDiv.Visible = false;

                        // altro guasto riscontrato
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_guasto_riscontrato_altro"].ToString()))
                            GuastoRiscontratoAltroValue.InnerText = dt.Rows[0]["sc_guasto_riscontrato_altro"].ToString();
                        else
                            GuastoRiscontratoAltroDiv.Visible = false;

                        // esito intervento
                        if (!String.IsNullOrEmpty(dt.Rows[0]["esito_intervento"].ToString()))
                            EsitoInterventoValue.InnerText = dt.Rows[0]["esito_intervento"].ToString();
                        else
                            EsitoInterventoDiv.Visible = false;

                        // data fine lavoro
                        DateTime dataFineLavoro;
                        if (DateTime.TryParse(dt.Rows[0]["sc_data_fine_lavoro"].ToString(), out dataFineLavoro))
                            DataFineLavoroValue.InnerText = dataFineLavoro.ToString(NextDateTime.StringFormats.DateIta);
                        else
                            DataFineLavoroDiv.Visible = false;

                        // ore manodopera
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_ora_manodopera_intervento"].ToString()) &&
                            dt.Rows[0]["sc_ora_manodopera_intervento"].ToString() != "0" && 
                            decimal.TryParse(dt.Rows[0]["sc_ora_manodopera_intervento"].ToString(), out oreManodopera))
                            OreManodoperaValue.InnerText = oreManodopera.ToString();
                        else
                            OreManodoperaDiv.Visible = false;

                        // prezzo manodopera
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_prezzo_manodopera"].ToString()) &&
                            dt.Rows[0]["sc_prezzo_manodopera"].ToString() != "0,0000" && !inGaranzia)
                        {
                            if (decimal.TryParse(dt.Rows[0]["sc_prezzo_manodopera"].ToString(), out costoManodopera))
                            {
                                costoManodopera = costoManodopera * oreManodopera;
                                PrezzoManodoperaValue.InnerText = NextString.FormatEuro(costoManodopera.ToString());
                            }
                            else
                                PrezzoManodoperaDiv.Visible = false;
                        }
                        else
                            PrezzoManodoperaDiv.Visible = false;

                        // ricambi utilizzati
                        sql = "SELECT *" +
                               " FROM sgtb_dettagli_schede" +
                              " WHERE dts_scheda_id = " + schedaId;
                        DataTable dtRic = NextPage.Connection.GetDataTable(sql);
                        if (dtRic.Rows.Count > 0)
                        {
                            RicambiUtilizzatiLista.DataSource = dtRic;
                            RicambiUtilizzatiLista.DataBind();                            
                        }
                        else
                            RicambiUtilizzatiDiv.Visible = false;

                        // note di chiusura
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_note_chiusura"].ToString()))
                            NoteChiusuraValue.InnerText = dt.Rows[0]["sc_note_chiusura"].ToString();
                        else
                            NoteChiusuraDiv.Visible = false;

                        // descrittori
                        sql = "SELECT *" +
                               " FROM sgtb_descrittori d" +
						  " LEFT JOIN srel_descrittori_schede r ON d.des_id = r.rds_descrittore_id AND r.rds_scheda_id = " + schedaId +
						  " LEFT JOIN sgtb_descrittori_raggruppamenti g ON d.des_raggruppamento_id = g.rag_id";
                        DataTable dtDesc = NextPage.Connection.GetDataTable(sql);
                        if (dtDesc.Rows.Count > 0)
                        {
                            Descrittori.DataSource = dtDesc;
                            Descrittori.DataBind();
                        }
                        else
                        {
                            TitoloControlliDiv.Visible = false;
                            DescrittoriDiv.Visible = false;
                        }

                        // costo presa in carico
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_costo_presa"].ToString()) &&
                            dt.Rows[0]["sc_costo_presa"].ToString() != "0,0000" && !inGaranzia)
                        {
                            if (decimal.TryParse(dt.Rows[0]["sc_costo_presa"].ToString(), out costoPresa))
                                CostoPresaValue.InnerText = NextString.FormatEuro(dt.Rows[0]["sc_costo_presa"].ToString());
                            else
                                CostoPresaDiv.Visible = false;
                        }
                        else
                            CostoPresaDiv.Visible = false;

                        // numero ddt di carico
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_numero_DDT_di_carico"].ToString()) &&
                            dt.Rows[0]["sc_numero_DDT_di_carico"].ToString() != "0")
                            NumeroDdtCaricoValue.InnerText = dt.Rows[0]["sc_numero_DDT_di_carico"].ToString();
                        else
                            NumeroDdtCaricoDiv.Visible = false;

                        // data ddt di carico
                        DateTime dataDdtCarico;
                        if (DateTime.TryParse(dt.Rows[0]["sc_data_DDT_di_carico"].ToString(), out dataDdtCarico))
                            DataDdtCaricoValue.InnerText = dataDdtCarico.ToString(NextDateTime.StringFormats.DateIta);
                        else
                            DataDdtCaricoDiv.Visible = false;

                        // costo riconsegna
                        if (!String.IsNullOrEmpty(dt.Rows[0]["sc_costo_riconsegna"].ToString()) &&
                            dt.Rows[0]["sc_costo_riconsegna"].ToString() != "0,0000" && !inGaranzia)
                        {
                            if (decimal.TryParse(dt.Rows[0]["sc_costo_riconsegna"].ToString(), out costoRiconsegna))
                                CostoRiconsegnaValue.InnerText = NextString.FormatEuro(dt.Rows[0]["sc_costo_riconsegna"].ToString());
                            else
                                CostoRiconsegnaDiv.Visible = false;
                        }
                        else
                            CostoRiconsegnaDiv.Visible = false;
                        
                        // numero ddt di riconsegna
                        if (!String.IsNullOrEmpty(dt.Rows[0]["numero_ddt"].ToString()) &&
                            dt.Rows[0]["numero_ddt"].ToString() != "0")
                            NumeroDdtRiconsegnaValue.InnerText = dt.Rows[0]["numero_ddt"].ToString();
                        else
                            NumeroDdtRiconsegnaDiv.Visible = false;

                        // data ddt di riconsegna
                        DateTime dataDdtRiconsegna;
                        if (DateTime.TryParse(dt.Rows[0]["data_ddt"].ToString(), out dataDdtRiconsegna))
                            DataDdtRiconsegnaValue.InnerText = dataDdtRiconsegna.ToString(NextDateTime.StringFormats.DateIta);
                        else
                            DataDdtRiconsegnaDiv.Visible = false;

                        // trasportatore
                        if (!String.IsNullOrEmpty(dt.Rows[0]["trasportatore_ddt"].ToString()))
                            TrasportatoreValue.InnerText = dt.Rows[0]["trasportatore_ddt"].ToString();
                        else
                            TrasportatoreDiv.Visible = false;

                        // riepilogo costi presa e riconsegna
                        if ((costoPresa + costoRiconsegna) > 0 && !inGaranzia)
                            CostoPresaRiconsegnaValue.InnerText = NextString.FormatEuro((costoPresa + costoRiconsegna).ToString());
                        else
                            CostoPresaRiconsegnaDiv.Visible = false;

                        // riepilogo costi manodopera
                        if (costoManodopera > 0 && !inGaranzia)
                            CostoManodoperaValue.InnerText = NextString.FormatEuro(costoManodopera.ToString());
                        else
                            CostoManodoperaDiv.Visible = false;

                        // riepilogo costi ricambi
                        if (costoRicambi > 0 && !inGaranzia)
                            CostoRicambiValue.InnerText = NextString.FormatEuro(costoRicambi.ToString());
                        else
                            CostoRicambiDiv.Visible = false;

                        // riepilogo costi totali
                        costoTotale = costoPresa + costoRiconsegna + costoManodopera + costoRicambi;
                        if (costoTotale > 0 && !inGaranzia)
                            CostoTotaleValue.InnerText = NextString.FormatEuro(costoTotale.ToString());
                        else
                            CostoTotaleDiv.Visible = false;

                        // se non ci sono elementi visibili nella sezione trasporti nasconde il titolo
                        if (!CostoPresaDiv.Visible && !NumeroDdtCaricoDiv.Visible && !DataDdtCaricoDiv.Visible &&
                            !CostoRiconsegnaDiv.Visible && !NumeroDdtRiconsegnaDiv.Visible &&
                            !DataDdtRiconsegnaDiv.Visible && !TrasportatoreDiv.Visible)
                            TitoloTrasportoDiv.Visible = false;

                        // se non ci sono elementi visibili nella sezione riepilogo nasconde il titolo
                        if (!CostoPresaRiconsegnaDiv.Visible && !CostoManodoperaDiv.Visible &&
                            !CostoRicambiDiv.Visible && !CostoTotaleDiv.Visible || inGaranzia)
                            TitoloRiepilogoDiv.Visible = false;

                        // centro assistenza
                        if (int.TryParse(dt.Rows[0]["sc_centro_assistenza_id"].ToString(), out centroAssId) && centroAssId > 0)
                            CentroAssistenzaValue.InnerText = dt.Rows[0]["centro_assistenza"].ToString();
                        else
                            CentroAssistenzaDiv.Visible = false;
                    }
                    else
                        Response.Redirect(NextPage.UrlHomePage);
                }
                else
                    Response.Redirect(NextPage.UrlHomePage);
            }
            else
                Response.Redirect(NextPage.UrlHomePage);
        }
        else
            Response.Redirect(NextPage.UrlHomePage);

        NextControlsTools.SetCssClass(this.Layer, "scheda");
    }

    #region data bounds

    protected void RicambiUtilizzatiLista_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
        {
            HtmlGenericControl codiceHd = (HtmlGenericControl)e.Row.FindControl("codiceHd");
            HtmlGenericControl ricambioHd = (HtmlGenericControl)e.Row.FindControl("ricambioHd");
            HtmlGenericControl prezzoHd = (HtmlGenericControl)e.Row.FindControl("prezzoHd");
            HtmlGenericControl quantitaHd = (HtmlGenericControl)e.Row.FindControl("quantitaHd");
            HtmlGenericControl scontoHd = (HtmlGenericControl)e.Row.FindControl("scontoHd");
            HtmlGenericControl totaleHd = (HtmlGenericControl)e.Row.FindControl("totaleHd");
            codiceHd.InnerText = "codice";
            ricambioHd.InnerText = "ricambio";
            quantitaHd.InnerText = "quantità";
            if (inGaranzia)
            {
                prezzoHd.Parent.Visible = false;
                scontoHd.Parent.Visible = false;
                totaleHd.Parent.Visible = false;
            }
            else
            {
                prezzoHd.InnerText = "prezzo";
                scontoHd.InnerText = "sconto";
                totaleHd.InnerText = "totale";
            }
        }

        else if (e.Row.RowType == DataControlRowType.DataRow)
        {
            HtmlGenericControl codice = (HtmlGenericControl)e.Row.FindControl("codice");
            HtmlGenericControl ricambio = (HtmlGenericControl)e.Row.FindControl("ricambio");
            HtmlGenericControl prezzo = (HtmlGenericControl)e.Row.FindControl("prezzo");
            HtmlGenericControl quantita = (HtmlGenericControl)e.Row.FindControl("quantita");
            HtmlGenericControl sconto = (HtmlGenericControl)e.Row.FindControl("sconto");
            HtmlGenericControl totale = (HtmlGenericControl)e.Row.FindControl("totale");
            DataRow dr = ((DataRowView)e.Row.DataItem).Row;
            codice.InnerText = dr["dts_ricambio_codice"].ToString();
            ricambio.InnerText = dr["dts_ricambio_nome"].ToString();
            quantita.InnerText = dr["dts_ricambio_qta"].ToString();
            if (inGaranzia)
            {
                prezzo.Parent.Visible = false;
                sconto.Parent.Visible = false;
                totale.Parent.Visible = false;
            }
            else
            {
                decimal prezzoRicambio = 0, totalePrezzo = 0;
                float scontoRicambio = 0;
                if (decimal.TryParse(dr["dts_ricambio_prezzo"].ToString(), out prezzoRicambio))
                    prezzo.InnerText = NextString.FormatEuro(prezzoRicambio.ToString());
                else
                    prezzo.Visible = false;
                if (float.TryParse(dr["dts_ricambio_sconto"].ToString(), out scontoRicambio))
                    sconto.InnerText = scontoRicambio.ToString() + " %";
                else
                    sconto.Visible = false;
                if (decimal.TryParse(dr["dts_prezzo_totale"].ToString(), out totalePrezzo))
                {
                    totale.InnerText = NextString.FormatEuro(totalePrezzo.ToString());
                    costoRicambi += totalePrezzo;
                }
                else
                    totale.Visible = false;
            }
        }
    }

    protected void Descrittori_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            HtmlGenericControl label = (HtmlGenericControl)e.Item.FindControl("DescrittoriLabel");
            HtmlGenericControl value = (HtmlGenericControl)e.Item.FindControl("DescrittoriValue");
            DataRow dr = ((DataRowView)e.Item.DataItem).Row;
            label.InnerText = dr["des_nome_it"].ToString() + ":";
            value.InnerText = (dr["rds_valore_it"].ToString() == "1" ? "Sì" : "No");
        }
    }

    #endregion

    protected void Invia_Click(object sender, EventArgs e)
    {
        int accessorioId = 0;
        string listaErrori = "", rifCliente = "", modello = "", matricola = "", negozioAcquisto = "",
               nScontrino = "", accessorioAltro = "", guastoAltro = "", noteCliente = "";
        DateTime dataAcquisto = DateTime.Today;

        #region valorizzazione campi e verifica riempimento
        if (cliente.Contatto.IsSocieta &&
            Request.Form[RiferimentoClienteInput.UniqueID] != null && Request.Form[RiferimentoClienteInput.UniqueID].ToString() != "")
            rifCliente = Request.Form[RiferimentoClienteInput.UniqueID].ToString();
        else if (cliente.Contatto.IsSocieta)
            listaErrori += "Campo obbligatorio 'riferimento cliente' non inserito.\n";

        if (ModelloInput.Visible && Request.Form[ModelloInput.UniqueID] != null && Request.Form[ModelloInput.UniqueID].ToString() != "")
            modello = Request.Form[ModelloInput.UniqueID].ToString();
        else if (ModelloInput.Visible)
            listaErrori += "Campo obbligatorio 'modello' non inserito.\n";

        if (Request.Form[MatricolaInput.UniqueID] != null && Request.Form[MatricolaInput.UniqueID].ToString() != "")
            matricola = Request.Form[MatricolaInput.UniqueID].ToString();
        else
            listaErrori += "Campo obbligatorio 'matricola' non inserito.\n";

        if (Request.Form[DataAcquistoInput.UniqueID] != null && Request.Form[DataAcquistoInput.UniqueID].ToString() != "")
            dataAcquisto = new DateTime(NextNumeric.ToInt(Request.Form[DataAcquistoInput.UniqueID].ToString().Split('/')[2]),
                                        NextNumeric.ToInt(Request.Form[DataAcquistoInput.UniqueID].ToString().Split('/')[1]),
                                        NextNumeric.ToInt(Request.Form[DataAcquistoInput.UniqueID].ToString().Split('/')[0]));
        else
            listaErrori += "Campo obbligatorio 'data acquisto' non inserito.\n";

        if (Request.Form[NegozioAcquistoInput.UniqueID] != null && Request.Form[NegozioAcquistoInput.UniqueID].ToString() != "")
            negozioAcquisto = Request.Form[NegozioAcquistoInput.UniqueID].ToString();

        if (Request.Form[NumeroScontrinoInput.UniqueID] != null && Request.Form[NumeroScontrinoInput.UniqueID].ToString() != "")
            nScontrino = Request.Form[NumeroScontrinoInput.UniqueID].ToString();
        else
            listaErrori += "Campo obbligatorio 'numero scontrino' non inserito.\n";

        if (Request.Form[AccessoriListaDdl.UniqueID] != null && Request.Form[AccessoriListaDdl.UniqueID].ToString() != "" &&
            NextNumeric.ToInt(Request.Form[AccessoriListaDdl.UniqueID].ToString()) > 0)
            accessorioId = NextNumeric.ToInt(Request.Form[AccessoriListaDdl.UniqueID].ToString());
        else if (Request.Form[AccessoriAltroInput.UniqueID] != null && Request.Form[AccessoriAltroInput.UniqueID].ToString() != "")
            accessorioAltro = Request.Form[AccessoriAltroInput.UniqueID].ToString();
        else
            listaErrori += "Campo obbligatorio 'accessorio' non inserito.\n";

        if (probId == 0 &&
            Request.Form[GuastoSegnalatoAltroInput.UniqueID] != null && Request.Form[GuastoSegnalatoAltroInput.UniqueID].ToString() != "")
            guastoAltro = Request.Form[GuastoSegnalatoAltroInput.UniqueID].ToString();
        else if (probId == 0)
            listaErrori += "Campo obbligatorio 'guasto segnalato' non inserito.\n";

        if (Request.Form[NoteClienteTxtarea.UniqueID] != null && Request.Form[NoteClienteTxtarea.UniqueID].ToString() != "")
            noteCliente = Request.Form[NoteClienteTxtarea.UniqueID].ToString();
        if (Request.Form[GaranziaCb.UniqueID] != null)
            reqGaranzia = GaranziaCb.Checked;
        #endregion

        // inserimento scheda se non mancano campi obbligatori
        if (String.IsNullOrEmpty(listaErrori))
        {
            schedaId = InfoschedeTools.InsertRichiesta(StnStatoSchedaId, cliente.Id, artId, 0, modello, matricola,
                                                       dataAcquisto, negozioAcquisto, nScontrino, true, probId,
                                                       guastoAltro, accessorioId, accessorioAltro, noteCliente, rifCliente);
            if (schedaId > 0)
            {
                queryString = "SCHEDAID=" + schedaId + "&CONFERMA=1" +
                              (Request.QueryString["MARID"] != null ? "&MARID=" + Request.QueryString["MARID"].ToString() : "");

                // email conferma
                NextPage.Alert.Email.SendPageFromAdminToContact("Richiesta di assistenza inviata correttamente", "",
                                                                NextPage.GetPageSiteUrlRedirect(StnPaginaEmailId, queryString),
                                                                StnAdminId, cliente.Contatto, "", true, true);
                // pagina conferma
                Response.Redirect(NextPage.GetPageSiteUrlRedirect(StnPaginaSchedaId, queryString));
            }
            else
                Response.Redirect(NextPage.UrlHomePage);
        }
        else
        {
            ErroriListaDiv.Visible = true;
            ErroriListaDiv.InnerHtml = NextString.HtmlEncode(listaErrori);
        }
    }
    
    protected void Indietro_Click(object sender, EventArgs e)
    {
        queryString = "";
        if (!string.IsNullOrEmpty(Request.QueryString["MARID"]))
            queryString = "MARID=" + Request.QueryString["MARID"].ToString();
        Response.Redirect(NextPage.GetPageSiteUrlRedirect(StnPaginaElencoId, queryString));
    }
}
