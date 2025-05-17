using System;
using System.Data;
using System.Configuration;
using System.Collections.Generic;
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

public partial class Plugin_RicercaRichieste : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente autenticato
    /// </summary>
    protected NextMembershipRivenditore cliente = NextMembershipRivenditore.CurrentCliente;

    #region Settings...

    /// <summary>
    /// Tipo di utente da associare all'elenco, valori accettati: supervisore,costruttore,altro.
    /// </summary>
    public string StnTipoUtente;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        try { StnTipoUtente = _settings["TipoUtente"]; }
        catch { StnTipoUtente = "altro"; };

        #endregion

        base.OnLoad(e);

        Titolo.InnerText = "Ricerca per:";
        NumeroSchedaLabel.InnerText = "N. scheda";
        StatoLabel.InnerText = "Stato della richiesta";
        MarcaLabel.InnerText = "Marca";
        ModelloLabel.InnerText = "Modello";
        RifClienteLabel.InnerText = "Riferimento cliente";
        NumeroDdtCaricoLabel.InnerText = "N. DDT di carico";
        DataRichiestaDaLabel.InnerText = "Data richiesta da";
        DataRichiestaALabel.InnerText = "Data richiesta a";
        RivenditoreLabel.InnerText = "Rivenditore";
        Cerca.Text = "Cerca";
        ViewAll.Text = "Vedi tutte";

        DataTable dt;
        string sql = "select sts_id, sts_nome_it from sgtb_stati_schede order by sts_nome_it";

        dt = NextPage.Connection.GetDataTable(sql);
        StatoDropDown.DataSource = dt;
        StatoDropDown.DataTextField = "sts_nome_it";
        StatoDropDown.DataValueField = "sts_id";
        StatoDropDown.DataBind();
        StatoDropDown.Items.Insert(0, new ListItem(String.Empty, String.Empty));

        if (StnTipoUtente.Equals("supervisore"))
        {
            sql = "select riv_id, (case issocieta when 1 then nomeorganizzazioneelencoindirizzi " +
                         "else nomeelencoindirizzi + ' ' + cognomeelencoindirizzi end) as nome " +
                         "from gtb_rivenditori left join tb_utenti on ut_id=riv_id left join tb_Indirizzario on idelencoindirizzi=ut_nextcom_id " +
                         "where riv_azienda_capogruppo_id is not null and riv_azienda_capogruppo_id<>riv_id and " +
                         " riv_azienda_capogruppo_id=" + NextMembershipRivenditore.Current.Id + " order by nome";

            dt = NextPage.Connection.GetDataTable(sql);
            RivenditoreDropDown.DataSource = dt;
            RivenditoreDropDown.DataTextField = "nome";
            RivenditoreDropDown.DataValueField = "riv_id";
            RivenditoreDropDown.DataBind();
            RivenditoreDropDown.Items.Insert(0, new ListItem(String.Empty, String.Empty));

            sql = "select mar_id,mar_nome_it from gtb_marche order by mar_nome_it";

            dt = NextPage.Connection.GetDataTable(sql);
            MarcaDropDown.DataSource = dt;
            MarcaDropDown.DataTextField = "mar_nome_it";
            MarcaDropDown.DataValueField = "mar_id";
            MarcaDropDown.DataBind();
            MarcaDropDown.Items.Insert(0, new ListItem(String.Empty, String.Empty));

            RifClienteDiv.Visible = false;
            NumeroDdtCaricoDiv.Visible = false;
        }

        else if (StnTipoUtente.Equals("costruttore"))
        {
            RivenditoreDiv.Visible = false;
            RifClienteDiv.Visible = false;
            NumeroDdtCaricoDiv.Visible = false;
            MarcaDiv.Visible = false;
        }

        else if (StnTipoUtente.Equals("altro"))
        {
            DataRichiestaADiv.Visible = false;
            DataRichiestaDaDiv.Visible = false;
            RivenditoreDiv.Visible = false;
            MarcaDiv.Visible = false;
        }

        // imposta eventualmente valori da ricerca precedente
        if (Session["numeroScheda"] != null && Session["numeroScheda"].ToString() != "")
            NumeroSchedaInput.Value = Session["numeroScheda"].ToString();
        if (Session["stato"] != null && Session["stato"].ToString() != "")
            StatoDropDown.SelectedValue = Session["stato"].ToString();
        if (Session["marca"] != null && Session["marca"].ToString() != "")
            MarcaDropDown.SelectedValue = Session["marca"].ToString();
        if (Session["modello"] != null && Session["modello"].ToString() != "")
            ModelloInput.Value = Session["modello"].ToString();
        if (Session["dataRichiestaDa"] != null && Session["dataRichiestaDa"].ToString() != "")
            DataRichiestaDaInput.Value = Session["dataRichiestaDa"].ToString();
        if (Session["dataRichiestaA"] != null && Session["dataRichiestaA"].ToString() != "")
            DataRichiestaAInput.Value = Session["dataRichiestaA"].ToString();
        if (Session["rivenditore"] != null && Session["rivenditore"].ToString() != "")
            RivenditoreDropDown.SelectedValue = Session["rivenditore"].ToString();
        if (Session["riferimentoCliente"] != null && Session["riferimentoCliente"].ToString() != "")
            RifClienteInput.Value = Session["riferimentoCliente"].ToString();
        if (Session["numeroDdtCarico"] != null && Session["numeroDdtCarico"].ToString() != "")
            NumeroDdtCaricoInput.Value = Session["numeroDdtCarico"].ToString();


        #region script datepicker

        string defaultDataString = "0",
               daysOfWeek = _nextLanguage.ChooseString("'do','lu','ma','me','gi','ve','sa'", "'su','mo','tu','we','th','fr','sa'"),
               datepicker_options = " showOn: 'both', \n" +
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
                                    " defaultDate: '" + defaultDataString + "' \n",
               datepicker_prefix = " var dates = ";

        NextPage.JQueryReadyManager.ApplyPluginDatePicker("#" + DataRichiestaDaInput.ClientID,
                                                          datepicker_options, datepicker_prefix);
        NextPage.JQueryReadyManager.ApplyPluginDatePicker("#" + DataRichiestaAInput.ClientID,
                                                          datepicker_options, datepicker_prefix);

        #endregion

        NextControlsTools.SetCssClass(this.Layer, "ricercarichieste");
    }

    protected void Cerca_Click(object sender, EventArgs e)
    {
        string filtro = "";
        if (((Button)sender).ID == "Cerca")
        {
            // validazione
            if (NumeroSchedaInputValid.IsValid && ModelloInputValid.IsValid)
            {
                if (Request.Form[NumeroSchedaInput.UniqueID] != null && Request.Form[NumeroSchedaInput.UniqueID].ToString() != "")
                {
                    filtro += " AND sc_numero = " + Request.Form[NumeroSchedaInput.UniqueID].ToString();
                    Session["numeroScheda"] = Request.Form[NumeroSchedaInput.UniqueID].ToString();
                }
                else
                    Session["numeroScheda"] = "";
                if (Request.Form[StatoDropDown.UniqueID] != null && Request.Form[StatoDropDown.UniqueID].ToString() != "")
                {
                    Session["stato"] = Request.Form[StatoDropDown.UniqueID].ToString();
                }
                else
                    Session["stato"] = "";
                if (Request.Form[MarcaDropDown.UniqueID] != null && Request.Form[MarcaDropDown.UniqueID].ToString() != "")
                {
                    filtro += " AND mar_id = " + Request.Form[MarcaDropDown.UniqueID].ToString();
                    Session["marca"] = Request.Form[MarcaDropDown.UniqueID].ToString();
                }
                else
                    Session["marca"] = "";
                if (Request.Form[ModelloInput.UniqueID] != null && Request.Form[ModelloInput.UniqueID].ToString() != "")
                {
                    List<String> listaCampi = new List<String>(new String[] {"art_nome_it", "art_nome_en", "art_cod_int",
                                                                             "art_cod_pro", "art_cod_alt"});
                    filtro += " AND EXISTS (SELECT art_id" +
                                            " FROM gtb_articoli" +
                                      " INNER JOIN grel_art_valori ON rel_art_id = art_id" +
                                           " WHERE rel_id = sc_modello_id" +
                                             " AND " + NextSql.TextSearch(Request.Form[ModelloInput.UniqueID].ToString(),
                                                                          listaCampi, true) + ")";
                    Session["modello"] = Request.Form[ModelloInput.UniqueID].ToString();
                }
                else
                    Session["modello"] = "";
                if (Request.Form[RifClienteInput.UniqueID] != null && Request.Form[RifClienteInput.UniqueID].ToString() != "")
                {
                    filtro += " AND " + NextSql.TextSearch(Request.Form[RifClienteInput.UniqueID].ToString(), "sc_rif_cliente", true);
                    Session["riferimentoCliente"] = Request.Form[RifClienteInput.UniqueID].ToString();
                }
                else
                    Session["riferimentoCliente"] = "";
                if (Request.Form[NumeroDdtCaricoInput.UniqueID] != null && Request.Form[NumeroDdtCaricoInput.UniqueID].ToString() != "")
                {
                    filtro += " AND " + NextSql.TextSearch(Request.Form[NumeroDdtCaricoInput.UniqueID].ToString(), "sc_numero_DDT_di_carico", true);
                    Session["numeroDdtCarico"] = Request.Form[NumeroDdtCaricoInput.UniqueID].ToString();
                }
                else
                    Session["numeroDdtCarico"] = "";
                if (Request.Form[DataRichiestaDaInput.UniqueID] != null && Request.Form[DataRichiestaDaInput.UniqueID].ToString() != "")
                {
                    string data = Request.Form[DataRichiestaDaInput.UniqueID].ToString();
                    filtro += " AND sc_data_ricevimento >= '" + data + "'";
                    Session["dataRichiestaDa"] = data;
                }
                else
                    Session["dataRichiestaDa"] = "";
                if (Request.Form[DataRichiestaAInput.UniqueID] != null && Request.Form[DataRichiestaAInput.UniqueID].ToString() != "")
                {
                    string data = Request.Form[DataRichiestaAInput.UniqueID].ToString();
                    filtro += " AND sc_data_ricevimento <= '" + data + "'";
                    Session["dataRichiestaA"] = data;
                }
                else
                    Session["dataRichiestaA"] = "";
                if (Request.Form[RivenditoreDropDown.UniqueID] != null && Request.Form[RivenditoreDropDown.UniqueID].ToString() != "")
                {
                    filtro += " AND sc_cliente_id = " + Request.Form[RivenditoreDropDown.UniqueID].ToString();
                    Session["rivenditore"] = Request.Form[RivenditoreDropDown.UniqueID].ToString();
                }
                else
                    Session["rivenditore"] = "";

                Session["filtroRichieste"] = filtro;

                Response.Redirect(NextPage.Url);
            }
            else
            {
                ErroriListaDiv.Visible = true;
                if (!NumeroSchedaInputValid.IsValid)
                    ErroriListaDiv.InnerHtml += NextString.HtmlEncode("\n" + NumeroSchedaLabel.InnerText + ": " +
                                                NumeroSchedaInputValid.ErrorMessage + "\n");
                if (!ModelloInputValid.IsValid)
                    ErroriListaDiv.InnerHtml += NextString.HtmlEncode("\n" + ModelloLabel.InnerText + ": " +
                                                ModelloInputValid.ErrorMessage + "\n");
                if (!RifClienteInputValid.IsValid)
                    ErroriListaDiv.InnerHtml += NextString.HtmlEncode("\n" + RifClienteLabel.InnerText + ": " +
                                                RifClienteInputValid.ErrorMessage + "\n");
                if (!NumeroDdtCaricoInputValid.IsValid)
                    ErroriListaDiv.InnerHtml += NextString.HtmlEncode("\n" + NumeroDdtCaricoLabel.InnerText + ": " +
                                                NumeroDdtCaricoInputValid.ErrorMessage + "\n");
            }
        }

        else
        {
            Session["numeroScheda"] = "";
            Session["stato"] = "";
            Session["marchio"] = "";
            Session["modello"] = "";
            Session["numeroDdtCarico"] = "";
            Session["filtroRichieste"] = "";
            Session["dataRichiestaA"] = "";
            Session["dataRichiestaDa"] = "";
            Session["filtroRichieste"] = "";
            Response.Redirect(NextPage.Url);
        }
    }
}
