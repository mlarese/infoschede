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

public partial class Plugin_SchedaStampa : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente autenticato
    /// </summary>
    protected NextMembershipRivenditore cliente = NextMembershipRivenditore.CurrentCliente;

    int schedaId = 0, clienteId = 0;
    decimal costoPresa = 0, costoRiconsegna = 0, costoManodopera = 0, oreManodopera = 0, costoRicambi = 0, costoTotale = 0, ivaTotale = 0;

    DataTable dtScheda;
    string sql;
    bool inGaranzia = false;


    #region Settings...

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

        StnTitolo = NextLanguage.ChooseString(_settings, "Titolo");
        StnDescrizione = NextLanguage.ChooseString(_settings, "Descrizione");
        #endregion

        base.OnLoad(e);

        if (!string.IsNullOrEmpty(StnTitolo))
        {
            Titolo.Visible = true;
            Titolo.InnerText = StnTitolo;
        }
        if (!string.IsNullOrEmpty(StnDescrizione))
        {
            Descrizione.Visible = true;
            Descrizione.InnerText = StnDescrizione;
        }

        if (int.TryParse(Request.QueryString["SCHEDAID"].ToString(), out schedaId) && schedaId > 0)
        {
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
                dtScheda = InfoschedeTools.GetSchedeDataTable(schedaId, 0, "", "", "");
                if (dtScheda.Rows.Count > 0)
                {
                    DataRow drScheda = dtScheda.Rows[0];

                    NumeroValue.InnerText = drScheda["sc_numero"].ToString();
                    DataValue.InnerText = DateTime.Parse(drScheda["sc_data_ricevimento"].ToString()).ToString(NextDateTime.StringFormats.DateIta);
                    StatoValue.InnerText = drScheda["stato"].ToString();

                    DestinatarioNomeValue.InnerText = cliente.Contatto.GetName().ToUpper();
                    DestinatarioViaValue.InnerText = cliente.Contatto.Indirizzo;
                    DestinatarioCittaValue.InnerText = cliente.Contatto.Cap + " " + cliente.Contatto.Citta.ToUpper() +
                                                       " " + cliente.Contatto.Provincia.ToUpper();
                    if (NextNumeric.ToInt(drScheda["numero_ddt"]) > 0)
                    {
                        ConsegnaRifValue.InnerText = "n°" + drScheda["numero_ddt"].ToString() + " del " + DateTime.Parse(drScheda["data_ddt"].ToString()).ToString(NextDateTime.StringFormats.DateIta);
                        TrasportatoreValue.InnerText = drScheda["trasportatore_ddt"].ToString();
                    }

                    if (!string.IsNullOrEmpty(drScheda["sc_rif_cliente"].ToString()))
                        RitiroRifValue.InnerText = drScheda["sc_rif_cliente"].ToString();
                    if (!string.IsNullOrEmpty(drScheda["sc_data_ddt_di_carico"].ToString()) && !string.IsNullOrEmpty(drScheda["sc_numero_ddt_di_carico"].ToString()))
                        RitiroRifDdtValue.InnerText = "n°" + drScheda["sc_numero_ddt_di_carico"].ToString() + " del " + DateTime.Parse(drScheda["sc_data_ddt_di_carico"].ToString()).ToString(NextDateTime.StringFormats.DateIta);

                    CostruttoreValue.InnerText = drScheda["mar_nome_it"].ToString();
                    ModelloValue.InnerText = drScheda["modello"].ToString();
                    MatricolaValue.InnerText = drScheda["sc_matricola"].ToString();
                    if (!string.IsNullOrEmpty(drScheda["sc_data_Acquisto"].ToString()))
                        DataAcquistoValue.InnerText = DateTime.Parse(drScheda["sc_data_Acquisto"].ToString()).ToString(NextDateTime.StringFormats.DateIta);
                    ScontrinoValue.InnerText = drScheda["sc_numero_scontrino"].ToString();

                    inGaranzia = NextBoolean.ToBoolean(drScheda["sc_in_garanzia"]);
                    if (inGaranzia)
                        GaranziaValue.InnerText = "Si";
                    else
                        GaranziaValue.InnerText = "No";
                    if (!(Request.QueryString["SHOW_GARANZIA"] != null && Request.QueryString["SHOW_GARANZIA"].ToString().ToLower() == "true"))
                        GaranziaValue.InnerText = "Previa valutazione da parte del centro assistenza";

                    AccessoriValue.InnerText = drScheda["accessorio"].ToString();
                    GuastoSegnalatoValue.InnerText = drScheda["guasto_segnalato"].ToString();
                    GuastoRiscontratoValue.InnerText = drScheda["guasto_riscontrato"].ToString();
                    NoteClienteValue.InnerText = drScheda["sc_note_cliente"].ToString();
                    NoteRiparazioneValue.InnerText = drScheda["sc_note_chiusura"].ToString();
                    EsitoRiparazioneValue.InnerText = drScheda["esito_intervento"].ToString();
                    if (!string.IsNullOrEmpty(drScheda["sc_data_fine_lavoro"].ToString()))
                        DataFineLavoroValue.InnerText = DateTime.Parse(drScheda["sc_data_fine_lavoro"].ToString()).ToString(NextDateTime.StringFormats.DateIta);

                    // ricambi utilizzati
                    sql = "SELECT *" +
                           " FROM sgtb_dettagli_schede" +
                          " WHERE dts_scheda_id = " + schedaId;
                    DataTable dtRic = NextPage.Connection.GetDataTable(sql);
                    if (dtRic.Rows.Count > 0)
                    {
                        RicambiNessunoTr.Visible = false;
                        RicambiUtilizzatiLista.DataSource = dtRic;
                        RicambiUtilizzatiLista.DataBind();
                    }
                    else
                    {
                        RicambiListaTr.Visible = false;
                        RicambiNessunoTr.Visible = true;
                    }

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
                        DescrittoriTitleTr.Visible = false;
                        DescrittoriListaTr.Visible = false;
                    }

                    if (!decimal.TryParse(drScheda["sc_ora_manodopera_intervento"].ToString(), out oreManodopera))
                        oreManodopera = 0;
                    ManodoperaOreValue.InnerText = drScheda["sc_ora_manodopera_intervento"].ToString();

                    if (!inGaranzia)
                    {
                        if (decimal.TryParse(drScheda["sc_prezzo_manodopera"].ToString(), out costoManodopera) && costoManodopera > 0)
                            ManodoperaPrezzoValue.InnerText = NextNumeric.FormatEuroPrice(costoManodopera);
                        else
                            ManodoperaPrezzoValue.InnerText = NextString.FormatEuro("0");

                        if (costoManodopera > 0 && oreManodopera > 0)
                        {
                            ManodoperaTotaleValue.InnerText = NextNumeric.FormatEuroPrice(costoManodopera * oreManodopera);
                        }
                        
                        if (decimal.TryParse(drScheda["sc_costo_riconsegna"].ToString(), out costoRiconsegna) && costoRiconsegna>0)
                            CostoRiconsegnaValue.InnerText = NextNumeric.FormatEuroPrice(costoRiconsegna);
                        else
                            CostoRiconsegnaValue.InnerText = NextString.FormatEuro("0");

                        if (decimal.TryParse(drScheda["sc_costo_presa"].ToString(), out costoPresa) && costoPresa > 0)
                            CostoPresaValue.InnerText = NextNumeric.FormatEuroPrice(costoPresa);
                        else
                            CostoPresaValue.InnerText = NextString.FormatEuro("0");

                        if (costoRicambi > 0)
                            TotaleRicambiValue.InnerText = NextNumeric.FormatEuroPrice(costoRicambi);
                        else
                            TotaleRicambiValue.InnerText = NextString.FormatEuro("0");

                        costoTotale = costoRicambi + costoPresa + costoRiconsegna + (costoManodopera * oreManodopera);
                        ivaTotale = NextNumeric.Percentuale(costoTotale, InfoschedeTools.IvaApplicata);
                        if (costoTotale > 0)
                        {
                            TotaleIvaValue.InnerText = NextNumeric.FormatEuroPrice(ivaTotale);
                            TotaleGeneraleValue.InnerText = NextNumeric.FormatEuroPrice(ivaTotale + costoTotale);
                        }
                        else
                        {
                            TotaleGeneraleValue.InnerText = NextString.FormatEuro("0");
                            TotaleIvaValue.InnerText = NextString.FormatEuro("0");
                        }


                    }
                    else
                    {
                        TotaleRicambiTr.Visible = false;
                        TotaleGeneraleTr.Visible = false;
                        TotaleIvaTr.Visible = false;
                    }

                }
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

}
