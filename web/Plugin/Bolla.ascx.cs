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
using NextFramework.NextCom;

public partial class Plugin_Bolla : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente autenticato
    /// </summary>
    protected NextMembershipRivenditore cliente = NextMembershipRivenditore.CurrentCliente;

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
        NextControlsTools.SetCssClass(this.Layer, "modulostampa");
        
        base.OnLoad(e);
        //NextControlsTools.SetCssClass(this.Layer, "scheda");

        int ddtId = 0, clienteId = 0, trasportatoreId = 0, destinazioneId = 0;

        if ((cliente == null || cliente.IsPublic) && Request.QueryString["CLIENTEID"] != null &&
            int.TryParse(Request.QueryString["CLIENTEID"].ToString(), out clienteId) && clienteId > 0 &&
            Request.QueryString["KEY"] != null && Request.QueryString["IDCNT"] != null)
            if (!cliente.SetPropertiesById(clienteId) ||
                Request.QueryString["IDCNT"].ToString() != cliente.Contatto.Id.ToString() &&
                Request.QueryString["KEY"].ToString() != cliente.Contatto.CodiceInserimento)
                Response.Redirect(NextPage.UrlHomePage);

        if (Request.QueryString["DDTID"] != null &&
            int.TryParse(Request.QueryString["DDTID"].ToString(), out ddtId) && ddtId > 0)
        {
            // recupera dati
            DataTable dt = InfoschedeTools.GetDdtDataTable(ddtId, 0, 0, "", "");

            if (int.TryParse(dt.Rows[0]["ddt_cliente_id"].ToString(), out clienteId) && clienteId == cliente.Id)
            {
                DdtTitle.InnerText = "DOCUMENTO DI TRASPORTO BENI VIAGGIANTI SU AUTOMEZZI DEL MITTENTE O DESTINATARIO " +
                                     "O VETTORE (D.P.R. 472/96 del 14/08/96)";
                DdtNumeroLabel.InnerText = "Numero";
                DdtNumeroValue.InnerText = dt.Rows[0]["ddt_numero"].ToString();
                DdtDataLabel.InnerText = "Data";
                DdtDataValue.InnerText = DateTime.Parse(dt.Rows[0]["ddt_data"].ToString()).ToString(NextDateTime.StringFormats.DateIta);

                cliente.Contatto.SetRecapiti();
                ClienteLabel.InnerText = "Cliente";
                ClienteValue.InnerText = cliente.Id.ToString();
                PartitaIvaLabel.InnerText = "Partita Iva";
                PartitaIvaValue.InnerText = !String.IsNullOrEmpty(cliente.Contatto.Partita_iva) ? cliente.Contatto.Partita_iva : cliente.Contatto.CF;
                CodiceFiscaleLabel.InnerText = "Cod. Fiscale";
                CodiceFiscaleValue.InnerText = cliente.Contatto.CF;
                AgenteLabel.InnerText = "Agente";
                AgenteValue.InnerText = cliente.Riv_agente_id > 0 ? cliente.Riv_agente_id.ToString() : "";
                TelefonoLabel.InnerText = "Telefono / Note";
                TelefonoValue.InnerText = cliente.Contatto.Telefono;
                
                DestinatarioTitle.InnerText = "Spett.le";
                DestinatarioNomeValue.InnerText = cliente.Contatto.GetName().ToUpper();
                DestinatarioViaValue.InnerText = cliente.Contatto.Indirizzo;
                DestinatarioCittaValue.InnerText = cliente.Contatto.Cap + " " + cliente.Contatto.Citta.ToUpper() +
                                                   " " + cliente.Contatto.Provincia.ToUpper();

                IndirizzoTitle.InnerText = "Destinazione merce";
                if (dt.Rows[0]["ddt_destinazione_id"] != null &&
                    int.TryParse(dt.Rows[0]["ddt_destinazione_id"].ToString(), out destinazioneId) && destinazioneId > 0)
                    cliente.Contatto.SetContatto(destinazioneId);
                IndirizzoNomeValue.InnerText = cliente.Contatto.GetName().ToUpper();
                IndirizzoViaValue.InnerText = cliente.Contatto.Indirizzo;
                IndirizzoCittaValue.InnerText = cliente.Contatto.Cap + " " + cliente.Contatto.Citta.ToUpper() +
                                                " " + cliente.Contatto.Provincia.ToUpper();

                string filtro = "SELECT 'pz. ' + CONVERT(NVARCHAR(20), count(art_cod_int)) AS Quantità" +
                                     ", art_cod_int AS Modello" +
                                     ", art_nome_it AS Descrizione" +
                                     ", CONVERT(NVARCHAR(20), sc_numero) AS Rif_NS_SCHEDA" +
                                     ", sc_numero_DDT_di_carico AS Rif_VS_DDT" +
                                     ", CASE sc_in_garanzia WHEN 1 THEN 'Sì' ELSE 'No' END AS Garanzia" +
                                     ", '' AS [Prezzo Unitario]" + 
                                     ", '' AS [Sconto]" +
                                 " FROM gtb_articoli" +
                           " INNER JOIN grel_art_valori ON rel_art_id = art_id" +
                           " INNER JOIN sgtb_schede ON sc_modello_id = rel_id" +
                                " WHERE sc_cliente_id = " + clienteId +
                                  " AND sc_rif_DDT_di_resa_id = " + ddtId +
                             " GROUP BY art_cod_int" +
                                     ", art_nome_it" +
                                     ", sc_numero" +
                                     ", sc_numero_DDT_di_carico" +
                                     ", sc_in_garanzia" +
                            " UNION " +
                            " SELECT 'pz. ' + CONVERT(NVARCHAR(20), dtd_articolo_qta) AS Quantità" +
                                    ", dtd_articolo_codice AS Modello" +
                                    ", dtd_articolo_nome AS Descrizione" +
                                    ", '' AS Rif_NS_SCHEDA" +
                                    ", dtd_rif_vs_ddt AS Rif_VS_DDT" +
                                    ", CASE dtd_in_garanzia WHEN 1 THEN 'Sì' ELSE 'No' END AS Garanzia" +
                                    ", CONVERT(NVARCHAR(20), dtd_articolo_prezzo_unitario) + ' €' AS [Prezzo Unitario]" +
                                    ", (CASE ISNULL(dtd_articolo_sconto, 0) WHEN 0 THEN '' ELSE CONVERT(NVARCHAR(20), dtd_articolo_sconto) + ' %' END) AS [Sconto]" + 
                            " FROM sgtb_dettagli_ddt WHERE dtd_ddt_id = " + ddtId;

                DataTable modelDt = NextPage.Connection.GetDataTable(filtro);
                bool disable_prezzo = true;
                bool disable_sconto = true;
                if (modelDt.Rows.Count > 0)
                {
                    TabellaDati.DataSource = modelDt;
                    TabellaDati.DataBind();
                    int x = 0;
                    while (x < modelDt.Rows.Count && (disable_prezzo || disable_sconto))
                    {
                        if (modelDt.Rows[x]["Prezzo Unitario"].ToString() != "")
                            disable_prezzo = false;
                        if (modelDt.Rows[x]["Sconto"].ToString() != "")
                            disable_sconto = false;
                        x++;
                    }

                    //aggiungo le classi alle celle della tabella
                    TabellaDati.HeaderRow.Cells[0].CssClass = "qta";
                    TabellaDati.HeaderRow.Cells[1].CssClass = "modello";
                    TabellaDati.HeaderRow.Cells[2].CssClass = "descrizione";
                    TabellaDati.HeaderRow.Cells[3].CssClass = "rif_ns_scheda";
                    TabellaDati.HeaderRow.Cells[4].CssClass = "rif_vs_ddt";
                    TabellaDati.HeaderRow.Cells[5].CssClass = "garanzia";
                    TabellaDati.HeaderRow.Cells[6].CssClass = "prz_unit";
                    TabellaDati.HeaderRow.Cells[7].CssClass = "sconto";
                    if (disable_prezzo)
                        TabellaDati.HeaderRow.Cells[6].CssClass = "disabled";
                    if (disable_sconto)
                        TabellaDati.HeaderRow.Cells[7].CssClass = "disabled";
                    for (int i = 0; i < TabellaDati.Rows.Count; i++)
                    {
                        for (int j = 0; j < TabellaDati.Rows[i].Cells.Count; j++)
                        {
                            TabellaDati.Rows[i].Cells[j].CssClass = TabellaDati.HeaderRow.Cells[j].CssClass;
                        }
                    }

                }
                else
                    TabellaDiv.Visible = false;

                AnnotazioniLabel.InnerText = "Annotazioni";
                AnnotazioniValue.InnerHtml = dt.Rows[0]["ddt_note"].ToString().Replace("\r\n","<br>");
                DdtCausaleLabel.InnerText = "Causale del trasporto";
                DdtCausaleValue.InnerText = dt.Rows[0]["cau_titolo_it"].ToString();
                DataOraTrasportoLabel.InnerText = "Data ed ora trasporto";
                DataOraTrasportoValue.InnerText = "";
                FirmaConducenteLabel.InnerText = "Firma del conducente";
                TrasportoCuraLabel.InnerText = "Trasporto a cura";
                TrasportoCuraValue.InnerText = dt.Rows[0]["tra_titolo_it"].ToString(); ;
                PortoLabel.InnerText = "Porto";
                PortoValue.InnerText = dt.Rows[0]["por_titolo_it"].ToString();
                ColliLabel.InnerText = "Colli";
                ColliValue.InnerText = dt.Rows[0]["ddt_numero_colli"].ToString();
                PesoLabel.InnerText = "Peso";
                PesoValue.InnerText = dt.Rows[0]["ddt_peso"].ToString();
                FirmaDestinatarioLabel.InnerText = "Firma del destinatario";
                VettoreLabel.InnerText = "Vettori (ditta - residenza/domicilio)";
                DataOraRitiroLabel.InnerText = "Data ed ora ritiro";
                DataOraRitiroValue.InnerHtml = "&nbsp;&nbsp;&nbsp;&nbsp;";
                FirmaVettoreLabel.InnerText = "Firma del vettore";
                FirmaVettoreValue.InnerHtml = "&nbsp;&nbsp;&nbsp;&nbsp;";

                NextMembershipRivenditore trasportatore = new NextMembershipRivenditore();
                if (int.TryParse(dt.Rows[0]["ddt_trasportatore_id"].ToString(), out trasportatoreId) &&
                    trasportatore.SetPropertiesById(trasportatoreId))
                    VettoreValue.InnerHtml = NextString.HtmlEncode(trasportatore.Contatto.GetName().ToUpper() + "\n" +
                                                                   trasportatore.Contatto.BLL.GetAddress() + "\n");

                Note.InnerText = "La merce viaggia a rischio e pericolo del committente anche se venduta franco destino. " +
                                 "Gli eventuali reclami debbonsi fare entro 3 giorni dal ricevimento della merce " +
                                 "con raccomandata R.R.; i ritorni di merce non autorizzati verranno respinti. " +
                                 "Danni causati dal trasporto dovranno esssere contestati alla " +
                                 "Hidroservices s.a.s. all'atto del ricevimento della merce.";
            }
            else
                Response.Redirect(NextPage.UrlHomePage);
        }
    }
}