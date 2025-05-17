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

public partial class Plugin_LetteraAccompagnamento : NextFramework.NextControls.NextUserControl
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

            if (int.TryParse(dt.Rows[0]["ddt_cliente_id"].ToString(), out clienteId) && clienteId == cliente.Id &&
                int.TryParse(dt.Rows[0]["ddt_trasportatore_id"].ToString(), out trasportatoreId) && trasportatoreId > 0)
            {
                NextMembershipRivenditore trasportatore = new NextMembershipRivenditore();
                if (trasportatore.SetPropertiesById(trasportatoreId))
                {
                    /*TestataTitolo.InnerText = "Hidroservices s.a.s.";
                    TestataLink.InnerText = "www.assistenza360gradi.com";
                    TestataLink.HRef = TestataLink.InnerText;
                    Indirizzo.InnerText = "Via Vittorio Veneto, 2/a";
                    Citta.InnerText = "30030 Salzano (VE)";
                    Fax.InnerText = "Fax 041 484691";
                    Email.InnerText = "E-Mail: infoservices@hidroservices.it";
                    Email.HRef = "mailto:infoservices@hidroservices.it";
                    PartitaIva.InnerText = "P.IVA 02485200279";
                    Telefono.InnerText = "Tel. 041.484691";*/


                    //TitoloLabel.InnerText = "Documento di trasporto :";
                    //DdtNumeroLabel.InnerText = "ddt n.";
                    TitoloLabel.InnerText = "Lettera d'accompagnamento ";
                    DdtNumeroLabel.InnerText = "n.";
                    DdtNumeroValue.InnerText = dt.Rows[0]["ddt_numero"].ToString();
                    DdtDataLabel.InnerText = "Data";
                    DdtDataValue.InnerText = DateTime.Parse(dt.Rows[0]["ddt_data"].ToString()).ToString(NextDateTime.StringFormats.DateIta);
                    
                    if (dt.Rows[0]["ddt_destinazione_id"] != null &&
                        int.TryParse(dt.Rows[0]["ddt_destinazione_id"].ToString(), out destinazioneId) && destinazioneId > 0)
                        cliente.Contatto.SetContatto(destinazioneId);
                    DestinatarioNomeValue.InnerText = cliente.Contatto.GetName().ToUpper();
                    DestinatarioNomeLabel.InnerText = "Destinatario";
                    NumeroOrdineLabel.InnerText = "numero ordine";
                    DestinatarioViaLabel.InnerText = "via";
                    DestinatarioViaValue.InnerText = cliente.Contatto.Indirizzo.ToUpper();
                    DestinatarioCapLabel.InnerText = "CAP";
                    DestinatarioCapValue.InnerText = cliente.Contatto.Cap;
                    DestinatarioCittaLabel.InnerText = "città";
                    DestinatarioCittaValue.InnerText = cliente.Contatto.Citta.ToUpper();
                    DestinatarioProvinciaValue.InnerText = cliente.Contatto.Provincia.ToUpper();
                    DestinatarioLocalitaLabel.InnerText = "località";
                    DdtCausaleLabel.InnerText = "Causale :";
                    DdtCausaleValue.InnerText = dt.Rows[0]["cau_titolo_it"].ToString();

                    string filtro = "SELECT art_cod_int AS Codice" +
		                                 ", art_nome_it AS Descrizione" +
		                                 ", count(art_cod_int) AS Quantità" +
	                                 " FROM gtb_articoli" +
                               " INNER JOIN grel_art_valori ON rel_art_id = art_id" +
                               " INNER JOIN sgtb_schede ON sc_modello_id = rel_id" +
                                    " WHERE sc_cliente_id = " + clienteId +
                                      " AND sc_rif_DDT_di_resa_id = " + ddtId +
                                 " GROUP BY art_cod_int" +
                                         ", art_nome_it" +
                            " UNION " +
                            " SELECT dtd_articolo_codice AS Codice" +
                                    ", dtd_articolo_nome AS Descrizione" +
                                    ", CONVERT(NVARCHAR(20), dtd_articolo_qta) AS Quantità" +
                            " FROM sgtb_dettagli_ddt WHERE dtd_ddt_id = " + ddtId;
                    DataTable modelDt = NextPage.Connection.GetDataTable(filtro);
                    if (modelDt.Rows.Count > 0)
                    {
                        TabellaDati.DataSource = modelDt;
                        TabellaDati.DataBind();
                    }
                    else
                        TabellaDiv.Visible = false;

                    TrasportatoreLabel.InnerText = "Trasporto a cura di";
                    TrasportatoreValue.InnerText = trasportatore.Contatto.GetName().ToUpper();
                    DataRitiroLabel.InnerText = "Data ritiro";
                    FirmaConducente.InnerText = "FIRMA CONDUCENTE";
                    FirmaVettore.InnerText = "FIRMA VETTORE";
                    FirmaDestinatario.InnerText = "FIRMA DESTINATARIO";
                }
            }
            else
                Response.Redirect(NextPage.UrlHomePage);
        }
    }
}
