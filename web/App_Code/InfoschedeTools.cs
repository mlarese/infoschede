using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using NextFramework;
using NextFramework.NextB2B;
using NextFramework.NextWeb;
using NextFramework.NextPassport;
using NextFramework.NextControls;


/// <summary>
/// Classe libreria di funzioni specifiche per infoschede
/// </summary>
public static class InfoschedeTools
{
    /// <summary>
    /// Codice modello di default (generico).
    /// </summary>
    public const string CodModelloDefault = "modello_default";
    /// <summary>
    /// Codice categoria di default (generico - relativo al modello di default).
    /// </summary>
    public const string CodategoriaDefault = "categoria_default";

    /// <summary>
    /// Percentuale di iva applicata.
    /// </summary>
    public const decimal IvaApplicata = 22;

    /// <summary>
    /// Definizione delle tipologie di messaggio visualizzabili nella bacheca generale.
    /// </summary>
    public struct Bacheca_TipoMessaggio
    {
        /// <summary>
        /// Messaggio generico.
        /// </summary>
        public const string Default = "default";
        /// <summary>
        /// Messaggio di tipo errore.
        /// </summary>
        public const string Errore = "errore";
        /// <summary>
        /// Messaggio di avviso generico.
        /// </summary>
        public const string Avviso = "avviso";
        /// <summary>
        /// Messaggio di ringraziamento.
        /// </summary>
        public const string Ok = "ok";
        /// <summary>
        /// Indicazione di aiuto/guida all'utente.
        /// </summary>
        public const string Help = "help";
    }


    /// <summary>
    /// Dichiarazione delle immagini di default in caso l'articolo non abbia immagini.
    /// </summary>
    public static class ImmaginiDefault
    {
        /// <summary>
        /// Immagine di default per il thumb
        /// </summary>
        public static string Thumb = NextFramework.NextB2B.ConfigurationNextB2B.NoImageAvailableThumb;

        /// <summary>
        /// Immagine di default per lo zoom.
        /// </summary>
        public static string Zoom = NextFramework.NextB2B.ConfigurationNextB2B.NoImageAvailableZoom;
    }

    public static void Bacheca_PubblicaMessaggio(string tipo, string titoloMessaggio, string corpoMessaggio)
    {}

    public static void Bacheca_PubblicaMessaggioInSessione(string tipo, string titoloMessaggio, string corpoMessaggio, NextUserControl sender)
    {}


    /// <summary>
    /// Restituisce un data table con i dati delle schede recuperate a seconda dei filtri.
    /// </summary>
    /// <param name="schedaId">Id scheda per cui filtrare l'elenco.</param>
    /// <param name="clienteId">Id cliente per cui filtrare l'elenco.</param>
    /// <param name="statiLista">Lista di id stato scheda separati da una virgola per cui filtrare l'elenco.</param>
    /// <param name="filtroAggiuntivo">Sql in aggiunta per cui filtrare l'elenco.</param>
    /// <param name="campiOrdinamento">Lista di campi per cui ordinare l'elenco.</param>
    /// <returns>DataTable contenente i dati scheda.</returns>
    public static DataTable GetSchedeDataTable(int schedaId, int clienteId, string statiLista,
                                               string filtroAggiuntivo, string campiOrdinamento)
    {
        return GetSchedeDataTable(schedaId, clienteId, statiLista, filtroAggiuntivo, campiOrdinamento, 0);
    }

    /// <summary>
    /// Restituisce un data table con i dati delle schede recuperate a seconda dei filtri.
    /// </summary>
    /// <param name="schedaId">Id scheda per cui filtrare l'elenco.</param>
    /// <param name="clienteId">Id cliente per cui filtrare l'elenco.</param>
    /// <param name="statiLista">Lista di id stato scheda separati da una virgola per cui filtrare l'elenco.</param>
    /// <param name="filtroAggiuntivo">Sql in aggiunta per cui filtrare l'elenco.</param>
    /// <param name="campiOrdinamento">Lista di campi per cui ordinare l'elenco.</param>
    /// <param name="richiestaRitiroId">Id richiesta di ritiro, per ulteriore filtro.</param>
    /// <returns>DataTable contenente i dati scheda.</returns>
    public static DataTable GetSchedeDataTable(int schedaId, int clienteId, string statiLista,
                                               string filtroAggiuntivo, string campiOrdinamento, int richiestaRitiroId)
    {
        string sql = "SELECT *" +
                          ", (CASE issocieta WHEN 1 THEN nomeorganizzazioneelencoindirizzi" +
                            " ELSE nomeelencoindirizzi + ' ' + cognomeelencoindirizzi END) AS nome_rivenditore" +
                          ", (SELECT sts_nome_it" +
                              " FROM sgtb_stati_schede" +
                             " WHERE sts_id = sc_stato_id) AS stato" +
                          ", (CASE WHEN art_id = " + GetModelloDefaultId() + " THEN sc_modello_altro" +
                            " ELSE (SELECT art_nome_it" +
                                    " FROM gtb_articoli" +
                              " INNER JOIN grel_art_valori ON rel_art_id = art_id" +
                                   " WHERE rel_id = sc_modello_id) END) AS modello" +
                          ", (SELECT art_cod_int" +
                              " FROM gtb_articoli" +
                        " INNER JOIN grel_art_valori ON rel_art_id = art_id" +
                             " WHERE rel_id = sc_modello_id) AS codice_modello" +
                          ", (CASE WHEN IsNull(sc_accessori_presenti_id,0)>0 then (SELECT acc_nome_it" +
                              " FROM sgtb_accessori" +
                             " WHERE acc_id = sc_accessori_presenti_id) ELSE sc_accessori_presenti_altro END) AS accessorio" +
                          ", (CASE WHEN IsNull(sc_guasto_segnalato_id,0)>0 then (SELECT prb_nome_it" +
                              " FROM sgtb_problemi" +
                             " WHERE prb_id = sc_guasto_segnalato_id) ELSE sc_guasto_segnalato_altro END) AS guasto_segnalato" +
                          ", (CASE WHEN IsNull(sc_guasto_riscontrato_id,0)>0 then (SELECT prb_nome_it" +
                              " FROM sgtb_problemi" +
                             " WHERE prb_id = sc_guasto_riscontrato_id) ELSE sc_guasto_riscontrato_altro END) AS guasto_riscontrato" +
                          ", (SELECT esi_nome_it" +
                              " FROM sgtb_esiti" +
                             " WHERE esi_id = sc_esito_intervento_id) AS esito_intervento" +
                          ", (SELECT ddt_numero" +
                              " FROM sgtb_ddt" +
                             " WHERE ddt_id = sc_rif_DDT_di_resa_id) AS numero_ddt" +
                          ", (SELECT ddt_data" +
                              " FROM sgtb_ddt" +
                             " WHERE ddt_id = sc_rif_DDT_di_resa_id) AS data_ddt" +
                          ", (SELECT ModoRegistra" +
                              " FROM sgtb_ddt" +
                        " INNER JOIN tb_Utenti ON ut_ID = ddt_trasportatore_id" +
                        " INNER JOIN tb_Indirizzario ON IDElencoIndirizzi = ut_NextCom_ID" +
                             " WHERE ddt_id = sc_rif_DDT_di_resa_id) AS trasportatore_ddt" +
                          ", (SELECT ddt_trasportatore_id" +
                              " FROM sgtb_ddt" +
                             " WHERE ddt_id = sc_rif_DDT_di_resa_id) AS trasportatore_ddt_id" +
                          ", (SELECT (CASE issocieta WHEN 1 THEN nomeorganizzazioneelencoindirizzi" +
                                    " ELSE nomeelencoindirizzi + ' ' + cognomeelencoindirizzi END) AS ca" +
                              " FROM tb_Utenti" +
                        " INNER JOIN tb_Indirizzario ON IDElencoIndirizzi = ut_NextCom_ID" +
                             " WHERE ut_id = sc_centro_assistenza_id) AS centro_assistenza" +
                      " FROM sgtb_schede" +
                 " LEFT JOIN grel_art_valori ON rel_id = sc_modello_id" +
                 " LEFT JOIN gtb_articoli ON rel_art_id = art_id" +
                 " LEFT JOIN gtb_marche ON art_marca_id = mar_id " +
                 " LEFT JOIN gtb_rivenditori ON riv_id = sc_cliente_id" +
                 " LEFT JOIN tb_utenti ON sc_cliente_id = ut_id " +
                 " LEFT JOIN tb_Indirizzario ON ut_NextCom_ID = IDElencoIndirizzi" +
                     " WHERE (1=1)";

        if (schedaId > 0)
            sql += " AND sc_id = " + schedaId;

        if (clienteId > 0)
            sql += " AND sc_cliente_id = " + clienteId;

        if (richiestaRitiroId > 0)
            sql += " AND sc_documento_ritiro_id = " + richiestaRitiroId;
        
        if (!String.IsNullOrEmpty(statiLista))
            sql += " AND sc_stato_id IN (" + statiLista + ")";

        if (!String.IsNullOrEmpty(filtroAggiuntivo))
            sql += " " + filtroAggiuntivo;

        if (!String.IsNullOrEmpty(campiOrdinamento))
            sql += " ORDER BY " + campiOrdinamento;

        return NextPage.Current.Connection.GetDataTable(sql);
    }


    /// <summary>
    /// Restituisce un data table con i dati dei ddt recuperati a seconda dei filtri.
    /// </summary>
    /// <param name="ddtId">Id ddt per cui filtrare l'elenco.</param>
    /// <param name="trasportatoreId">Id trasportatore per cui filtrare l'elenco.</param>
    /// <param name="clienteId">Id cliente per cui filtrare l'elenco.</param>
    /// <param name="filtroAggiuntivo">Sql in aggiunta per cui filtrare l'elenco.</param>
    /// <param name="campiOrdinamento">Lista di campi per cui ordinare l'elenco.</param>
    /// <returns>DataTable contenente i dati ddt.</returns>
    public static DataTable GetDdtDataTable(int ddtId, int trasportatoreId, int clienteId,
                                            string filtroAggiuntivo, string campiOrdinamento)
    {
        string sql = "SELECT * " +
                      " FROM sgtb_ddt INNER JOIN sgtb_ddt_categorie ON sgtb_ddt.ddt_categoria_id = sgtb_ddt_categorie.cat_id " + 
                           " INNER JOIN gv_rivenditori ON sgtb_ddt.ddt_cliente_id = gv_rivenditori.riv_id " + 
                           " LEFT JOIN sgtb_ddt_causali ON sgtb_ddt.ddt_causale_id = sgtb_ddt_causali.cau_id " +
                           " LEFT JOIN sgtb_ddt_porto ON sgtb_ddt.ddt_porto_id = sgtb_ddt_porto.por_id " + 
                           " LEFT JOIN sgtb_ddt_trasporto ON sgtb_ddt.ddt_trasporto_id = sgtb_ddt_trasporto.tra_id " +
                      " WHERE (1=1)";

        if (ddtId > 0)
            sql += " AND ddt_id = " + ddtId;

        if (trasportatoreId > 0)
            sql += " AND ddt_trasportatore_id = " + trasportatoreId;

        if (clienteId > 0)
            sql += " AND ddt_cliente_id = " + clienteId;

        if (!String.IsNullOrEmpty(filtroAggiuntivo))
            sql += " " + filtroAggiuntivo;

        if (!String.IsNullOrEmpty(campiOrdinamento))
            sql += " ORDER BY " + campiOrdinamento;

        return NextPage.Current.Connection.GetDataTable(sql);
    }


    /// <summary>
    /// Inserisce a db una richiesta di assistenza.
    /// </summary>
    /// <param name="statoId"></param>
    /// <param name="clienteId"></param>
    /// <param name="artId"></param>
    /// <param name="matricola"></param>
    /// <param name="statoId"></param>
    /// <param name="dataAcquisto"></param>
    /// <param name="nScontrino"></param>
    /// <param name="probId"></param>
    /// <param name="accessorioAltro"></param>
    /// <param name="noteCliente"></param>
    /// <param name="rifCliente"></param>
    /// <returns>Id della richiesta di assistenza inserita.</returns>
    public static int InsertRichiesta(int statoId, int clienteId, int artId, int relId, string modelloGenerico,
                                      string matricola, DateTime dataAcquisto, string negozioAcquisto,
                                      string nScontrino, bool reqGaranzia, int guastoId, string guastoAltro,
                                      int accessorioId, string accessorioAltro, string noteCliente, string rifCliente)
    {
        SqlCommand command = (SqlCommand)NextPage.Current.Connection.CreateCommand();

        // id variante
        if (relId == 0)
        {
            command.CommandText = "SELECT TOP 1 rel_id" +
                                   " FROM grel_art_valori" +
                                  " WHERE rel_art_id = " + artId;
            relId = (int)NextPage.Current.Connection.ExecuteScalar(command);
        }

        // numero progressivo
        command.CommandText = "SELECT MAX(sc_numero) + 1" +
                               " FROM sgtb_schede";
        int numero = (int)NextPage.Current.Connection.ExecuteScalar(command);

        // inserimento richiesta di assistenza
        command.CommandText =
            "INSERT INTO sgtb_schede(sc_stato_id, sc_numero, sc_data_ricevimento, sc_cliente_id, sc_modello_id" +
                                  ", sc_modello_altro, sc_matricola, sc_data_acquisto, sc_negozio_acquisto" +
                                  ", sc_numero_scontrino, sc_guasto_segnalato_id, sc_guasto_segnalato_altro" +
                                  ", sc_accessori_presenti_id, sc_accessori_presenti_altro" +
                                  ", sc_note_cliente, sc_rif_cliente, sc_richiesta_garanzia, sc_in_garanzia)" +
                " VALUES (" + statoId + ", " + numero + ", @data_ricevimento, " + clienteId + ", " + relId +
                       ", '" + NextSql.EncodeSql(modelloGenerico) + "', '" + NextSql.EncodeSql(matricola) + "', @data_acquisto" + ", '" + NextSql.EncodeSql(negozioAcquisto) +
                      "', '" + NextSql.EncodeSql(nScontrino) + "', " + guastoId + ", '" + NextSql.EncodeSql(guastoAltro) + "', " + accessorioId +
                       ", '" + NextSql.EncodeSql(accessorioAltro) + "', '" + NextSql.EncodeSql(noteCliente) + "', '" + NextSql.EncodeSql(rifCliente) +
                      "', " + (reqGaranzia ? "1" : "0") + ", 0)";

        Database.AddParameter(command, "@data_acquisto", DbType.DateTime, dataAcquisto, false, null);
        Database.AddParameter(command, "@data_ricevimento", DbType.DateTime, DateTime.Today.Date, false, null);

        return (int)NextPage.Current.Connection.ExecuteWithId(command);
    }


    /// <summary>
    /// Restituisce l'id del modello di default partendo dal codice <remarks>CodModelloDefault</remarks>.
    /// </summary>
    /// <returns>Id del modello di default.</returns>
    public static int GetModelloDefaultId()
    {
        SqlCommand command = (SqlCommand)NextPage.Current.Connection.CreateCommand();
        command.CommandText = "SELECT art_id" +
                               " FROM gtb_articoli" +
                              " WHERE art_cod_int = '" + CodModelloDefault + "'";
        return (int)NextPage.Current.Connection.ExecuteScalar(command);
    }
}
