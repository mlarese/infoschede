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

public partial class Plugin_ElencoRichieste : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente autenticato
    /// </summary>
    protected NextMembershipRivenditore cliente = NextMembershipRivenditore.CurrentCliente;

    int modelloDefaultId = InfoschedeTools.GetModelloDefaultId();

    #region Settings...

    /// <summary>
    /// Pagina scheda.
    /// </summary>
    public int StnPaginaSchedaId;

    /// <summary>
    /// Titolo.
    /// </summary>
    public string StnTitolo;

    /// <summary>
    /// Descrizione.
    /// </summary>
    public string StnDescrizione;

    /// <summary>
    /// Tipo di utente da associare all'elenco, valori accettati: supervisore,costruttore,altro.
    /// </summary>
    public string StnTipoUtente;

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        try { StnPaginaSchedaId = int.Parse(_settings["paginaSchedaId"]); } catch { };

        try { StnTitolo = _settings["Titolo"]; } catch { };
        if (String.IsNullOrEmpty(StnTitolo))
            StnTitolo = "Richieste di assistenza già effettuate";

        try { StnDescrizione = _settings["Descrizione"]; } catch { };
        if (String.IsNullOrEmpty(StnDescrizione))
            StnDescrizione = "Qui potrai controllare lo stato di avanzamento delle richieste di assistenza già effettuate ma non ancora portate a compimento.";

        try { StnTipoUtente = _settings["TipoUtente"]; }
        catch { StnTipoUtente = "altro"; };

        #endregion

        base.OnLoad(e);

        if (Session["filtroRichieste"] == null) Session["filtroRichieste"] = "";
        if (Session["sortExpression"] == null) Session["sortExpression"] = "";
        if (Session["stato"] == null) Session["stato"] = "";

        if (cliente != null && !cliente.IsPublic)
        {
            Titolo.InnerText = StnTitolo;
            Descrizione.InnerText = StnDescrizione;

            DataTable dt;
            if (StnTipoUtente.Equals("costruttore"))
            {
                dt = InfoschedeTools.GetSchedeDataTable(0, 0, Session["stato"].ToString(),
                                                        Session["filtroRichieste"].ToString() +
                                                        " AND mar_id = " + NextNumeric.ToInt(Request.QueryString["MARID"]) +
                                                        " AND sc_in_garanzia = 1", " sc_data_ricevimento DESC");
                RichiesteElenco.Columns.RemoveAt(4);
            }
            else if (StnTipoUtente.Equals("supervisore"))
                dt = InfoschedeTools.GetSchedeDataTable(0, 0, Session["stato"].ToString(),
                                                        Session["filtroRichieste"].ToString() +
                                                        " AND riv_azienda_capogruppo_id = " + NextMembershipRivenditore.Current.Id,
                                                        " sc_data_ricevimento DESC");
            else
            {
                dt = InfoschedeTools.GetSchedeDataTable(0, cliente.Id, Session["stato"].ToString(),
                                                        Session["filtroRichieste"].ToString(), " sc_data_ricevimento DESC");
                RichiesteElenco.Columns.RemoveAt(4);
            }

            if (dt.Rows.Count > 0)
            {
                Summary.InnerHtml = "Trovate n° " + dt.Rows.Count + " richieste in " +
                                    (dt.Rows.Count / RichiesteElenco.PageSize + 1) + " pagine";
                RichiesteElenco.DataSource = dt;
                RichiesteElenco.DataBind();
            }
        }
        else
            Response.Redirect(NextPage.UrlHomePage);

        NextControlsTools.SetCssClass(this.Layer, "elencorichieste");
    }
    
    protected void RichiesteElenco_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
        {
            LinkButton statoHd = (LinkButton)e.Row.FindControl("statoHd");
            LinkButton numeroHd = (LinkButton)e.Row.FindControl("numeroHd");
            LinkButton dataRicevimentoHd = (LinkButton)e.Row.FindControl("dataRicevimentoHd");
            LinkButton modelloHd = (LinkButton)e.Row.FindControl("modelloHd");
            LinkButton riferimentoClienteHd = (LinkButton)e.Row.FindControl("riferimentoClienteHd");
            LinkButton nomeRivenditoreHd = (LinkButton)e.Row.FindControl("nomeRivenditoreHd");
            LinkButton numeroCaricoHd = (LinkButton)e.Row.FindControl("numeroCaricoHd");
            LinkButton dataCaricoHd = (LinkButton)e.Row.FindControl("dataCaricoHd");
            HtmlGenericControl linkHd = (HtmlGenericControl)e.Row.FindControl("linkHd");

            #region gestione link ordinamento colonne
            if (Session["sortDirection_stato"] == null)
                Session["sortDirection_stato"] = SortDirection.Ascending;
            if (Session["sortDirection_sc_numero"] == null)
                Session["sortDirection_sc_numero"] = SortDirection.Ascending;
            if (Session["sortDirection_sc_data_ricevimento"] == null)
                Session["sortDirection_sc_data_ricevimento"] = SortDirection.Descending;
            if (Session["sortDirection_modello"] == null)
                Session["sortDirection_modello"] = SortDirection.Ascending;
            if (Session["sortDirection_sc_rif_cliente"] == null)
                Session["sortDirection_sc_rif_cliente"] = SortDirection.Ascending;
            if (Session["sortDirection_sc_numero_DDT_di_carico"] == null)
                Session["sortDirection_sc_numero_DDT_di_carico"] = SortDirection.Ascending;
            if (Session["sortDirection_sc_data_DDT_di_carico"] == null)
                Session["sortDirection_sc_data_DDT_di_carico"] = SortDirection.Descending;
            
            statoHd.Text = "Stato richiesta";
            statoHd.Attributes.Add("sortExpression", "stato");
            statoHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "stato")
                NextControlsTools.SetCssClass(statoHd.Parent, ((SortDirection)Session["sortDirection_stato"] == SortDirection.Descending ? "down" : "up"));
            numeroHd.Text = "N.";
            numeroHd.Attributes.Add("sortExpression", "sc_numero");
            numeroHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "sc_numero")
                NextControlsTools.SetCssClass(numeroHd.Parent, ((SortDirection)Session["sortDirection_sc_numero"] == SortDirection.Descending ? "down" : "up"));
            dataRicevimentoHd.Text = "Data richiesta";
            dataRicevimentoHd.Attributes.Add("sortExpression", "sc_data_ricevimento");
            dataRicevimentoHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "sc_data_ricevimento")
                NextControlsTools.SetCssClass(dataRicevimentoHd.Parent, ((SortDirection)Session["sortDirection_sc_data_ricevimento"] == SortDirection.Descending ? "down" : "up"));
            modelloHd.Text = "Modello";
            modelloHd.Attributes.Add("sortExpression", "modello");
            modelloHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "modello")
                NextControlsTools.SetCssClass(modelloHd.Parent, ((SortDirection)Session["sortDirection_modello"] == SortDirection.Descending ? "up" : "down"));
            riferimentoClienteHd.Text = "Riferimento cliente";
            riferimentoClienteHd.Attributes.Add("sortExpression", "sc_rif_cliente");
            riferimentoClienteHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "sc_rif_cliente")
                NextControlsTools.SetCssClass(riferimentoClienteHd.Parent, ((SortDirection)Session["sortDirection_sc_rif_cliente"] == SortDirection.Descending ? "down" : "up"));
            numeroCaricoHd.Text = "N. DDT";
            numeroCaricoHd.Attributes.Add("sortExpression", "sc_numero_DDT_di_carico");
            numeroCaricoHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "sc_numero_DDT_di_carico")
                NextControlsTools.SetCssClass(numeroCaricoHd.Parent, ((SortDirection)Session["sortDirection_sc_numero_DDT_di_carico"] == SortDirection.Descending ? "down" : "up"));
            dataCaricoHd.Text = "Data DDT";
            dataCaricoHd.Attributes.Add("sortExpression", "sc_data_DDT_di_carico");
            dataCaricoHd.Click += new EventHandler(Sort_Click);
            if (Session["sortExpression"].ToString() == "sc_data_DDT_di_carico")
                NextControlsTools.SetCssClass(dataCaricoHd.Parent, ((SortDirection)Session["sortDirection_sc_data_DDT_di_carico"] == SortDirection.Descending ? "down" : "up"));

            if (StnTipoUtente.Equals("supervisore"))
            {
                if (Session["sortDirection_nome_rivenditore"] == null)
                    Session["sortDirection_nome_rivenditore"] = SortDirection.Ascending;
                nomeRivenditoreHd.Text = "Nome Rivenditore";
                nomeRivenditoreHd.Attributes.Add("sortExpression", "nome_rivenditore");
                nomeRivenditoreHd.Click += new EventHandler(Sort_Click);
                if (Session["sortExpression"].ToString() == "nome_rivenditore")
                    NextControlsTools.SetCssClass(nomeRivenditoreHd.Parent, ((SortDirection)Session["sortDirection_nome_rivenditore"] == SortDirection.Descending ? "down" : "up"));
            }
            #endregion

            linkHd.InnerText = "Dettagli";
        }
        else if (e.Row.RowType == DataControlRowType.DataRow)
        {
            HtmlGenericControl stato = (HtmlGenericControl)e.Row.FindControl("stato");
            HtmlGenericControl numero = (HtmlGenericControl)e.Row.FindControl("numero");
            HtmlGenericControl dataRicevimento = (HtmlGenericControl)e.Row.FindControl("dataRicevimento");
            HtmlGenericControl modello = (HtmlGenericControl)e.Row.FindControl("modello");
            HtmlGenericControl riferimentoCliente = (HtmlGenericControl)e.Row.FindControl("riferimentoCliente");            
            HtmlGenericControl numeroCarico = (HtmlGenericControl)e.Row.FindControl("numeroCarico");
            HtmlGenericControl dataCarico = (HtmlGenericControl)e.Row.FindControl("dataCarico");
            HtmlAnchor link = (HtmlAnchor)e.Row.FindControl("link");

            DataRow dr = ((DataRowView)e.Row.DataItem).Row;
            stato.InnerText = dr["stato"].ToString();
            numero.InnerText = dr["sc_numero"].ToString();
            dataRicevimento.InnerText = DateTime.Parse(dr["sc_data_ricevimento"].ToString()).ToString(NextDateTime.StringFormats.DateIta);
            modello.InnerText = dr["modello"].ToString();
            riferimentoCliente.InnerText = dr["sc_rif_cliente"].ToString();            
            numeroCarico.InnerText = dr["sc_numero_DDT_di_carico"].ToString();
            DateTime dataDdt;
            if (DateTime.TryParse(dr["sc_data_DDT_di_carico"].ToString(), out dataDdt))
                dataCarico.InnerText = dataDdt.ToString(NextDateTime.StringFormats.DateIta);
            link.InnerText = "Apri";
            string query_string = "SCHEDAID=" + dr["sc_id"].ToString();
            if (!string.IsNullOrEmpty(Request.QueryString["MARID"]))
                query_string += "&MARID=" + Request.QueryString["MARID"].ToString();
            link.HRef = NextPage.GetPageSiteUrl(StnPaginaSchedaId, query_string);
            if (StnTipoUtente.Equals("supervisore"))
            {
                HtmlGenericControl nomeRivenditore = (HtmlGenericControl)e.Row.FindControl("nomeRivenditore");
                nomeRivenditore.InnerText = dr["nome_rivenditore"].ToString();
            }
        }
    }
    
    void Sort_Click(object sender, EventArgs e)
    {
        string sortExpression = ((LinkButton)sender).Attributes["sortExpression"].ToString();
        Session["sortExpression"] = sortExpression;
        if ((SortDirection)Session["sortDirection_" + sortExpression] == SortDirection.Descending)
        {
            Session["sortDirection_" + sortExpression] = SortDirection.Ascending;
            RichiesteElenco.Sort(sortExpression, SortDirection.Ascending);
        }
        else
        {
            Session["sortDirection_" + sortExpression] = SortDirection.Descending;
            RichiesteElenco.Sort(sortExpression, SortDirection.Descending);
        }
    }

    protected void RichiesteElenco_Sorting(object sender, GridViewSortEventArgs e)
    {
        GridView gv = (GridView)sender;
        string sortDirection = "";
        DataTable dt;

        if (e.SortDirection == SortDirection.Descending)
            sortDirection = " DESC";

        if (StnTipoUtente.Equals("costruttore"))
            dt = InfoschedeTools.GetSchedeDataTable(0, 0, Session["stato"].ToString(),
                                                    Session["filtroRichieste"].ToString() +
                                                    " AND mar_id = " + NextNumeric.ToInt(Request.QueryString["MARID"]),
                                                    e.SortExpression + sortDirection);
        else if (StnTipoUtente.Equals("supervisore"))
            dt = InfoschedeTools.GetSchedeDataTable(0, 0,Session["stato"].ToString(),
                                                    Session["filtroRichieste"].ToString() +
                                                    " AND riv_azienda_capogruppo_id = " + NextMembershipRivenditore.Current.Id,
                                                    e.SortExpression + sortDirection);
        else
            dt = InfoschedeTools.GetSchedeDataTable(0, cliente.Id, Session["stato"].ToString(),
                                                    Session["filtroRichieste"].ToString(), e.SortExpression + sortDirection);
        if (dt.Rows.Count > 0)
        {
            gv.DataSource = dt;
            gv.DataBind();
        }
    }

    protected void GridView_PageChanging(object sender, GridViewPageEventArgs e)
    {
        ((GridView)sender).PageIndex = e.NewPageIndex;
        ((GridView)sender).DataBind();
    }
}