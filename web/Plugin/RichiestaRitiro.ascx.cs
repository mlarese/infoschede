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

public partial class Plugin_RichiestaRitiro : NextFramework.NextControls.NextUserControl
{
    /// <summary>
    /// dati utente autenticato
    /// </summary>
    protected NextMembershipRivenditore cliente = NextMembershipRivenditore.CurrentCliente;

    #region Settings...

    #endregion

    protected override void OnLoad(EventArgs e)
    {
        #region Settings init...

        #endregion

        base.OnLoad(e);
        //NextControlsTools.SetCssClass(this.Layer, "scheda");

        int ddtId = 0, clienteId = 0, trasportatoreId = 0;

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
                    cliente.Contatto.SetRecapiti();

                    TrasportatoreDiv.InnerHtml = "<p class=\"titolo\"><span>spett.le " + trasportatore.Contatto.GetName().ToUpper() + "</span>" +
                                                 "<span>Richiesta ritiro</span></p>";
                    DescrizioneDiv.InnerHtml = "<p>Vogliate eseguire i sottoindicati ritiri e darcene debito sul ns contratto</p>";
                    MittenteDiv.InnerHtml = "<p><span>Hidroservices s.a.s. di Manente Daniele & C</span>";
                                            //"<span>cod. 0592894</span></p>";
                    DestinatarioDiv.InnerHtml = "<p>ritiro presso:</p>" +
                                                "<p>" + cliente.Contatto.GetName() + "</p>" +
                                                "<p><span class=\"label\">via</span><span>" + cliente.Contatto.Indirizzo + "</span></p>" +
                                                "<p><span class=\"label\">città</span><span>" + cliente.Contatto.Citta + "</span></p>" +
                                                "<p><span class=\"label\">provincia</span><span>" + cliente.Contatto.Provincia + "</span></p>" +
                                                "<p>riferimenti:</p>" +
                                                "<p class=\"riferimenti\"><label>telefono</label><span>" + cliente.Contatto.Telefono + "</span></p>" +
                                                "<p class=\"riferimenti\"><label>fax</span><label>" + cliente.Contatto.Fax + "</span></p>";
                    DataRitiroDiv.InnerHtml = "<p>DATA RITIRO RICHIESTA</p>";
                    DatiDiv.InnerHtml = "<p><span class=\"mini\">ritiro num.</span>" +
                                        "<span class=\"mini numero\">" + dt.Rows[0]["ddt_numero"].ToString() + "</span>" +
                                        "<span class=\"mini\">data</span>" +
                                        "<span class=\"mini data\">" + DateTime.Parse(dt.Rows[0]["ddt_data"].ToString()).ToString(NextDateTime.StringFormats.DateIta) + "</span></p>" +
                                        "<p><span class=\"label\">descrizione</span></p>" +
                                        "<span>" + dt.Rows[0]["ddt_note"].ToString() + "</span>" +
                                        "<p><span class=\"label\">num. colli</span>" +
                                        "<span>" + dt.Rows[0]["ddt_numero_colli"].ToString() + "</span></p>" +
                                        "<p><span class=\"label\">peso kg</span>" +
                                        "<span>" + dt.Rows[0]["ddt_peso"].ToString() + "</span></p>" +
                                        "<p><span class=\"label\">volume mq</span>" +
                                        "<span>" + dt.Rows[0]["ddt_volume"].ToString() + "</span></p>";
                    ConsegnaDiv.InnerHtml = "<p>CONSEGNA PRESSO NS SEDE</p>" +
                                            "<p>" + trasportatore.Riv_codice + "</span>";
                }
                else
                    this.Layer.Visible = false;
            }
            else
                this.Layer.Visible = false;
        }
    }
}
