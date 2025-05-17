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

public partial class Plugin_LetteraVettura : NextFramework.NextControls.NextUserControl
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

            if (int.TryParse(dt.Rows[0]["ddt_cliente_id"].ToString(), out clienteId) && clienteId > 0 &&
                int.TryParse(dt.Rows[0]["ddt_trasportatore_id"].ToString(), out trasportatoreId) && trasportatoreId > 0)
            {
                NextMembershipRivenditore trasportatore = new NextMembershipRivenditore();
                if (trasportatore.SetPropertiesById(trasportatoreId) && cliente.SetPropertiesById(clienteId))
                {
                    cliente.Contatto.SetRecapiti();
                    MittenteDiv.InnerHtml = "<p>MITTENTE : </p>" +      //tolto: 0592894
                                            "<p>HIDROSERVICES s.r.l.</p>" +
                                            "<p>VIA VITTORIO VENETO 2/A</p>" +
                                            "<p>30030 SALZANO VE</p>" +
                                            "<p>TEL. E FAX 041484691</p>";
                    DestinatarioDiv.InnerHtml = "<h2>DESTINATARIO</h2>" +
                                                "<p>" + cliente.Contatto.GetName() + "</p>" +
                                                "<p><span>Via</span><span>" + cliente.Contatto.Indirizzo + "</span></p>" +
                                                "<p><span>Località</span><span>" + cliente.Contatto.Stato + "</span></p>" +
                                                "<p><span>Cap</span><span>" + cliente.Contatto.Cap + "</span>" +
                                                "<span>Città</span><span>" + cliente.Contatto.Citta + "</span></p>" +
                                                "<p><span>Provincia</span><span>" + cliente.Contatto.Provincia + "</span>" +
                                                "<span>Telefono</span><span>" + cliente.Contatto.Telefono + "</span></p>";
                    FirmaMittenteDiv.InnerHtml = "<h2>FIRMA MITTENTE</h2>";
                    TrasportatoreDiv.InnerHtml = "<p>" + trasportatore.Contatto.GetName().ToUpper() + "</p>" +
                                                 "<p>" + trasportatore.Contatto.BLL.GetAddress().ToUpper() + "</p>" +
                                                 "<p>FILIALE : " + trasportatore.Contatto.Qualifica + "</p>";
                    DatiDiv.InnerHtml = "<p><span>NUMERO</span>" +
                                        "<input readonly value=\"" + dt.Rows[0]["ddt_numero"].ToString() + "\" style=\"width:41px;text-align:right;\"/>" +
                                        "<span style=\"margin-left:2px;\">DATA</span>" +
                                        "<input readonly value=\"" +
                                        DateTime.Parse(dt.Rows[0]["ddt_data"].ToString()).ToString(NextDateTime.StringFormats.DateIta) +
                                        "\" style=\"width:140px;text-align:right;\"/></p>" +
                                        "<p><span>RESA</span><input readonly value=\"" + dt.Rows[0]["por_titolo_it"].ToString() + "\" style=\"width:270px;\"/></p>" +
                                        "<p><span>CONTRASSEGNO</span><input readonly value=\"" + dt.Rows[0]["ddt_contrassegno"].ToString() + "\" style=\"width:190px;\"/>" +
                                        "<p><span>DOC.</span><input readonly value=\"DDT\" style=\"width:142px;\"/>" +
                                        "<span>NUM.</span><input readonly value=\"" + dt.Rows[0]["ddt_numero"].ToString() + "\" style=\"width:72px;\"/></p>" +
                                        "<p><span style=\"width:97px;\">N° COLLI</span>" +
                                        "<span style=\"width:97px;\">PESO</span>" +
                                        "<span style=\"width:105px;\">VOLUME M3</span></p>" +
                                        "<p><input readonly value=\"" + dt.Rows[0]["ddt_numero_colli"].ToString() + "\" style=\"width:94px;\"/>" +
                                        "<input readonly value=\"" + dt.Rows[0]["ddt_peso"].ToString() + "\" style=\"width:94px;\"/>" +
                                        "<input readonly value=\"" + dt.Rows[0]["ddt_volume"].ToString() + "\" style=\"width:105px;\"/></p>" +
                                        "<p><span>DESCRIZIONE</span></p>" +
                                        "<p><input readonly value=\"" + dt.Rows[0]["ddt_note"].ToString() + "\" style=\"width:320px;\"/></p>";
                    FirmaRitiroDiv.InnerHtml = "<h2>FIRMA RITIRO</h2>";
                }
                else
                    this.Layer.Visible = false;
            }
            else
                this.Layer.Visible = false;
        }
    }
}
