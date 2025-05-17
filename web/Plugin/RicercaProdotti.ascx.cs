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

public partial class Plugin_RicercaProdotti : NextFramework.NextControls.NextUserControl
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

        Titolo.InnerHtml = NextString.HtmlEncode("Ricerca per\nnome / codice");
        Cerca.Text = "Cerca";
        ViewAll.Text = "Vedi tutte";
        
        // imposta eventualmente valore da ricerca precedente
        if (Session["prodotto"] != null && Session["prodotto"].ToString() != "")
            ProdottoInput.Value = Session["prodotto"].ToString();

        NextControlsTools.SetCssClass(this.Layer, "ricercarichieste");
    }

    protected void Cerca_Click(object sender, EventArgs e)
    {
        string filtro = "";
        if (((Button)sender).ID == "Cerca")
        {
            // validazione
            if (ProdottoInputValid.IsValid)
            {
                if (Request.Form[ProdottoInput.UniqueID] != null && Request.Form[ProdottoInput.UniqueID].ToString() != "")
                {
                    List<String> listaCampi = new List<String>(new String[] {"art_nome_it", "art_nome_en", "art_cod_int",
                                                                             "art_cod_pro", "art_cod_alt"});
                    filtro += " AND " + NextSql.TextSearch(Request.Form[ProdottoInput.UniqueID].ToString(),
                                                           listaCampi, true);
                    Session["prodotto"] = Request.Form[ProdottoInput.UniqueID].ToString();
                }
                else
                    Session["prodotto"] = "";
                Session["filtroProdotti"] = filtro;

                Response.Redirect(NextPage.Url);
            }
            else
            {
                ErroriListaDiv.Visible = true;
                ErroriListaDiv.InnerHtml += NextString.HtmlEncode("\n" + ProdottoInputValid.ErrorMessage + "\n");
            }
        }

        else
        {
            Session["prodotto"] = "";
            Session["filtroProdotti"] = "";
            Response.Redirect(NextPage.Url);
        }
    }
}
