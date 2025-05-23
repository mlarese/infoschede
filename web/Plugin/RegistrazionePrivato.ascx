<%@ Control Language="C#" AutoEventWireup="true" CodeFile="RegistrazionePrivato.ascx.cs" Inherits="Plugin_RegistrazionePrivato" %>
<%@ Register Assembly="NextFramework" Namespace="NextFramework.NextControls" TagPrefix="nextFramework" %>

<h1 id="Titolo" class="titolo_sez" runat="server"></h1>

<div class="background">
    <div class="container contattaci">

        <nextFramework:Contattaci ID="Contattaci" runat="server"
                                    OnDataBound="Contattaci_DataBound" 
                                    OnOdsInserting="CheckUser" OnOdsUpdating="CheckUser"
                                    OnFormSave="Contattaci_FormSave"
                                    EnableViewState="true" ValidationGroup="Registrazione">
            
            <InsertItemTemplate>
            
                <%--<asp:ValidationSummary ID="ValidationSummary1" runat="server" DisplayMode="List" CssClass="form_mandatory" ValidationGroup="Registrazione" />
                <asp:HiddenField ID="Id" runat="server" Value='<%# Bind("Id") %>'/>
                <asp:HiddenField ID="UtenteId" runat="server" Value='<%# Bind("UtenteId") %>'/>--%>
                <asp:TextBox ID="IsSocieta" Visible="false" runat="server" Text='<%# Bind("IsSocieta") %>' CssClass="is_soc">false</asp:TextBox>
                
                <%--<table cellpadding="0" cellspacing="0" width="100%" class="form">--%>

                    <tr>
                        <td class="form_label first_td"></td>
                        <td class="form_input first_td" colspan="3"></td>
                    </tr>
                                    
                    <tr>
                        <td class="sezione form_title" colspan="4">
                            <label id="sezione0" class="sezione">I tuoi dati</label>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label12" class="lb_edit campi_sx">Nome</label>
                        </td>
                        <td class="form_input" colspan="3">
                            <asp:TextBox ID="Nome" runat="server" Text='<%# Bind("Nome") %>' CssClass="campi_dx nome"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" ControlToValidate="Nome"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Nome' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label13" class="lb_edit campi_sx">Cognome</label>
                        </td>
                        <td class="form_input" colspan="3">
                            <asp:TextBox ID="Cognome" runat="server" Text='<%# Bind("Cognome") %>' CssClass="campi_dx cognome"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ControlToValidate="Cognome"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Cognome' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                    </tr>
                    
                    <tr id="trCF" runat="server" visible="false">
                        <td class="form_label">
                            <label id="Label6" class="lb_edit campi_sx">Codice Fiscale</label>
                        </td>
                        <td class="form_input" colspan="3">
                            <asp:TextBox ID="CF" runat="server" Text='<%# Bind("CF") %>' CssClass="campi_dx codice_fiscale"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="CF"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Codice Fiscale' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                    </tr>
                    <tr>
                        <td class="form_label">
                            <label id="Label2" class="campi_sx">Indirizzo</label>
                        </td>
                        <td class="form_input1">
                            <asp:TextBox ID="Indirizzo" runat="server" Text='<%# Bind("Indirizzo") %>' CssClass="campi_dx indirizzo"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="Indirizzo"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Indirizzo' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                        
                        <td class="form_label2">
                            <label id="Label7" class="campi_sx">Cap</label>
                        </td>
                        <td class="form_input2">
                            <asp:TextBox ID="Cap" runat="server" Text='<%# Bind("Cap") %>' CssClass="campi_dx cap"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ControlToValidate="Cap"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Cap' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label3" class="lb_edit campi_sx">Citt�</label>
                        </td>
                        <td class="form_input1">
                            <asp:TextBox ID="Citta" runat="server" Text='<%# Bind("Citta") %>' CssClass="campi_dx citta"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator6" runat="server" ControlToValidate="Citta"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Citt�' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                        
                        <td class="form_label2">
                            <label id="Label8" class="lb_edit campi_sx">Provincia</label>
                        </td>
                        <td class="form_input2">
                            <asp:TextBox ID="Provincia" runat="server" Text='<%# Bind("Provincia") %>' CssClass="campi_dx provincia"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ControlToValidate="Provincia"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Provincia' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                        </td>
                    </tr>
                    
                    <%--<tr>
                        <td class="form_label1">
                            <label id="Label14" class="campi_sx">Localit�:</label>
                        </td>
                        <td class="form_input2">
                            <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("Localita") %>' CssClass="campi_dx localita"></asp:TextBox>
                        </td>
                    </tr>--%>

                    <tr>
                        <td class="form_label">
                            <label id="Label9" class="campi_sx">Telefono</label>
                        </td>
                        <td class="form_input1" colspan="3">
                            <asp:TextBox ID="Telefono" runat="server" Text='<%# Bind("Telefono") %>' CssClass="campi_dx telefono"></asp:TextBox>
                        </td>
                    </tr>
                    
                    <tr id="trFax" runat="server" visible="false">
                        <td class="form_label">
                            <label id="Label10" class="lb_edit campi_ds">Fax</label>
                        </td>
                        <td class="form_input1" colspan="3">
                            <asp:TextBox ID="Fax" runat="server" Text='<%# Bind("Fax") %>' CssClass="campi_dx fax"></asp:TextBox>
                        </td>
                    </tr>
                    
                    <tr class="pre_title">
                        <td class="form_label">
                            <label id="Label11" class="campi_sx">E-mail</label>
                        </td>
                        <td class="form_input" colspan="3">
                            <asp:TextBox ID="Email" runat="server" Text='<%# Bind("Email") %>' CssClass="campi_dx email"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ControlToValidate="Email"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Email' � obbligatorio." Text="&lt;--"
                                ValidationGroup="Registrazione" Display="Dynamic" />
                            <asp:RegularExpressionValidator CssClass="form_mandatory form_mandatory_error" ID="RegularExpressionValidator1" runat="server"
                                ControlToValidate="Email" Text="&lt;--" ValidationExpression="<%# NextFramework.Messaggi.NextEmail.REAddress %>"
                                ErrorMessage='<%#NextFramework.NextCom.BLLContatto.ExcEmailErrata %>' ValidationGroup="Registrazione" />
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="sezione form_title" colspan="4">
                            <label id="sezione1" class="sezione">Dati di accesso</label>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label sfondo">
                            <label id="Label18" class="lb_login campi_sx">Login</label>
                        </td>
                        <td class="form_input sfondo" colspan="3">
                            <asp:TextBox ID="Login" runat="server" CssClass="campi_dx login"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="Login"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Login' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                            <asp:RegularExpressionValidator CssClass="form_mandatory form_mandatory_error" ID="RegularExpressionValidatorLogin" runat="server"
                                ControlToValidate="Login" Text="&lt;--" ValidationExpression="<%# NextFramework.NextPassport.BLLUser.RELogin %>"
                                ErrorMessage='<%#NextFramework.NextPassport.BLLUser.ExcLoginNonValido %>' ValidationGroup="Registrazione" />
                        </td>
                    </tr>

                    <tr>
                        <td class="form_label sfondo">
                            <label id="Label20" class="lb_new_pwd campi_sx">Password</label>
                        </td>
                        <td class="form_input sfondo" colspan="3">
                            <asp:TextBox ID="Password" TextMode="Password" runat="server" CssClass="campi_dx new_pwd"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="Password"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Password' � obbligatorio." Text="&lt;--" ValidationGroup="Registrazione" Display="Dynamic" />
                            <asp:RegularExpressionValidator CssClass="form_mandatory form_mandatory_error" ID="RegularExpressionValidatorPassword" runat="server"
                                ControlToValidate="Password" Text="&lt;--" ValidationExpression="<%# NextFramework.NextPassport.BLLUser.REPassword %>"
                                ErrorMessage='<%#NextFramework.NextPassport.BLLUser.ExcPasswordNonValida %>' ValidationGroup="Registrazione" />
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label sfondo ultima">
                            <label id="Label21" class="lb_conf_pwd campi_sx">Conferma password</label>
                        </td>
                        <td class="form_input sfondo ultima" colspan="3">
                            <asp:TextBox ID="ConfermaPassword" TextMode="password" runat="server" CssClass="campi_dx conf_pwd"></asp:TextBox>
                            <span class="form_mandatory">(*)</span>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" ControlToValidate="ConfermaPassword"
                                CssClass="form_mandatory form_mandatory_error" ErrorMessage="Il campo 'Conferma password' � obbligatorio." Text="&lt;--"
                                ValidationGroup="Registrazione" Display="Dynamic" />
                            <asp:RegularExpressionValidator CssClass="form_mandatory form_mandatory_error" ID="RegularExpressionValidatorConfermaPassword" runat="server"
                                ControlToValidate="ConfermaPassword" Text="&lt;--" ValidationExpression="<%# NextFramework.NextPassport.BLLUser.REPassword %>"
                                ErrorMessage='<%#NextFramework.NextPassport.BLLUser.ExcPasswordNonValida %>' ValidationGroup="Registrazione" />
                            <asp:CompareValidator CssClass="form_mandatory form_mandatory_error" ID="CompareValidatorConfermaPassword" runat="server" ValidationGroup="Registrazione"
                                ControlToValidate="ConfermaPassword" ControlToCompare="Password" Text="<--" Display="Dynamic"
                                ErrorMessage='<%#NextFramework.NextPassport.BLLUser.ExcPasswordNotRetypedCorrectly %>'/>
                        </td>
                    </tr>
                    
                    <tr id="trMessaggio" runat="server" visible="false">
                        <td class="form_label">
                            <label id="Label4" class="lb_edit campi_sx">Messaggio</label>
                        </td>
                        <td class="form_input" colspan="3">
                            <asp:TextBox ID="Note" TextMode="MultiLine" runat="server" Text='<%# Bind("Note") %>' CssClass="campi_dx note"></asp:TextBox>
                        </td>
                    </tr>

                <%--</table>--%> 
                
            </InsertItemTemplate>
            
            <ItemTemplate>
            
                <%--<table cellpadding="0" cellspacing="0" width="100%" class="form">--%>
                    
                    
                    <tr>
                        <td class="form_label first_td"></td>
                        <td class="form_input first_td" colspan="3"></td>
                    </tr>
                
                    <tr>
                        <td class="form_label">
                            <label id="Label16" class="campi_sx">Nome</label>
                        </td>
                        <td class="form_value" colspan="3">
                            <asp:Label ID="Nome" runat="server" Text='<%# Bind("Nome") %>' CssClass="campi_dx nome"></asp:Label>
                        </td>
                    </tr>

                    <tr>
                        <td class="form_label">
                            <label id="Label12" class="campi_sx">Cognome</label>
                        </td>
                        <td class="form_value" colspan="3">
                            <asp:Label ID="Cognome" runat="server" Text='<%# Bind("Cognome") %>' CssClass="campi_dx cognome"></asp:Label>
                        </td>
                    </tr>

                    <tr>
                        <td class="form_label">
                            <label id="Label6" class="campi_sx">Codice Fiscale</label>
                        </td>
                        <td class="form_value" colspan="3">
                            <asp:Label ID="CF" runat="server" Text='<%# Bind("CF") %>' CssClass="campi_dx codice_fiscale"></asp:Label>
                        </td>
                    </tr>
                    
                    <%--<tr>
                        <td class="form_label">
                            <label id="Label5" class="campi_sx">Attivit�:</label>
                        </td>
                        <td class="form_value">
                            <asp:Label ID="Label17" runat="server" Text='<%# Bind("Attivita") %>' CssClass="campi_dx localita"></asp:Label>
                        </td>
                    </tr>--%>

                    <tr>
                        <td class="form_label">
                            <label id="Label2" class="campi_sx">Indirizzo</label>
                        </td>
                        <td class="form_value1">
                            <asp:Label ID="Indirizzo" runat="server" Text='<%# Bind("Indirizzo") %>' CssClass="campi_dx indirizzo"></asp:Label>
                        </td>
                   
                        <td class="form_label2">
                            <label id="Label7" class="campi_sx">Cap</label>
                        </td>
                        <td class="form_value2">
                            <asp:Label ID="Cap" runat="server" Text='<%# Bind("Cap") %>' CssClass="campi_dx cap"></asp:Label>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label1" class="campi_sx">Citt�</label>
                        </td>
                        <td class="form_value1">
                            <asp:Label ID="Citta" runat="server" Text='<%# Bind("Citta") %>' CssClass="campi_dx citta"></asp:Label>
                        </td>
                    
                        <td class="form_label2">
                            <label id="Label8" class="campi_sx">Provincia</label>
                        </td>
                        <td class="form_value2">
                            <asp:Label ID="Provincia" runat="server" Text='<%# Bind("Provincia") %>' CssClass="campi_dx provincia"></asp:Label>
                        </td>
                    </tr>
                    
                    
                    <%--<tr>
                        <td class="form_label1">
                            <label id="Label3" class="campi_sx">Localit�</label>
                        </td>
                        <td class="form_input2">
                            <asp:Label ID="Localita" runat="server" Text='<%# Bind("Localita") %>' CssClass="campi_dx localita"></asp:Label>
                        </td>
                    </tr>--%>
                   
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label9" class="campi_sx">Telefono</label>
                        </td>
                        <td class="form_value" colspan="3">
                            <asp:Label ID="Telefono" runat="server" Text='<%# Bind("Telefono") %>' CssClass="campi_dx telefono"></asp:Label>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label10" class="campi_sx">Fax</label>
                        </td>
                        <td class="form_value" colspan="3">
                            <asp:Label ID="Fax" runat="server" Text='<%# Bind("Fax") %>' CssClass="campi_dx fax"></asp:Label>
                        </td>
                    </tr>
                    
                    <tr class="pre_title">
                        <td class="form_label">
                            <label id="Label11" class="campi_sx">E-mail</label>
                        </td>
                        <td class="form_value" colspan="3">
                            <%-- mailto --%>
                            <asp:HyperLink ID="Email" runat="server" NavigateUrl='<%# Bind("Email", "mailto:{0}") %>' Text='<%# Bind("Email") %>' CssClass="campi_dx email"></asp:HyperLink>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="sezione form_title" colspan="4">
                            <label id="sezione1" class="sezione">Sezione di accesso</label>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label sfondo">
                            <label id="Label19" class="campi_sx">Login</label>
                        </td>
                        <td class="form_value sfondo" colspan="3">
                            <span id="Login" runat="server" class="campi_dx login"></span>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label sfondo">
                            <label id="Label22" class="campi_sx">Password</label>
                        </td>
                        <td class="form_value sfondo" colspan="3">
                            <span id="Password" runat="server" class="campi_dx password"></span>
                        </td>
                    </tr>
                    
                    <tr>
                        <td class="form_label">
                            <label id="Label17" class="campi_sx">Note</label>
                        </td>
                        <td class="form_input1" colspan="3">
                            <asp:Label ID="Note" runat="server" Text='<%# Bind("Note") %>' CssClass="campi_dx note"></asp:Label>
                        </td>
                    </tr>

                <%--</table>--%>                               

            </ItemTemplate>
            
        </nextFramework:Contattaci>
        
    </div>
</div>