<%@ Control Language="C#" AutoEventWireup="true" CodeFile="Login.ascx.cs" Inherits="Plugin_Login" %>
    <div class="intestazione" runat="server" id="Intestazione"></div>
     <p id="Test" runat="server"></p>
    <asp:Login ID="Login" runat="server" DisplayRememberMe="false"
           PasswordRequiredErrorMessage="Manca la password."  UserNameRequiredErrorMessage="Manca l'user name."
           OnLoggedIn="Login_LoggedIn" OnPreRender="Login_PreRender">
        <LayoutTemplate>
             <div id="Container" class="up_container" runat="server">               
                <div class="User">
                    <span class="label">
                        User Name: 
                    </span>
                    <asp:TextBox ID="UserName" runat="server"></asp:TextBox><asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="UserName" ValidationGroup="<%# this.Parent.ID %>">*</asp:RequiredFieldValidator></div>
                <div class="Password">
                    <span class="label">
                        Password:
                    </span>
                    <asp:TextBox ID="Password" runat="server" TextMode="Password"></asp:TextBox><asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="Password" ValidationGroup="<%# this.Parent.ID %>">*</asp:RequiredFieldValidator></div>
                <div class="Submit">
                    <asp:Button ID="LoginButton" runat="server" CommandName="Login" Text="Entra" value="ACCEDI" name="EXECUTE" ValidationGroup="ctl00<%# this.Parent.ID %>Login1" /></div>
            </div>
        </LayoutTemplate>
    </asp:Login>