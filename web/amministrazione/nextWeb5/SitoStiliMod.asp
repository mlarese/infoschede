<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="intestazione.asp" -->
<!--#INCLUDE FILE="../library/Tools4Color.asp" -->
<%
'check dei permessi dell'utente
CALL CheckAutentication(index.ChkPrm(prm_stili_accesso, 0))

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	Server.Execute("SitoStiliSalva.asp")
end if

dim dicitura
set dicitura = New testata 
dicitura.iniz_sottosez(0)
dicitura.sezione = "Gestione siti - stili di testo - modifica"
dicitura.puls_new = "INDIETRO"
dicitura.link_new = "SitoStili.asp"
dicitura.scrivi_con_sottosez() 

dim conn, rs, sql, cssO
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString")
set rs = Server.CreateObject("ADODB.RecordSet")

if request("goto")<>"" then
	CALL GotoRecord(conn, rs, session("STILI_SQL"), "style_id", "SitoStiliMod.asp")
end if

'genera stili una volta sola anche per tutti gli iframe che mostrano gli stili.
'La variabile di sessione eviene rimossa nell'esecuzione dell'iframe dove c'e' il parametro LAST=1
set cssO = new CssManager
Session("TMP_STILI_TESTO_" & session("AZ_ID")) = cssO.GenerateCss(conn, session("AZ_ID"), false)

sql = " SELECT * FROM (tb_webs INNER JOIN tb_css_groups ON tb_webs.id_webs = tb_css_groups.grp_id_webs) " + _
	  " INNER JOIN tb_css_styles ON tb_css_groups.grp_id = tb_css_styles.style_grp_id " + _
	  " WHERE style_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
%>

<script language="JavaScript" type="text/javascript">
	function Preview_UpdateStyleProperty(PropertyName, PropertyValue){
		//recupera elemento di esempio e anteprima
		//var preview_element = window.frames.preview_frame.document.getElementById("esempio");
		var preview_element = document.getElementById("preview_frame").contentDocument.getElementById("esempio")
		preview_element.style.setAttribute(PropertyName, PropertyValue, 0);
	}
	
	//imposta font family
	function OnChange_FontFamily(obj){
		Preview_UpdateStyleProperty('fontFamily', obj.options[obj.selectedIndex].value);
	}
	
	//imposta font-size
	function OnChange_FontSize(obj){
		Preview_UpdateStyleProperty('fontSize', toNumber(obj.options[obj.selectedIndex].value) + "%");
	}
	
	//impostazione grassetto
	function OnChange_FontWeight(obj){
		Preview_UpdateStyleProperty('fontWeight', obj.options[obj.selectedIndex].value);
	}
	
	//impostazione corsivo/stile
	function OnChange_FontStyle(obj){
		Preview_UpdateStyleProperty('fontStyle', obj.options[obj.selectedIndex].value);
	}
	
	//impostazione spaziatura caratteri
	function OnChange_LetterSpacing(obj){
		Preview_UpdateStyleProperty('letterSpacing', obj.options[obj.selectedIndex].value);
	}
	
	//impostazione allineamento testo
	function OnChange_TextAlign(obj){
		Preview_UpdateStyleProperty('textAlign', obj.options[obj.selectedIndex].value);
	}
	
	//impostazione altezza riga
	function OnChange_LineHeight(obj){
		Preview_UpdateStyleProperty('lineHeight', obj.options[obj.selectedIndex].value + "%");
	}
	
	//impostazione sottolineatura link
	function OnChange_TextDecoration(obj){
		Preview_UpdateStyleProperty('textDecoration', obj.options[obj.selectedIndex].value);
	}
	
	//imposta colore di sfondo
	function OnChange_BackgroundColor(obj){
		Preview_UpdateStyleProperty('backgroundColor', obj.value);
	}
	
	//imposta colore del carattere
	function OnChange_FontColor(obj){
		Preview_UpdateStyleProperty('color', obj.value);
	}
</script>

<div id="content">
	<form action="" method="post" id="form1" name="form1">
	<input type="hidden" name="grp_id_webs" value="<%= rs("grp_id_webs") %>">
	<input type="hidden" name="style_grp_id" value="<%= rs("style_grp_id") %>">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<caption>	
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td class="caption">Modifica stile "<%= rs("style_description") %>" del "<%= rs("grp_name") %>"</td>
					<td align="right" style="font-size: 1px;">
						<a class="button" href="?ID=<%= request("ID") %>&goto=PREVIOUS" title="stile precedente" <%= ACTIVE_STATUS %>>
							&lt;&lt; PRECEDENTE
						</a>
						&nbsp;
						<a class="button" href="?ID=<%= request("ID") %>&goto=NEXT" title="stile successivo" <%= ACTIVE_STATUS %>>
							SUCCESSIVO &gt;&gt;
						</a>
					</td>
				</tr>
			</table>
		</caption>
		<tr><th colspan="4">FORMATTAZIONI DISPONIBILI</th></tr>
		<% if not IsNull(rs("style_font_family")) OR _
			  not IsNull(rs("style_font_size")) OR _
			  not IsNull(rs("style_font_style")) OR _
			  not IsNull(rs("style_font_weight")) OR _
			  not IsNull(rs("style_color")) OR _
			  not IsNull(rs("style_letter_spacing")) then %>
			<tr><th class="L2" colspan="4">carattere</th></tr>
			<% if not IsNull(rs("style_font_family")) then %>
				<tr>
					<td class="label">Tipo di carattere</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.FONT_FAMILY, "tft_style_font_family", rs("style_font_family"), true, "onchange=""OnChange_FontFamily(this);"" style=""width:150px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_font_size")) then %>
				<tr>
					<td class="label">dimensione</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.FONT_SIZE, "tft_style_font_size", rs("style_font_size"), true, "onchange=""OnChange_FontSize(this);"" style=""width:100px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_font_weight")) then %>
				<tr>
					<td class="label">Grassetto</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.FONT_WEIGHT_USER, "tft_style_font_weight", rs("style_font_weight"), true, "onchange=""OnChange_FontWeight(this);"" style=""width:150px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_font_style")) then %>
				<tr>
					<td class="label">Stile</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.FONT_STYLE_USER, "tft_style_font_style", rs("style_font_style"), true, "onchange=""OnChange_FontStyle(this);"" style=""width:150px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_color")) then %>
				<tr>
					<td class="label">Colore</td>
					<td class="content" colspan="3">
						<% CALL WriteColorPicker_Input("form1", "tft_style_color", rs("style_color"), "", false, false, "OnChange_FontColor(sender);") %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_letter_spacing")) then %>
				<tr>
					<td class="label">Spaziatura caratteri</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.LETTER_SPACING_USER, "tft_style_letter_spacing", rs("style_letter_spacing"), true, "onchange=""OnChange_LetterSpacing(this);"" style=""width:100px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
		end if 
		if not IsNull(rs("style_line_height")) OR _
		   not IsNull(rs("style_text_align")) then%>
			<tr><th class="L2" colspan="4">paragrafo</th></tr>
			<% if not IsNull(rs("style_text_align")) then %>
				<tr>
					<td class="label">Allineamento</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.TEXT_ALIGN_USER, "tft_style_text_align", rs("style_text_align"), true, "onchange=""OnChange_TextAlign(this);"" style=""width:100px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_line_height")) then %>
				<!-- <tr>
					<td class="label">Altezza della linea</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.LINE_HEIGHT_USER, "tft_style_line_height", rs("style_line_height"), true, "onchange=""OnChange_LineHeight(this);"" style=""width:100px;""", LINGUA_ITALIANO) %>
					</td>
				</tr> -->
				<input type="hidden" name="tft_style_line_height" value="<%= rs("style_line_height") %>">
			<% end if
		end if
		if not IsNull(rs("style_background_color")) OR _
		   not IsNull(rs("style_text_decoration")) then%>
			<tr><th class="L2" colspan="4">altre formattazioni</th></tr>
			<% if not IsNull(rs("style_background_color")) then %>
				<tr>
					<td class="label">Colore di sfondo</td>
					<td class="content" colspan="3">
						<% CALL WriteColorPicker_Input("form1", "tft_style_background_color", rs("style_background_color"), "", true, false, "OnChange_BackgroundColor(sender);") %>
					</td>
				</tr>
			<% end if
			if not IsNull(rs("style_text_decoration")) then %>
				<tr>
					<td class="label">Sottolineatura</td>
					<td class="content" colspan="3">
						<% CALL DropDownDictionary(cssO.TEXT_DECORATION_USER, "tft_style_text_decoration", rs("style_text_decoration"), true, "onchange=""OnChange_TextDecoration(this);"" style=""width:100px;""", LINGUA_ITALIANO) %>
					</td>
				</tr>
			<% end if
		end if%>
		<tr><th colspan="4">ANTEPRIMA DELLO STILE</th></tr>
		<tr>
			<td class="content_center" colspan="4" style="height:44px; vertical-align:middle;">
				<iframe src="SitoStiliPreview.asp?ID=<%= rs("style_id") %>&last=1" frameborder="0" scrolling="No" 
						id="preview_frame" style="width:99%; height:80px;">
				</iframe>
			</td>
		</tr>
		<% 	'visualizza dati modifica
			CALL Form_DatiModifica(conn, rs, "style_") %>
		<tr>
			<td class="footer" colspan="4">
				(*) Campi obbligatori.
				<input type="submit" class="button" name="mod" value="SALVA">
			</td>
		</tr>
	</table>
	</form>
</div>
</body>
</html>
<%
rs.close
set rs = nothing
conn.Close
set conn = nothing
set cssO = nothing
%>