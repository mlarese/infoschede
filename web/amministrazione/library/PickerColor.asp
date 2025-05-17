<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<% response.buffer = false %>
<% server.ScriptTimeout = 2 %>
<!--#INCLUDE FILE="TOOLS.asp" -->
<!--#INCLUDE FILE="Tools4Color.asp" -->
<html>
<head>
	<title>SELEZIONA IL COLORE..</title>
	<link rel="stylesheet" type="text/css" href="stili.css">
	<SCRIPT LANGUAGE="javascript"  src="utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<script language="JavaScript" type="text/javascript">
	function HighlightColor(color, advise){
		if (VerifyColor(color, advise)){
			var View_Color = document.getElementById("ColoreAttuale_Color");
			var View_Code = document.getElementById("ColoreAttuale_Code");
			View_Color.style.backgroundColor = color;
			View_Code.innerHTML = color;
		}
	}
	
	function SelectColor(color){
		if(VerifyColor(color, true)){
			//recupera controlli per impostazione colori
			var color_input = opener.document.getElementById("<%= request("field_name") %>_input");
			
			//inposta il colore nell'input di apertura
			color_input.value = color;
			opener.<%= request("field_name") %>_change(color_input);
			
			window.close();
		}
	}
	
	</script>
<% dim ColoreSelezionato, i, color, r, g, b
ColoreSelezionato = request.querystring("field_value")
%>
<body rightmargin="0" leftmargin="3" topmargin="5" onload="window.focus()">
	<table cellspacing="1" cellpadding="0" class="tabella_madre">
		<form id="FormColors" name="FormColors">
		<caption>
			Selezione colore
		</caption>
		<tr><th colspan="3">COLORI SCELTI</th></tr>
		<tr>
			<td class="label" style="width:30%;">colore selezionato</td>
			<% if ColoreSelezionato<>"" then
				if not IsColor(ColoreSelezionato) then %>
					<td class="errore" colspan="2">
						il codice "<%= ColoreSelezionato %>" non &egrave; un colore valido
					</td>
				<% else %>
					<td class="content" style="background-color:<%= ColoreSelezionato %>">&nbsp;</td>
					<td class="content" width="35%"><%= ColoreSelezionato %></td>
				<% end if
			else%>
				<td class="content">&nbsp;</td>
				<td class="content" width="35%">&nbsp;</td>
			<% end if %>
		</tr>
		<tr>
			<td class="label">colore attuale</td>
			<td class="content" id="ColoreAttuale_Color">&nbsp;</td>
			<td class="content" id="ColoreAttuale_Code" width="35%">&nbsp;</td>
		</tr>
		<tr><th colspan="3">COLORI DISPONIBILI</th></tr>
		<tr>
			<td class="label" colspan="3">
				clicca sul colore per selezionarlo
			</td>
		</tr>
		<tr>
			<td colspan="3" style="height:220px; overflow:auto; vertical-align:top;">
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<th class="L2" colspan="36">Colori "sicuri" per tutti i Browser</th>
					</tr>
					<tr><td>
					<% 	i = 0
						for r = 0 to 255 step 51
							for g = 0 to 255 step 51
								for b = 0 to 255 step 51
									color = "#" + FixLenght(Hex(r), "0", 2) + FixLenght(Hex(g), "0", 2) + FixLenght(Hex(b), "0", 2)
									if i mod 36 = 0 then
										if i> 1 then%> 
											</tr>
										<% end if %>
										<tr>
									<% end if
									CALL Write_ColorCell(color, 16)
									i = i + 1
								next
							next
						next %>
					</td></tr>
				</table>
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
					<tr>
						<th class="L2" colspan="32">Scala di grigi</th>
					</tr>
					<% for i = 255 to 0 step -1
						color = "#" + FixLenght(Hex(i), "0", 2) + FixLenght(Hex(i), "0", 2) + FixLenght(Hex(i), "0", 2)
						if (i+1) mod 32 = 0 then
							if i> 255 then%> 
								</tr>
							<% end if %>
							<tr>
						<% end if
						CALL Write_ColorCell(color, 9)
					next %>
				</table>
			</td>
		</tr>
		<tr><th colspan="3">ALTRO COLORE</th></tr>
		<tr>
			<td class="label">
				codice HTML del colore
			</td>
			<td class="content" colspan="2">
				<table cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<input type="text" maxlength="7" name="input_color" value="<%= ColoreSelezionato %>" size="7" 
								   onchange="HighlightColor(document.FormColors.input_color.value, true);"
								   onmouseover="HighlightColor(document.FormColors.input_color.value, false);"
								   style="letter-spacing:1px; padding-left:3px; text-align:center;" title="inserisci il codice html del colore desiderato">
						</td>
						<td>
							<a href="javascript:void(0);" title="conferma e seleziona il colore inserito"  <%= ACTIVE_STATUS %>
				   				class="button_input" onclick="SelectColor(document.FormColors.input_color.value)">
								CONFERMA
							</a>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td class="footer" colspan="3">
				<a class="button" href="javascript:window.close();" title="annulla" <%= ACTIVE_STATUS %>>
					ANNULLA
				</a>
			</td>
		</tr>
		</form>
	</table>
</body>
</html>
<script language="JavaScript" type="text/javascript">
<!--
	FitWindowSize(this);
//-->
</script>
<% 

sub Write_ColorCell(color, width) %>
	<td onclick="SelectColor('<%= color %>')" onmouseover="HighlightColor('<%= color %>', false)" 
		style="background-color:<%= color%>; height:12px; font-size:1px; width:<%= width %>px; cursor:pointer;" 
		title="click per selezionare il colore <%= color %>">
		&nbsp;
	</td>
<%end sub

Sub Luminosita(r, g, b)
	dim lumin
	for lumin = 1 to 10					'10 gradazioni di luminosita
		if r > 0 AND r < 255 then
			r = r + 17
		end if
		if r > 255 then r = 255 end if
		if (g > 0 OR r = 255) AND g < 255 then
			g = g + 17
		end if
		if g > 255 then g = 255 end if
		if (b > 0 OR g = 255) AND b < 255 then
			b = b + 17
		end if
		if b > 255 then b = 255 end if
		
		color = "#" + FixLenght(Hex(r), "0", 2) + FixLenght(Hex(g), "0", 2) + FixLenght(Hex(b), "0", 2)
		CALL Write_ColorCell(color, 16)
	next
End Sub
%>