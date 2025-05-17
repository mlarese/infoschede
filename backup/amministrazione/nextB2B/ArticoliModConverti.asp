<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<!--#INCLUDE FILE="Tools_B2B.asp" -->
<!--#INCLUDE FILE="Tools4Save_B2B.asp" -->
<% 
dim conn, rs, rsv, sql
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
set rs = Server.CreateObject("ADODB.Recordset")
set rsv = Server.CreateObject("ADODB.Recordset")

if Request.ServerVariables("REQUEST_METHOD")="POST" then
	if request("SALVA")<>"" then
	
		dim objVariante 
		set objVariante = new GestioneVariante
		set objVariante.conn = conn
	
		conn.begintrans
		
		'converte articolo in articolo con varianti
		sql = "UPDATE gtb_articoli SET art_varianti=1 WHERE art_id=" & cIntero(request("ID"))
		conn.execute(sql)
		
		'assegna valore variante
		dim var, listaValori, RelId
		listaValori = ""
		
		relId = cIntero(GetValueList(conn, rs, "SELECT TOP 1 rel_id FROM grel_art_valori WHERE rel_art_id=" & request("ID")))
		
		for each var in request.form
			if left(var, 7) = "valori_" then
				'inserisce valore variante per l'articolo
				if cIntero(request.form(var))>0 then
					listaValori = listaValori & IIF(listaValori <> "", ",", "") & cIntero(request.form(var))
				end if
			end if
		next
		
		if listaValori<>"" then
		
			CALL objVariante.InserisciValoriVariante(RelId, listaValori)
			CALL objVariante.ImpostaOrdineVariante(RelId)
			
			conn.committrans
			%>
			<script language="JavaScript" type="text/javascript">
				opener.location.reload(true);
				window.close();
			</script>
			<% response.end
		else
			conn.rollbacktrans
		end if
	end if
end if

sql = " SELECT * FROM gtb_articoli INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + _
	  " WHERE art_id=" & cIntero(request("ID"))
rs.open sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

%>

<%'--------------------------------------------------------
sezione_testata = "converti in Articol con Varianti"%>
<!--#INCLUDE FILE="../library/Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
%>

<div id="content_ridotto">
	<form action="" method="post" id="form1" name="form1">
		<table cellspacing="1" cellpadding="0" class="tabella_madre">
			<caption>
				Converti in Articolo con Varianti
			</caption>
			<tr><th colspan="7">DATI ARTICOLO</th></tr>
			<% CALL ArticoloScheda (conn, rs, rsv) %>
			<tr><th colspan="7">VARIANTI DELL'ARTICOLO</th></tr>
			<% sql = " SELECT * FROM gtb_varianti INNER JOIN gtb_valori ON gtb_varianti.var_id = gtb_valori.val_var_id " + _
					 " ORDER BY var_nome_it, val_nome_it"
			rsv.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
			if rsv.eof then%>
				<tr><td class="content_b alert" colspan="7">Nessuna variante definita.<br>Prima di inserire l'articolo inserire le varianti ed i relativi valori.</th></tr>
			<%else
				dim Current%>
				<tr>
					<td colspan="7">
						<table cellspacing="1" cellpadding="0" class="tabella_madre" style="border:0px;">
							<tr>
								<th class="L2">VARIANTE</th>
								<th class="L2" width="4%">&nbsp;</th>
								<th class="L2" width="8%">CODICE</th>
								<th class="L2">VALORE</th>
							</tr>
							<%Current = ""
							while not rsv.eof %>
								<tr>
									<% if Current <> rsv("var_id") then
										Current = rsv("var_id") %>
										<td class="content"><%= rsv("var_nome_it") %></td>
									<% else %>	
										<td class="content">&nbsp;</td>
									<% end if %>
									<td class="content"><input type="radio" class="checkbox" id="valore_<%= rsv("var_id") %>_<%= rsv("val_id") %>" name="valori_<%= rsv("var_id") %>" value=" <%= rsv("val_id") %> " <%= chk(instr(1, request("valori_" & rsv("var_id")) , " " & rsv("val_id") & " ", vbTextCompare)) %>></td>
									<td class="content"><%= rsv("val_cod_int") %></td>
									<td class="content"><%= rsv("val_nome_it") %></td>
								</tr>
								<%rsv.MoveNext
							wend%>
						</table>
					</td>
				</tr>
			<% end if %>
			
			<tr>
				<td class="footer" colspan="7">
					<input style="width:23%;" type="submit" class="button" name="salva" value="SALVA">
					<input style="width:23%;" type="button" class="button" name="annulla" value="ANNULLA" onclick="window.close();">
				</td>
			</tr>
		</table>
	</form>
</div>
</body>
</html>
<% rs.close
conn.close
set rs = nothing
set rsv = nothing
set conn = nothing %>
			