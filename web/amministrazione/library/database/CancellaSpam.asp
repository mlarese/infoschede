<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>  
<% response.buffer = false %>
<!--#INCLUDE FILE="Tools4DataBase.asp" -->
<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="../Tools4Admin.asp" -->

<% 
'*****************************************************************************************************************
'verifica dei permessi
CALL VerificaPermessiUtente(true)
'*****************************************************************************************************************
%>
<html>
<head>
	<title>Amministrazione aggiornamenti database</title>
	<link rel="stylesheet" type="text/css" href="../stili.css">
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
	<meta name="robots" content="noindex,nofollow" />
	<meta name="copyright" content="Copyright © <%= Year(Date())%> - Next-aim" />
</head>
<body leftmargin="0" topmargin="0,5"  onload="window.focus();">
<%
dim strCap,strNote,strbox,strinizio,strcanc 
dim conn, rs, field, sql, sql_list, i,checks
set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application(request("ConnString")), "", ""
%>
<form name="form1" id="form1" action="" method="post">
<table width="740" cellspacing="0" cellpadding="0" border="0">
	<caption style="border:0px;">
		<table width="100%" cellspacing="0" cellpadding="0" border="0">                                                                                                                     
	  		<tr>
				<td align="right" style="padding-right:10px;">
					<a href="javascript:close();" class="menu" name="top">CHIUDI</a>
				</td>
	  		</tr>
		</table>
	</caption>
	<tr>
		<td >		
				<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:10px;">
					<% 
					strCap="((isnumeric(CAPElencoIndirizzi)=0)and([CAPElencoIndirizzi]!= '') and (LEN(NomeElencoIndirizzi)>5))"
					strNote= "((( NoteElencoIndirizzi)like '%http%')or (( NoteElencoIndirizzi)like '%comment%')or ((NomeElencoIndirizzi)like '%è%')) "
					strbox= "(IDElencoIndirizzi not IN (" & Request("MyCheckBox") &"))"
			        strinizio="select* from v_Indirizzario where"
					strcanc =" DELETE from v_Indirizzario where"
					
				 if request("MOSTRA SPAM")="ANTEPRIMA SPAM" then  
					 sql= strinizio+ strCap
				   elseif request("MOSTRA SPAM2")="ANTEPRIMA SPAM" then 
					 sql=strinizio + strNote 
				  end if
														
										
			      if request("cancella")="CANCELLA SELEZIONATO" then		              
			     	 sql= strcanc+"(IDElencoIndirizzi IN (" & Request("MyCheckBox") &"))"	
					   conn.execute(sql)
                     Response.write "spam cancellati"					 
                  End If
				    
					
				  if request("cancella1")="CANCELLA NON SELEZIONATO" then		              
					 sql= strcanc+strCap+"and"+strbox	
                       conn.execute(sql)			
                    Response.write "spam cancellati"					 
                  End If
					  
			      if request("cancella2")="CANCELLA NON SELEZIONATO" then		              
					 sql= strcanc+strNote+"and"+strbox	
                        conn.execute(sql)			
                      Response.write "spam cancellati"						
                  End If
					
					
					sql_list = split(sql, ";")
					
					conn.beginTrans
					
					for each sql in sql_list
						if trim(sql)<>"" then %>
						
						    <tr>
								<td class="caption">ANTEPRIMA DI PULIZIA SPAM "<%= request("ConnString") %>"</td>
								<td align="right" style="font-size: 1px;">
								</td>
							</tr>
						
							<tr>
								<th>CODICE ESEGUITO:</th>
							</tr>
							<tr>
								<td class="content">
									<%= TextHtmlEncode(sql) %>
								</td>
							</tr>
							<% set rs = conn.execute(sql, , adCmdText)
                            if rs.state = adStateOpen then %>
							<tr>
								<td class="content">
									rs.bof = <%= rs.bof %><br>
									rs.recordcount = <%= rs.recordcount %><br>
									rs.eof = <%= rs.eof %><br>
								</td>
							</tr>
                            <% end if %>
							<tr>
								<td class="content_b ok">
									<%if request("ESEGUI")="ESEGUI" then%>
										SPAM PRESENTI:
									<% else %>
										SPAM PRESENTI:
									<% end if%>
								</td>
							</tr>
							<%if rs.state = adStateOpen then%>
								<tr>
									<td>
										<table width="100%" border="0" cellspacing="1" cellpadding="0">								
									      <tr>	 
									         <th class="center ok" nowrap>chekbox </th>
										 
										     	<th class="center ok" nowrap>n. riga</th>
											    	<%for each Field in rs.Fields%>
													  <th>&nbsp;<%= Field.name %></th>
												     <%next%>
												  <th class="center ok" nowrap>n. riga</th>										
											  
											
									      </tr>
											  <%i = 1
											    while not rs.eof %>
											  <tr>
												
												      <td class="content_center ok">
													    <input type=checkbox value=<%= rs("IDElencoIndirizzi")%> name=MyCheckBox>
												     </td>	
													
													<td class="content_center ok"><%= i %></td>
													<%for each Field in rs.Fields%>
														<td class="content">
															<%if isNull(Field.value) then%>
																NULL
															<%else%>
																&nbsp;<%= Field.value %>
															<%end if%>
														</td>
													<%next%>
													<td class="content_center ok"><%= i %></td>
															
												</tr>
																																				
												<%i = i + 1
												rs.MoveNext
											wend %>
									      	
										
										</table>		

                                     <tr > 
										  <td class="label_no_width" colspan="3">
										      
										    <input style="width:3.5%;" type="submit" class="button" value='CANCELLA SELEZIONATO' onclick="if(confirm('Eliminare spam?'))name='cancella'"/>								
										     &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; cancella contatti identificati come span selezionati
										  </td>
									
								    </tr>
									
									<tr > 
										  <td class="label_no_width" colspan="3">	  
										     <%if request("MOSTRA SPAM")="ANTEPRIMA SPAM" then %>
                                              <input style="width:3.5%;" type="submit" class="button" value='CANCELLA NON SELEZIONATO' onclick="if(confirm('Eliminare spam?'))name='cancella1'"/>
											   &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;  cancella contatti identificati come span non selezionati
											  <%elseif request("MOSTRA SPAM2")="ANTEPRIMA SPAM" then%>											  
                                              <input style="width:3.5%;" type="submit" class="button" value='CANCELLA NON SELEZIONATO' onclick="if(confirm('Eliminare spam?'))name='cancella2'"/>	 
                                               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; cancella contatti identificati come span selezionati
											  <%end if %>  
										   </td>
										   
								</tr>
								
							
								
							<%else%>
								<tr><td class="content">Istruzione senza risultati</td></tr>
							<%end if
							set rs = nothing
						end if
					next
					
					if request("ESEGUI")="ESEGUI" then
						conn.CommitTrans
					else
						conn.RollbackTrans
					end if %>
					
				</table>
		</td>
	</tr>
		
</table>


<table width="740" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td>
			<table cellspacing="1" cellpadding="0" class="tabella_madre">
				<caption>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td class="caption"> PULIZIA CONTATTI SPAM PER DATABASE </td>
							<td align="right" style="font-size: 1px;">
							</td>
						</tr>
					</table>
				</caption>
				 <th colspan="4" >filtro contatti identificati come spam attraverso un confronto dei 'CAPElencoIndirizzi'non numerici presenti nel database </th>
				 
				<tr>
				    <td class="label_no_width" colspan="3">
				       elenco  contatti identificati come spam
                     </td>
                        <td class="content_center">
                                   <input style="width:50%;" type="submit" class="button" name="MOSTRA SPAM" value="ANTEPRIMA SPAM">
							    </td>
				</tr>
				
				

				 <th colspan="4" >filtro contatti identificati come  spam attraverso un confronto delle 'NoteElencoIndirizzi' presenti nel database</th>
				
				<tr>
				    <td class="label_no_width" colspan="3">
				         elenco  contatti identificati come spam
                     </td>
                        <td class="content_center">
                                   <input style="width:50%;" type="submit" class="button" name="MOSTRA SPAM2" value="ANTEPRIMA SPAM">
							    </td>
				</tr>
				
			
								
			</table>
		</td>
	</tr>
	
</table>
</form>
</body>


</html>


<% 
conn.close
set conn = nothing
%>