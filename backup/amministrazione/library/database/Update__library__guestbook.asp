<%
'*******************************************************************************************
'INSTALLA NEXT-GUESTBOOK
'...........................................................................................
'aggiunge database per gestione next-guestbook
'...........................................................................................

function Install__NEXTGUESTBOOK(conn)
	Install__NEXTGUESTBOOK = _
		" CREATE TABLE tb_guestbook (" & _
		" 	IdGuest int IDENTITY (1, 1) NOT NULL, " & _
		" 	Data datetime NULL, " & _
		" 	Visibile bit NOT NULL, " & _
		" 	Id_contatto int NOT NULL, " & _
		" 	Messaggio " + SQL_CharField(Conn, 0) + "NULL, " &_
		"	Oggetto " + SQL_CharField(Conn, 250) + "NULL, " &_
		"	Log_richiesta " + SQL_CharField(Conn, 0) + "NULL , " &_
		" 	risposta " + SQL_CharField(Conn, 0) + "NULL); " &_
		" ALTER TABLE tb_guestbook ADD CONSTRAINT FK_tb_guestbook__tb_Indirizzario " &_
		" 	FOREIGN KEY (Id_contatto) REFERENCES Tb_Indirizzario (IDElencoIndirizzi) " &_
		" 	ON UPDATE CASCADE ON DELETE CASCADE; "
end function

'*******************************************************************************************
'ATTIVAZIONE NEXT-GUESTBOOK CON RELATIVI PARAMETRI
'...........................................................................................
function Activate_GUESTBOOK(conn)
	Activate_GUESTBOOK ""
end function


'*******************************************************************************************
'AGGIORNAMENTO GUESTBOOK 1
'...........................................................................................
'aggiungo cancellazione in cascata per i commenti e l'indice altrimenti
'ogni volta che cancello una voce dell'indice dovrei fare il controllo (cancellazioni automatiche?)
'...........................................................................................
function Aggiornamento__GUESTBOOK__1(conn)
	Aggiornamento__GUESTBOOK__1 = _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_guestbook ADD" + vbCrLf + _
		" 	risposta " + SQL_CharField(Conn, 0) + ";" + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO GUESTBOOK 2
'...........................................................................................
'	Giacomo, 07/02/2012
'...........................................................................................
'aggiungo tabella categorie
'...........................................................................................
function Aggiornamento__GUESTBOOK__2(conn)
	Aggiornamento__GUESTBOOK__2 = _
		" CREATE TABLE "& SQL_Dbo(conn) &"tb_guestbook_categorie (" + vbCrLf + _
		"	cat_id "& SQL_PrimaryKey(conn, "tb_guestbook_categorie") + ", " + vbCrLf + _
		SQL_MultiLanguageFieldComplete(conn, " cat_nome_<lingua> " + SQL_CharField(Conn, 500)) + "," + vbCrLf + _
		" 	cat_ordine int NULL, " + vbCrLf + _
		" 	cat_visibile bit NULL " + vbCrLf + _
		"); " + vbCrLf + _
		" ALTER TABLE "& SQL_Dbo(conn) &"tb_guestbook ADD" + vbCrLf + _
		" 	id_categoria int NULL; " + vbCrLf + _
		SQL_AddForeignKey(conn, "tb_guestbook", "id_categoria", "tb_guestbook_categorie", "cat_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO GUESTBOOK 3
'...........................................................................................
'	Giacomo, 07/02/2012
'...........................................................................................
'   aggiunge parametro per attivare le categorie
'...........................................................................................
function Aggiornamento__GUESTBOOK__3(conn)
	Aggiornamento__GUESTBOOK__3 = " SELECT * FROM AA_Versione "
end function

sub AggiornamentoSpeciale__GUESTBOOK__3(conn)
	if cIntero(GetValueList(conn, NULL, "SELECT COUNT(*) FROM tb_siti WHERE id_sito = " & NEXTGUESTBOOK)) > 0 then
		CALL AddParametroSito(conn, "ATTIVA_CATEGORIE_GUESTBOOK", _
									null, _
									"Attiva le categorie del guestbook.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTGUESTBOOK, _
									null, null, null, null, null)
	end if
end sub
'*******************************************************************************************


%>