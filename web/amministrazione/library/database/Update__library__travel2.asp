<%
'...........................................................................................
'...........................................................................................
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-travel 2.0
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE TRAVEL2
'...........................................................................................
function Install__TRAVEL2(conn)
	Install__TRAVEL2 = _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_viaggi (" + _
		"	vi_id " & SQL_PrimaryKey(conn, "TAtb_viaggi") + ", " +_ 
		"	vi_categoria_id INT NOT NULL, " + _
		"	vi_destinazione_id INT NOT NULL, " + _
			SQL_MultiLanguageField("vi_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	vi_ordine INT NULL, " + _
		" 	vi_visibile BIT NULL, " + _
		"	vi_NEXTweb_ps INT NULL, " + _
		"	vi_file " + SQL_CharField(Conn, 255) + ", " + _
			SQL_MultiLanguageField("vi_descr_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			SQL_MultiLanguageField("vi_partenza_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			SQL_MultiLanguageField("vi_durata_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
			AddInsModFields("vi") + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_destinazioni (" + _
		"	de_id " & SQL_PrimaryKey(conn, "TAtb_destinazioni") + ", " +_ 
			SQL_MultiLanguageField("de_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageField("de_descr_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	de_ordine INT NULL " + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_richieste(" + _
		"	ri_id " & SQL_PrimaryKey(conn, "TAtb_richieste") + ", " +_ 
		"	ri_viaggio_id INT NOT NULL, " + _
		" 	ri_indirizzario_id INT NOT NULL, " + _
		"	ri_data SMALLDATETIME NULL, "& vbCrLf & _
		"	ri_note " + SQL_CharField(Conn, 0) + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_categorie(" + _
		"	cat_id " & SQL_PrimaryKey(conn, "TAtb_categorie") + ", " +_ 
			SQL_MultiLanguageField("cat_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	cat_foto " + SQL_CharField(Conn, 255) + ", " + _
		"	cat_codice " + SQL_CharField(Conn, 255) + ", " + _
			SQL_MultiLanguageField("cat_descr_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
		"	cat_foglia BIT NULL, " + _
		"	cat_livello INT NULL, " + _
		"	cat_padre_id INT NULL, " + _
		"	cat_ordine INT NULL, " + _
		"	cat_ordine_assoluto " + SQL_CharField(Conn, 255) + ", " + _
		"	cat_tipologia_padre_base INT NULL, " + _
		"	cat_visibile BIT NULL, " + _
		"	cat_albero_visibile BIT NULL, " + _
		"	cat_tipologie_padre_lista " + SQL_CharField(Conn, 255) + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_raggruppamenti_descrittori(" + _
		"	rag_id " & SQL_PrimaryKey(conn, "TAtb_raggruppamenti_descrittori") + ", " +_ 
			SQL_MultiLanguageField("rag_titolo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	rag_ordine INT NULL, " + _
		"	rag_codice " + SQL_CharField(Conn, 255) + ", " + _
		"	rag_note " + SQL_CharField(Conn, 0) + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_descrittori(" + _
		"	des_id " & SQL_PrimaryKey(conn, "TAtb_descrittori") + ", " +_ 
			SQL_MultiLanguageField("des_nome_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	des_raggruppamento_id INT NULL, " + _
		"	des_tipo INT NULL, " + _
		"	des_principale BIT NULL, " + _
			SQL_MultiLanguageField("des_unita_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	des_img " + SQL_CharField(Conn, 255) + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TArel_descrittori_categorie(" + _
		"	rdc_id " & SQL_PrimaryKey(conn, "TArel_descrittori_categorie") + ", " +_ 
		"	rdc_categoria_id INT NULL, " + _
		"	rdc_descrittore_id INT NULL, " + _
		"	rdc_ordine INT NULL " + _
		" ); " + _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TArel_viaggi_descrittori(" + _
		"	rvd_id " & SQL_PrimaryKey(conn, "TArel_viaggi_descrittori") + ", " +_ 
		"	rvd_viaggio_id INT NULL, " + _
		"	rvd_descrittore_id INT NULL, " + _
			SQL_MultiLanguageField("rvd_valore_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
			SQL_MultiLanguageField("rvd_memo_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
		"	rvd_data SMALLDATETIME NULL " + _
		" ); " + _
		AddInsModRelations(conn, "TAtb_viaggi", "vi") + _
		SQL_AddForeignKey(conn, "TAtb_richieste", "ri_viaggio_id", "TAtb_viaggi", "vi_id", true, "") + _
		SQL_AddForeignKey(conn, "TAtb_richieste", "ri_indirizzario_id", "tb_indirizzario", "idElencoIndirizzi", true, "") + _
		SQL_AddForeignKey(conn, "TAtb_viaggi", "vi_destinazione_id", "TAtb_destinazioni", "de_id", true, "") + _
		SQL_AddForeignKey(conn, "TAtb_viaggi", "vi_categoria_id", "TAtb_categorie", "cat_id", true, "") + _
		SQL_AddForeignKey(conn, "TAtb_categorie", "cat_padre_id", "TAtb_categorie", "cat_id", false, "") + _
		SQL_AddForeignKey(conn, "TAtb_categorie", "cat_tipologia_padre_base", "TAtb_categorie", "cat_id", false, "") + _
		SQL_AddForeignKey(conn, "TAtb_descrittori", "des_raggruppamento_id", "TAtb_raggruppamenti_descrittori", "rag_id", false, "") + _
		SQL_AddForeignKey(conn, "TArel_descrittori_categorie", "rdc_descrittore_id", "TAtb_descrittori", "des_id", true, "") + _
		SQL_AddForeignKey(conn, "TArel_descrittori_categorie", "rdc_categoria_id", "TAtb_categorie", "cat_id", true, "") + _
		SQL_AddForeignKey(conn, "TArel_viaggi_descrittori", "rvd_descrittore_id", "TAtb_descrittori", "des_id", true, "") + _
		SQL_AddForeignKey(conn, "TArel_viaggi_descrittori", "rvd_viaggio_id", "TAtb_viaggi", "vi_id", true, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-TRAVEL2 1
'...........................................................................................
'   aggiunge il codice alle escursioni
'...........................................................................................
function Aggiornamento__TRAVEL2__1(conn)
    Aggiornamento__TRAVEL2__1 = _
        " ALTER TABLE " & SQL_Dbo(Conn) & "TAtb_viaggi ADD"& vbCrLf & _
	  	"		vi_codice_it " + SQL_CharField(Conn, 100) + " NULL,"& vbCrLf & _
	  	"		vi_codice_en " + SQL_CharField(Conn, 100) + " NULL,"& vbCrLf & _
	  	"		vi_codice_fr " + SQL_CharField(Conn, 100) + " NULL,"& vbCrLf & _
	  	"		vi_codice_es " + SQL_CharField(Conn, 100) + " NULL,"& vbCrLf & _
	  	"		vi_codice_de " + SQL_CharField(Conn, 100) + " NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TRAVEL2 2
'...........................................................................................
'   modifica per aggiunta agenzie
'...........................................................................................
function Aggiornamento__TRAVEL2__2(conn)
    Aggiornamento__TRAVEL2__2 = _
		" CREATE TABLE " & SQL_Dbo(Conn) & "TAtb_agenzie ("& vbCrLf & _
		"	age_id " & SQL_PrimaryKeyInt(conn, "TAtb_agenzie") + ", "& vbCrLf & _
			SQL_MultiLanguageField("age_descrizione_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", "& vbCrLf & _
		"	age_ordine INT NULL"& vbCrLf & _
		" ); "& vbCrLf & _
		SQL_RemoveForeignKey(conn, "TAtb_richieste", "ri_indirizzario_id", "tb_indirizzario", true, "FK_TAtb_richieste__tb_indirizzario") & vbCrLf & _
		SQL_AddForeignKey(conn, "TAtb_richieste", "ri_indirizzario_id", "tb_indirizzario", "idElencoIndirizzi", false, "") & vbCrLf & _
		SQL_AddForeignKey(conn, "TAtb_agenzie", "age_id", "tb_utenti", "ut_id", true, "") & vbCrLf & _
		" CREATE VIEW " & SQL_Dbo(Conn) & "TAv_agenzie AS"& vbCrLf & _
		" 	SELECT * FROM (TAtb_agenzie a"& vbCrLf & _
		" 	INNER JOIN tb_utenti u ON a.age_id = u.ut_id)"& vbCrLf & _
		"	INNER JOIN tb_indirizzario i ON u.ut_nextCom_id = i.idElencoIndirizzi;"& vbCrLf & _
        " ALTER TABLE " & SQL_Dbo(Conn) & "TAtb_viaggi ADD"& vbCrLf & _
	  	"		vi_agenzia_id INT NULL;"& vbCrLf & _
		SQL_AddForeignKey(conn, "TAtb_viaggi", "vi_agenzia_id", "TAtb_agenzie", "age_id", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TRAVEL2 3
'...........................................................................................
'   modifica il campo memo della tabella TArel_viaggi_descrittori modificandone il tipo
'   in memo.
'...........................................................................................
function Aggiornamento__TRAVEL2__3(conn)
    Aggiornamento__TRAVEL2__3 = _
		" ALTER TABLE TArel_viaggi_descrittori "& vbCrLf & _
		" ALTER COLUMN rvd_memo_it " + SQL_CharField(Conn, 0) + " NULL; " & vbCrLf & _
		" ALTER TABLE TArel_viaggi_descrittori "& vbCrLf & _
		" ALTER COLUMN rvd_memo_en " + SQL_CharField(Conn, 0) + " NULL; " & vbCrLf & _
		" ALTER TABLE TArel_viaggi_descrittori "& vbCrLf & _
		" ALTER COLUMN rvd_memo_fr " + SQL_CharField(Conn, 0) + " NULL; " & vbCrLf & _
		" ALTER TABLE TArel_viaggi_descrittori "& vbCrLf & _
		" ALTER COLUMN rvd_memo_de " + SQL_CharField(Conn, 0) + " NULL; " & vbCrLf & _
		" ALTER TABLE TArel_viaggi_descrittori "& vbCrLf & _
		" ALTER COLUMN rvd_memo_es " + SQL_CharField(Conn, 0) + " NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-TRAVEL2 4
'...........................................................................................
'   aggiunge campi 'città partenza' 'periodo interesse' 'numero partecipanti' alle richieste
'...........................................................................................
function Aggiornamento__TRAVEL2__4(conn)
    Aggiornamento__TRAVEL2__4 = _
        " ALTER TABLE " & SQL_Dbo(Conn) & "TAtb_richieste ADD"& vbCrLf & _
	  	"		ri_citta_partenza " + SQL_CharField(Conn, 250) + " NULL,"& vbCrLf & _
	  	"		ri_periodo " + SQL_CharField(Conn, 250) + " NULL,"& vbCrLf & _
	  	"		ri_numero_partecipanti " + SQL_CharField(Conn, 250) + " NULL"
end function
'*******************************************************************************************




%>