<%
'*******************************************************************************************
'INSTALLA NEXT-COMMENT
'...........................................................................................
'aggiunge database per gestione next-comment
'...........................................................................................
function Install__NEXTCOMMENT(conn)
	Install__NEXTCOMMENT = _
		" CREATE TABLE dbo.tb_comments_valutazioni ( " + vbCrLF + _
		"	val_id INT IDENTITY(1, 1) NOT NULL CONSTRAINT PK_tb_comments_valutazioni PRIMARY KEY CLUSTERED, " + vbCrLF + _
		"	val_valutazione INT NULL, " + vbCrLF + _
		"	val_nome_it NVARCHAR(100) NULL, " + vbCrLF + _
		"	val_nome_en NVARCHAR(100) NULL, " + vbCrLF + _
		"	val_nome_fr NVARCHAR(100) NULL, " + vbCrLF + _
		"	val_nome_es NVARCHAR(100) NULL, " + vbCrLF + _
		"	val_nome_de NVARCHAR(100) NULL, " + vbCrLF + _
		"	val_icona NVARCHAR(255) NULL, " + vbCrLF + _
		"	val_default BIT NULL " + vbCrLF + _
		" ); " + vbCrLF + _
		" CREATE TABLE dbo.tb_comments ( " + vbCrLF + _
		"	com_id " + SQL_PrimaryKey(conn, "tb_comments") + ", " + vbCrLF + _
		"	com_idx_id INT NOT NULL, " + vbCrLF + _
		"	com_contatto_id INT NULL, " + vbCrLF + _
		"	com_val_id INT NULL, " + vbCrLF + _
		"	com_comment NTEXT NULL, " + vbCrLF + _
		"	com_validate BIT NULL, " + vbCrLF + _
		"	com_validateData SMALLDATETIME NULL, " + vbCrLF + _
		"	com_validateAdmin_id INT NULL, " + vbCrLF + _
		AddInsModFields("com") + _
		" ); " + vbCrLF + _
		" ALTER TABLE dbo.tb_comments ADD CONSTRAINT FK_tb_comments_tb_contents_index" + vbCrLf + _
		" 	FOREIGN KEY (com_idx_id) REFERENCES tb_contents_index (idx_id);" + vbCrLF + _
		" ALTER TABLE dbo.tb_comments ADD CONSTRAINT FK_tb_comments_tb_indirizzario" + vbCrLf + _
		" 	FOREIGN KEY (com_contatto_id) REFERENCES tb_indirizzario (idElencoIndirizzi);" + vbCrLf + _
		" ALTER TABLE dbo.tb_comments NOCHECK CONSTRAINT FK_tb_comments_tb_indirizzario;" + vbCrLf + _
		" ALTER TABLE dbo.tb_comments ADD CONSTRAINT FK_tb_comments_tb_admin_val" + vbCrLf + _
		" 	FOREIGN KEY (com_validateAdmin_id) REFERENCES tb_admin (id_admin);" + vbCrLf + _
		" ALTER TABLE dbo.tb_comments NOCHECK CONSTRAINT FK_tb_comments_tb_admin_val;" + vbCrLf + _
		AddInsModRelations(conn, "tb_comments", "com") + _
		" ALTER TABLE tb_comments ADD CONSTRAINT FK_tb_comments__tb_comments_valutazioni " + vbCrLf + _
	    "	FOREIGN KEY (com_val_id) REFERENCES tb_comments_valutazioni (val_id) " + vbCrLf + _
        "   ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************

'*******************************************************************************************
'ATTIVAZIONE NEXT-COMMENT CON RELATIVI PARAMETRI
'...........................................................................................
function Activate_NEXTCOMMENT(conn)
    Activate_NEXTCOMMENT = _
                " INSERT INTO tb_siti(sito_nome, sito_dir, sito_p1, sito_amministrazione, id_sito, sito_prmEsterni_Admin, sito_prmesterni_sito ) " + _
				"     VALUES('NEXT-comment [gestione commenti degli utenti]', 'NEXTcomment', 'COMMENT_USER', 1, " & _
				NEXTCOMMENT & ", '', ''); " + vbCrLf + _
                " INSERT INTO tb_rubriche (nome_rubrica, locked_rubrica, rubrica_esterna, note_rubrica) " + _
                "     VALUES('Commenti - elenco completo', 1, 0, 'Utilizzata da NEXT-COMMENT'); " + _
                " INSERT INTO tb_siti_parametri (par_key, par_value, par_sito_id ) " + _
                "     SELECT 'RUBRICA_ID', "& SQL_String(conn, "id_rubrica") &", " & NEXTCOMMENT & _
                "       FROM tb_rubriche " + _
                "       WHERE nome_rubrica LIKE 'Commenti%' AND note_rubrica LIKE '%NEXT-COMMENT%' ; " + _
				" INSERT INTO tb_siti_parametri (par_key, par_value, par_sito_id ) " + _
                "     SELECT TOP 1 'RUBRICA_ALTERNATIVA_ID', "& SQL_String(conn, "id_rubrica") &", " & NEXTCOMMENT & _
                "       FROM tb_rubriche " + _
                "       WHERE nome_rubrica LIKE '%Contatti%' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO COMMENT 1
'...........................................................................................
'aggiungo cancellazione in cascata per i commenti e l'indice altrimenti
'ogni volta che cancello una voce dell'indice dovrei fare il controllo (cancellazioni automatiche?)
'...........................................................................................
function Aggiornamento__COMMENT__1(conn)
	Aggiornamento__COMMENT__1 = _
		" ALTER TABLE tb_comments DROP CONSTRAINT FK_tb_comments_tb_contents_index;" + vbCrLf + _
		" ALTER TABLE dbo.tb_comments ADD CONSTRAINT FK_tb_comments_tb_contents_index" + vbCrLf + _
		" 	FOREIGN KEY (com_idx_id) REFERENCES tb_contents_index (idx_id)" + vbCrLf + _
        "   ON UPDATE CASCADE ON DELETE CASCADE;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO COMMENT 2
'...........................................................................................
' la valutazione non è più campo obbligatorio del commento
'...........................................................................................
function Aggiornamento__COMMENT__2(conn)
	Aggiornamento__COMMENT__2 = _
		" ALTER TABLE dbo.tb_comments DROP CONSTRAINT FK_tb_comments__tb_comments_valutazioni;" + vbCrLf + _
		" ALTER TABLE dbo.tb_comments ADD CONSTRAINT FK_tb_comments__tb_comments_valutazioni " + vbCrLf + _
	    "	FOREIGN KEY (com_val_id) REFERENCES tb_comments_valutazioni (val_id) " + vbCrLf + _
        "   ON UPDATE SET NULL ON DELETE SET NULL;" + vbCRLF + _
		"  ALTER TABLE dbo.tb_comments NOCHECK CONSTRAINT FK_tb_comments__tb_comments_valutazioni;"
end function
'*******************************************************************************************

%>