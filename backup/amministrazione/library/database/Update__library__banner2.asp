<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-banner 2.0
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE NEXT-BANNER 2.0
'...........................................................................................
function Install__NEXTBANNER2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__NEXTBANNER2 = _
					""
		case DB_SQL
			Install__NEXTBANNER2 = _
				" CREATE TABLE dbo.ADtb_tipiBanner( " + vbCrLf + _
				"   tipoB_id int IDENTITY(1,1) NOT NULL, " + vbCrLf + _
				"	tipoB_nome nvarchar(50) NULL, " + vbCrLf + _
				"	tipoB_note ntext NULL, " + vbCrLf + _
				"   CONSTRAINT PK_ADtb_tipiBanner PRIMARY KEY CLUSTERED (tipoB_id) " + vbCrLF + _
                " ) ; " + vbCrLf + _
                " CREATE TABLE dbo.ADtb_banner( " + vbCrLf + _
                "   ban_id int IDENTITY(1,1) NOT NULL, " + vbCrLf + _
                "	ban_nome nvarchar(50) NULL, " + vbCrLf + _
                "   ban_image nvarchar(50) NULL, " + vbCrLf + _
                "   ban_link nvarchar(250) NULL, " + vbCrLf + _
                "   ban_alt ntext NULL, " + vbCrLf + _
                "	ban_tipo_id int NOT NULL, " + vbCrLf + _
                "   ban_cliente_id int NOT NULL, " + vbCrLf + _
                "   CONSTRAINT PK_ADtb_banner PRIMARY KEY CLUSTERED (ban_id), " + vbCrLf + _
                "   CONSTRAINT FK_ADtb_banner__ADtb_tipiBanner " + vbCrLF + _
                "       FOREIGN KEY(ban_tipo_id) REFERENCES ADtb_tipiBanner(tipoB_id) " + vbCrLf + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + vbCrLf + _
                "   CONSTRAINT FK_ADtb_banner__tb_Indirizzario " + vbCRLF + _
                "       FOREIGN KEY(ban_cliente_id) REFERENCES tb_Indirizzario(IDElencoIndirizzi) " + vbCrLf + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + vbCrLf + _
                " ) ; " + vbCrLF + _
                " CREATE TABLE dbo.ADtb_banner_posizioni( " + vbCrLf + _
                "	bp_id int IDENTITY(1,1) NOT NULL, " + vbCrLf + _
                "   bp_impress_iniz int NULL, " + vbCrLf + _
                "   bp_impress int NULL, " + vbCrLf + _
                "   bp_data_iniz smalldatetime NULL, " + vbCrLf + _
                "   bp_data_fine smalldatetime NULL, " + vbCrLf + _
                "   bp_click_iniz int NULL, " + vbCrLf + _
                "   bp_click int NULL, " + vbCrLf + _
                "   bp_banner_id int NOT NULL, " + vbCrLf + _
                "   bp_index_id int NOT NULL, " + vbCrLF + _
                "   CONSTRAINT PK_ADtb_banner_posizioni PRIMARY KEY CLUSTERED (bp_id), " + vbCrLf + _
                "   CONSTRAINT FK_ADtb_banner_posizioni__ADtb_banner " + vbCrLF + _
                "       FOREIGN KEY(bp_banner_id) REFERENCES ADtb_banner (ban_id) " + vbCrLf + _
                "       ON UPDATE CASCADE ON DELETE CASCADE, " + vbCrLF + _
                "   CONSTRAINT FK_ADtb_banner_posizioni__tb_contents_index " + vbCrLF + _
                "       FOREIGN KEY(bp_index_id) REFERENCES tb_contents_index (idx_id) " + vbCrLf + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + vbCrLf + _
				" ); " + vbCrLf + _
                " CREATE TABLE dbo.ADtb_storico_impress( " + vbCrLf + _
                "	sti_id int IDENTITY(1,1) NOT NULL, " + vbCrLf + _
                "   sti_data SMALLDATETIME NULL, " + vbCrLF + _
                "   sti_ora INT NULL, " + vbCRLF + _
                "   sti_count INT NULL, " + vbCrLF + _
                "   sti_posizione_id INT NOT NULL, " + vbCrLf + _
                "   CONSTRAINT PK_ADtb_storico_impress PRIMARY KEY CLUSTERED (sti_id), " + vbCrLf + _
                "   CONSTRAINT FK_ADtb_storico_impress__ADtb_banner_posizioni " + vbCrLf + _
                "       FOREIGN KEY (sti_posizione_id) REFERENCES ADtb_banner_posizioni(bp_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + vbcrLf + _
                " ); " + vbCrLf + _
                " CREATE TABLE dbo.ADlog_impress( " + vbCrLf + _
                "	log_id int IDENTITY(1,1) NOT NULL, " + vbCrLf + _
                "   log_data SMALLDATETIME NULL, " + vbCrLF + _
                "   log_posizione_id INT NOT NULL, " + vbCrLf + _
                "   CONSTRAINT PK_ADlog_impress PRIMARY KEY CLUSTERED (log_id), " + vbCrLf + _
                "   CONSTRAINT FK_ADlog_impress__ADtb_banner_posizioni " + vbCrLf + _
                "       FOREIGN KEY (log_posizione_id) REFERENCES ADtb_banner_posizioni(bp_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + vbcrLf + _
                " ); " + vbCrLf + _
                " CREATE TABLE dbo.ADlog_click( " + vbCrLf + _
                "	log_id int IDENTITY(1,1) NOT NULL, " + vbCrLf + _
                "   log_data SMALLDATETIME NULL, " + vbCrLF + _
                "   log_posizione_id INT NOT NULL, " + vbCrLf + _
                "   log_request NTEXT NOT NULL, " + vbCrLf + _
                "   log_ip nvarchar(15) NOT NULL, " + vbCrLf + _
                "   CONSTRAINT PK_ADlog_click PRIMARY KEY CLUSTERED (log_id), " + vbCrLf + _
                "   CONSTRAINT FK_ADlog_click__ADtb_banner_posizioni " + vbCrLf + _
                "       FOREIGN KEY (log_posizione_id) REFERENCES ADtb_banner_posizioni(bp_id) " + vbCrLF + _
                "       ON UPDATE CASCADE ON DELETE CASCADE " + vbcrLf + _
                " ); "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER  1
'...........................................................................................
'rimuove tabella banner e posizioni per unificare il tutto in una tabella contratti
'...........................................................................................
function Aggiornamento__NEXTBANNER2__1(conn)
    Aggiornamento__NEXTBANNER2__1 = _
            " ALTER TABLE ADtb_banner DROP CONSTRAINT FK_ADtb_banner__ADtb_tipiBanner; " + _
            " ALTER TABLE ADtb_banner DROP CONSTRAINT FK_ADtb_banner__tb_Indirizzario; " + _
            " ALTER TABLE ADtb_banner_posizioni DROP CONSTRAINT FK_ADtb_banner_posizioni__ADtb_banner; " + _
            " ALTER TABLE ADtb_banner_posizioni DROP CONSTRAINT FK_ADtb_banner_posizioni__tb_contents_index; " + _
            " ALTER TABLE ADtb_storico_impress DROP CONSTRAINT FK_ADtb_storico_impress__ADtb_banner_posizioni; " + _
            " ALTER TABLE ADlog_impress DROP CONSTRAINT FK_ADlog_impress__ADtb_banner_posizioni; " + _
            " ALTER TABLE ADlog_click DROP CONSTRAINT FK_ADlog_click__ADtb_banner_posizioni; " + _
            " DROP TABLE ADtb_banner;" + _
            " DROP TABLE ADtb_banner_posizioni; " + _
            " CREATE TABLE " + SQL_Dbo(Conn) + "ADtb_contratti_banner ( " + _
            "   cb_id " + SQL_PrimaryKey(conn, "ADtb_contratti_banner") + ", " + vbCrLF + _
            "	cb_tipo_id int NOT NULL, " + vbCrLf + _
            "   cb_cliente_id int NOT NULL, " + vbCrLf + _
            "   cb_riferimento_interno " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLF + _
            SQL_MultiLanguageField("   cb_banner_link_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
            "   cb_banner_file " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLF + _
            SQL_MultiLanguageField("   cb_banner_title_<lingua>" + SQL_CharField(Conn, 255) + " NULL ") + ", " + _
            SQL_MultiLanguageField("   cb_banner_text_<lingua>" + SQL_CharField(Conn, 0) + " NULL ") + ", " + _
            "   cb_banner_site " + SQL_CharField(Conn, 255) + " NULL, " + vbCrLF + _
            "   cb_data_inizio smalldatetime NULL, " + vbCrLf + _
            "   cb_data_fine smalldatetime NULL, " + vbCrLf + _
            "   cb_impression_iniziali int NULL, " + vbCrLf + _
            "   cb_impression_attuali int NULL, " + vbCrLf + _
            "   cb_click_iniziali int NULL, " + vbCrLf + _
            "   cb_click_attuali int NULL, " + vbCrLf + _
            "   CONSTRAINT FK_ADtb_contratti_banner__ADtb_tipiBanner " + vbCrLF + _
            "       FOREIGN KEY(cb_tipo_id) REFERENCES ADtb_tipiBanner(tipoB_id) " + vbCrLf + _
            "       ON UPDATE CASCADE ON DELETE CASCADE, " + vbCrLf + _
            "   CONSTRAINT FK_ADtb_contratti_banner__tb_Indirizzario " + vbCRLF + _
            "       FOREIGN KEY(cb_cliente_id) REFERENCES tb_Indirizzario(IDElencoIndirizzi) " + vbCrLf + _
            "       ON UPDATE CASCADE ON DELETE CASCADE " + vbCrLf + _
            " ) ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 2
'...........................................................................................
'aggiunge riferimenti di inserimento e modifica ai contratti
'...........................................................................................
function Aggiornamento__NEXTBANNER2__2(conn)
    Aggiornamento__NEXTBANNER2__2 = _
            " ALTER TABLE ADtb_contratti_banner ADD " + _
            "   cb_insData	SMALLDATETIME NULL, " + _
            "   cb_insAdmin_id INT NULL, " + _
            "   cb_modData	SMALLDATETIME NULL, " + _
            "   cb_modAdmin_id INT NULL " + _
            " ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 3
'...........................................................................................
'aggiunge note e flag per disattivazione ai contratti
'...........................................................................................
function Aggiornamento__NEXTBANNER2__3(conn)
    Aggiornamento__NEXTBANNER2__3 = _
            " ALTER TABLE ADtb_contratti_banner ADD " + _
            "   cb_attivo BIT NULL, " + _
            "   cb_note " + SQL_CharField(Conn, 0) + " NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 4
'...........................................................................................
'aggiunge vista sui banner
'...........................................................................................
function Aggiornamento__NEXTBANNER2__4(conn)
    Aggiornamento__NEXTBANNER2__4 = _
            " CREATE VIEW " + SQL_Dbo(Conn) + "ADv_contratti_banner AS " + vbCrLF + _
            "   SELECT * FROM (ADtb_contratti_banner " + vbCrLF + _
            "                  INNER JOIN tb_indirizzario ON ADtb_contratti_banner.cb_cliente_id = tb_indirizzario.IdElencoIndirizzi) " + vbCrLF + _
            "                  INNER JOIN ADtb_tipiBanner ON ADtb_contratti_banner.cb_tipo_id = ADtb_tipiBanner.tipoB_id " + vbCrLf + _
            " ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 5
'...........................................................................................
'aggiunge tabella per registrazione pubblicazioni banner
'...........................................................................................
function Aggiornamento__NEXTBANNER2__5(conn)
    Aggiornamento__NEXTBANNER2__5 = _
            " CREATE TABLE " + SQL_Dbo(Conn) + "ADtb_contratti_banner_pubblicazioni ( " + _
            "   pub_id " + SQL_PrimaryKey(conn, "ADtb_contratti_banner_pubblicazioni") + ", " + _
            "   pub_cb_id INT NOT NULL, " + _
            "   pub_index_id INT NOT NULL, " + _
            "   pub_su_Ramo BIT NOT NULL, " + _
            "   pub_attiva BIT NOT NULL, " + _
            "   pub_impression INT NULL, " + _
            "   pub_click INT NULL, " + _
            "   pub_insData	SMALLDATETIME NULL, " + _
            "   pub_insAdmin_id INT NULL, " + _
            "   pub_modData	SMALLDATETIME NULL, " + _
            "   pub_modAdmin_id INT NULL, " + _
            "   CONSTRAINT FK_ADtb_contratti_banner_pubblicazioni__ADtb_contratti_banner " + _
            "       FOREIGN KEY(pub_cb_id) REFERENCES ADtb_contratti_banner (cb_id) " + vbCrLf + _
            "       ON UPDATE CASCADE ON DELETE CASCADE, " + vbCrLF + _
            "   CONSTRAINT FK_ADtb_contratti_banner_pubblicazioni__tb_contents_index " + vbCrLF + _
            "       FOREIGN KEY(pub_index_id) REFERENCES tb_contents_index (idx_id) " + vbCrLf + _
            "       ON UPDATE CASCADE ON DELETE CASCADE " + vbCrLf + _
			" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 6
'...........................................................................................
'corregge storico click ed impression
'...........................................................................................
function Aggiornamento__NEXTBANNER2__6(conn)
    Aggiornamento__NEXTBANNER2__6 = _
            " ALTER TABLE ADlog_impress DROP COLUMN log_posizione_id; " + _
            " ALTER TABLE ADlog_click DROP COLUMN log_posizione_id; " + _
            " ALTER TABLE ADlog_impress ADD " + _
            "   log_pub_id INT NULL, " + _
            "   log_index_id INT NULL, " + _
            "   CONSTRAINT FK_ADlog_impress__ADtb_contratti_banner_pubblicazioni " + _
            "       FOREIGN KEY(log_pub_id) REFERENCES ADtb_contratti_banner_pubblicazioni(pub_id) " + _
            "       ON UPDATE CASCADE ON DELETE CASCADE, " + _
            "   CONSTRAINT FK_ADlog_impress__tb_contents_index " + _
            "       FOREIGN KEY(log_index_id) REFERENCES tb_contents_index(idx_id) " + _
            "       ON UPDATE NO ACTION ON DELETE NO ACTION; " + _
            " ALTER TABLE ADlog_impress ALTER COLUMN log_pub_id INT NOT NULL; "+ _
            " ALTER TABLE ADlog_impress NOCHECK CONSTRAINT FK_ADlog_impress__tb_contents_index; " + _
            " ALTER TABLE ADlog_click ADD " + _
            "   log_pub_id INT NULL, " + _
            "   log_index_id INT NULL, " + _
            "   CONSTRAINT FK_ADlog_click__ADtb_contratti_banner_pubblicazioni " + _
            "       FOREIGN KEY(log_pub_id) REFERENCES ADtb_contratti_banner_pubblicazioni(pub_id) " + _
            "       ON UPDATE CASCADE ON DELETE CASCADE, " + _
            "   CONSTRAINT FK_ADlog_click__tb_contents_index " + _
            "       FOREIGN KEY(log_index_id) REFERENCES tb_contents_index(idx_id) " + _
            "       ON UPDATE NO ACTION ON DELETE NO ACTION; " + _
            " ALTER TABLE ADlog_click ALTER COLUMN log_pub_id INT NOT NULL; "+ _
            " ALTER TABLE ADlog_click NOCHECK CONSTRAINT FK_ADlog_click__tb_contents_index; " + _
            " ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 7
'...........................................................................................
'corregge storico click ed impression
'...........................................................................................
function Aggiornamento__NEXTBANNER2__7(conn)
    Aggiornamento__NEXTBANNER2__7 = _
            " DELETE FROM adlog_click" + vbCrLF + _
			" WHERE EXISTS (SELECT log_ip FROM adlog_click l" + vbCrLF + _
			" 				WHERE l.log_ip = adlog_click.log_ip" + vbCrLF + _
			" 			    GROUP BY log_ip" + vbCrLF + _
			"			  	HAVING COUNT(*) >= 30);" + vbCrLF + _
			" UPDATE adtb_contratti_banner_pubblicazioni" + vbCrLF + _
			" SET pub_click = (SELECT COUNT(*) FROM adlog_click WHERE log_pub_id = pub_id)" + vbCrLF + _
			" UPDATE adtb_contratti_banner" + vbCrLF + _
			" SET cb_click_attuali = (SELECT COUNT(*) FROM adlog_click l" + vbCrLF + _
			" 						  INNER JOIN adtb_contratti_banner_pubblicazioni p ON (l.log_pub_id = p.pub_id)" + vbCrLF + _
			"						  WHERE pub_cb_id = cb_id) ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 8
'...........................................................................................
'crea tabella per filtri log
'...........................................................................................
function Aggiornamento__NEXTBANNER2__8(conn)
    Aggiornamento__NEXTBANNER2__8 = _
            " CREATE TABLE " + SQL_Dbo(Conn) + "ADtb_filtri ( " + _
            "   fil_id " + SQL_PrimaryKey(conn, "ADtb_filtri") + ", " + _
            "   fil_parametro " + SQL_CharField(Conn, 50) + " NOT NULL, " + _
            "   fil_valore " + SQL_CharField(Conn, 255) + " NULL," + _
			"	fil_tipo INT NULL" + _
			" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER 9
'...........................................................................................
'cancella tabella per filtri log
'...........................................................................................
function Aggiornamento__NEXTBANNER2__9(conn)
    Aggiornamento__NEXTBANNER2__9 = _
			DropObject(conn, "ADtb_filtri", "TABLE")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-BANNER  10
'...........................................................................................
' rimuove integrita' referenziale banner/posizione su indice ed aggiunge campo per mantenere
' il nodo dell'indice anche in caso di cancellazione.
'...........................................................................................
function Aggiornamento__NEXTBANNER2__10(conn)
    Aggiornamento__NEXTBANNER2__10 = _
            " ALTER TABLE ADtb_contratti_banner_pubblicazioni DROP CONSTRAINT FK_ADtb_contratti_banner_pubblicazioni__tb_contents_index; " + _
			SQL_AddForeignKey(conn, "ADtb_contratti_banner_pubblicazioni", "pub_index_id", "tb_contents_index", "idx_id", false, "") + _
			" ALTER TABLE ADtb_contratti_banner_pubblicazioni ADD " + _
			"	pub_index_name " + SQL_CharField(Conn, 3000) + " NULL " + _
			" ; " + _
			" UPDATE ADtb_contratti_banner_pubblicazioni SET " + _
			"	pub_index_name = ( SELECT co_titolo_IT FROM v_indice WHERE idx_id = ADtb_contratti_banner_pubblicazioni.pub_index_id ) ;"
end function
'*******************************************************************************************


%>