<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-memo
'...........................................................................................
'...........................................................................................

'*******************************************************************************************
'INSTALLAZIONE MEMO
'...........................................................................................
function Install__MEMO(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__MEMO = _
				"CREATE TABLE tb_categorieCircolari (" + _
				"catC_id COUNTER CONSTRAINT PK_tb_categorieCircolari PRIMARY KEY ," + _
				"catC_nome_it TEXT(255) WITH COMPRESSION NULL ," + _
				"catC_nome_en TEXT(255) WITH COMPRESSION NULL ," + _
				"catC_codice TEXT(50) WITH COMPRESSION NULL ," + _
				"catC_foglia BIT NULL ," + _
				"catC_livello INTEGER NULL ," + _
				"catC_padre_id INTEGER NULL , " + _
				"catC_ordine INTEGER NULL , " + _
				"catC_ordine_assoluto TEXT(250) WITH COMPRESSION NULL ," + _
				"catC_descr_it NTEXT NULL , " + _
				"catC_descr_en NTEXT NULL , " + _
				"catC_tipologia_padre_base INTEGER NULL, " + _
				"catC_visibile BIT NULL ," + _
				"catC_albero_visibile BIT NULL , " + _
				"catC_tipologie_padre_lista TEXT(255) WITH COMPRESSION NULL " + _
				"); " + _
				"CREATE TABLE tb_Circolari (" + _
				"CI_id COUNTER CONSTRAINT PK_tb_Circolari PRIMARY KEY ," + _
				"CI_Numero TEXT(50) WITH COMPRESSION NULL ," + _
				"CI_Titolo TEXT(250) WITH COMPRESSION NULL ," + _
				"CI_Estratto NTEXT NULL , " + _
				"CI_Pubblicazione DATETIME NULL ," + _
				"CI_Scadenza DATETIME NULL ," + _
				"CI_File NTEXT NULL , " + _
				"CI_Visibile BIT NULL ," + _
				"CI_Protetto BIT NULL ," + _
				"CI_idcategoria INTEGER NULL " + _
				"); " + _
				"ALTER TABLE tb_Circolari ADD CONSTRAINT FK_tb_Circolari_tb_categorieCircolari " + _
				"FOREIGN KEY (CI_idcategoria) REFERENCES tb_categorieCircolari (catC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"
 
				
		case DB_SQL
			Install__MEMO = _
				"CREATE TABLE dbo.tb_categorieCircolari (" + _
				"catC_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_tb_categorieCircolari PRIMARY KEY  CLUSTERED ," + _
				"catC_nome_it NVARCHAR(255) NULL , " + _
				"catC_nome_en NVARCHAR(255) NULL ," + _
				"catC_codice NVARCHAR(50) NULL ," + _
				"catC_foglia BIT NULL ," + _
				"catC_livello INTEGER NULL ," + _
				"catC_padre_id INTEGER NULL ," + _
				"catC_ordine INTEGER NULL ," + _
				"catC_ordine_assoluto NVARCHAR(250) NULL ," + _				
				"catC_descr_it NTEXT NULL ," + _
				"catC_descr_en NTEXT NULL , " + _
				"catC_tipologia_padre_base INTEGER NULL ," + _
				"catC_visibile BIT NULL , " + _
				"catC_albero_visibile BIT NULL , " + _
				"catC_tipologie_padre_lista NVARCHAR(255) NULL , " + _
				"); " + _
				"CREATE TABLE dbo.tb_Circolari (" + _
				"CI_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_tb_Circolari PRIMARY KEY  CLUSTERED ," + _
				"CI_Numero NVARCHAR(50) NULL ," + _
				"CI_Titolo NVARCHAR(250) NULL ," + _	
				"CI_Estratto NTEXT NULL , " + _
				"CI_Pubblicazione SMALLDATETIME NULL ," + _
				"CI_Scadenza SMALLDATETIME NULL ," + _
				"CI_File NTEXT NULL , " + _
				"CI_Visibile BIT NULL ," + _
				"CI_Protetto BIT NULL ," + _
				"CI_idcategoria INTEGER NULL " + _
				"); " + _	
				"ALTER TABLE tb_Circolari ADD CONSTRAINT FK_tb_Circolari_tb_categorieCircolari " + _
				"FOREIGN KEY (CI_idcategoria) REFERENCES tb_categorieCircolari (catC_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"
				
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO MEMO 2
'...........................................................................................
'
'...........................................................................................
function Install__MEMO_2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__MEMO_2 = _
				"ALTER TABLE tb_categorieCircolari ADD "+ _
					"catC_nome_fr TEXT(255) WITH COMPRESSION NULL , "+ _
					"catC_nome_de TEXT(255) WITH COMPRESSION NULL , "+ _
					"catC_nome_es TEXT(255) WITH COMPRESSION NULL , "+ _
					"catC_descr_fr NTEXT NULL , " + _
					"catC_descr_de NTEXT NULL , " + _
					"catC_descr_es NTEXT NULL ; "
					
		case DB_SQL
			Install__MEMO_2 = _
				"ALTER TABLE tb_categorieCircolari ADD "+ _
					"catC_nome_fr NVARCHAR(255) NULL , "+ _
					"catC_nome_de NVARCHAR(255) NULL , "+ _
					"catC_nome_es NVARCHAR(255) NULL , "+ _
					"catC_descr_fr NTEXT NULL ," + _
					"catC_descr_de NTEXT NULL ," + _
					"catC_descr_es NTEXT NULL;"
	end select
end function
'*******************************************************************************************											


'*******************************************************************************************
'AGGIORNAMENTO MEMO 3
'...........................................................................................
'
'...........................................................................................
function Install__MEMO_3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__MEMO_3 = _
				"CREATE TABLE log_circolari (" + _
				"log_id COUNTER CONSTRAINT PK_log_circolari PRIMARY KEY ," + _
				"log_ut_id INTEGER NULL ," + _
				"log_dip_id INTEGER NULL ," + _
				"log_ci_id INTEGER NULL ," + _
				"log_data DATETIME NULL ," + _
				"); " + _
				"ALTER TABLE log_circolari ADD CONSTRAINT FK_log_circolari_tb_categorieCircolari " + _
				"FOREIGN KEY (log_ci_id) REFERENCES tb_Circolari (CI_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"
				
		case DB_SQL
			Install__MEMO_3 = _
				"CREATE TABLE dbo.log_circolari (" + _
				"log_id int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_log_circolari PRIMARY KEY  CLUSTERED ," + _
				"log_ut_id INTEGER NULL ," + _
				"log_dip_id INTEGER NULL ," + _
				"log_ci_id INTEGER NULL ," + _
				"log_data SMALLDATETIME NULL ," + _
				"); " + _	
				"ALTER TABLE log_circolari ADD CONSTRAINT FK_log_circolari_tb_categorieCircolari " + _
				"FOREIGN KEY (log_ci_id) REFERENCES tb_Circolari (CI_id) " + _
				"ON UPDATE CASCADE ON DELETE CASCADE;"
				
	end select
end function
'*******************************************************************************************	

				
'*******************************************************************************************
'AGGIORNAMENTO SPECIALE NEXT-MEMO  1
'...........................................................................................
'ClassCategorie: aggiunge il campo per la gestione della lista degli IDs dei padri
'...........................................................................................
function AggiornamentoSpeciale__MEMO__1(DB, rs, version)
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "tb_categorieCircolari", "catC")
    AggiornamentoSpeciale__MEMO__1 = "SELECT * FROM AA_versione"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 1
'...........................................................................................
'
'...........................................................................................
function Aggiornamento__MEMO__1(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__MEMO__1 = _
				"ALTER TABLE tb_categorieCircolari ADD "+ _
					"catC_nome_ru TEXT(255) WITH COMPRESSION NULL , "+ _
					"catC_nome_cn TEXT(255) WITH COMPRESSION NULL , "+ _
					"catC_descr_ru NTEXT NULL , " + _
					"catC_descr_cn NTEXT NULL;" 	
		case DB_SQL
			Aggiornamento__MEMO__1 = _
				"ALTER TABLE tb_categorieCircolari ADD "+ _
					"catC_nome_ru NVARCHAR(255) NULL , "+ _
					"catC_nome_cn NVARCHAR(255) NULL , "+ _
					"catC_descr_ru NTEXT NULL ," + _
					"catC_descr_cn NTEXT NULL;"
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-MEMO 2
'...........................................................................................
'
'...........................................................................................
function Aggiornamento__MEMO__2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__MEMO__2 = _
				"ALTER TABLE tb_categorieCircolari ADD "+ _
					"catC_nome_pt TEXT(255) WITH COMPRESSION NULL , "+ _
					"catC_descr_pt NTEXT NULL;" 	
		case DB_SQL
			Aggiornamento__MEMO__2 = _
				"ALTER TABLE tb_categorieCircolari ADD "+ _
					"catC_nome_pt NVARCHAR(255) NULL , "+ _
					"catC_descr_pt NTEXT NULL;"
	end select
end function
'*******************************************************************************************

%>