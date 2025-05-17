<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-info
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'INSTALLAZIONE NEXT-INFO
'...........................................................................................
function Install__NEXTINFO(conn)
	Select case DB_Type(conn)
		case DB_Access
			Install__NEXTINFO = _
				""
		case DB_SQL
			Install__NEXTINFO = _
				"CREATE TABLE dbo.itb_anagrafiche_descrRag ( " + vbCrLf + _
				"	adr_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
				"	adr_titolo_it nvarchar (250) NULL , " + vbCrLF + _
				"	adr_titolo_en nvarchar (250) NULL , " + vbCRLF + _
				"	adr_titolo_fr nvarchar (250) NULL , " + vbCrLF + _
				"	adr_titolo_es nvarchar (250) NULL , " + vbCrLf + _
				"	adr_titolo_de nvarchar (250) NULL , " + vbCrLf + _
				"	adr_ordine int NULL , " + vbCrLf + _
				"	adr_external_id nvarchar (250) NULL , " + vbCrLF + _
				"	adr_external_source nvarchar (250) NULL " + vbCrLf + _
				" ) ; " + _
				"CREATE TABLE dbo.itb_anagrafiche_tipi (" + vbCrLf + _
				"	ant_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	ant_nome_it nvarchar (250) NULL ," + vbCrLf + _
				"	ant_nome_en nvarchar (250) NULL ," + vbCrLf + _
				"	ant_nome_fr nvarchar (250) NULL ," + vbCrLf + _
				"	ant_nome_es nvarchar (250) NULL ," + vbCrLf + _
				"	ant_nome_de nvarchar (250) NULL ," + vbCrLf + _
				"	ant_foto nvarchar (255) NULL ," + vbCrLf + _
				"	ant_codice nvarchar (50) NULL ," + vbCrLf + _
				"	ant_descr_it ntext NULL ," + vbCrLf + _
				"	ant_descr_en ntext NULL ," + vbCrLf + _
				"	ant_descr_fr ntext NULL ," + vbCrLf + _
				"	ant_descr_es ntext NULL ," + vbCrLf + _
				"	ant_descr_de ntext NULL ," + vbCrLf + _
				"	ant_foglia bit NULL ," + vbCrLf + _
				"	ant_livello int NULL ," + vbCrLf + _
				"	ant_padre_id int NULL ," + vbCrLf + _
				"	ant_ordine int NULL ," + vbCrLf + _
				"	ant_ordine_assoluto nvarchar (250) NULL ," + vbCrLf + _
				"	ant_external_id nvarchar (250) NULL ," + vbCrLf + _
				"	ant_external_source nvarchar (250) NULL ," + vbCrLf + _
				"	ant_tipologia_padre_base int NULL ," + vbCrLf + _
				"	ant_visibile bit NULL ," + vbCrLf + _
				"	ant_albero_visibile bit NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_aree (" + vbCrLf + _
				"	are_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	are_nome_it nvarchar (250) NULL ," + vbCrLf + _
				"	are_nome_en nvarchar (250) NULL ," + vbCrLf + _
				"	are_nome_fr nvarchar (250) NULL ," + vbCrLf + _
				"	are_nome_es nvarchar (250) NULL ," + vbCrLf + _
				"	are_nome_de nvarchar (250) NULL ," + vbCrLf + _
				"	are_foto nvarchar (255) NULL ," + vbCrLf + _
				"	are_codice nvarchar (50) NULL ," + vbCrLf + _
				"	are_descr_it ntext NULL ," + vbCrLf + _
				"	are_descr_en ntext NULL ," + vbCrLf + _
				"	are_descr_fr ntext NULL ," + vbCrLf + _
				"	are_descr_es ntext NULL ," + vbCrLf + _
				"	are_descr_de ntext NULL ," + vbCrLf + _
				"	are_foglia bit NULL ," + vbCrLf + _
				"	are_livello int NULL ," + vbCrLf + _
				"	are_padre_id int NULL ," + vbCrLf + _
				"	are_ordine int NULL ," + vbCrLf + _
				"	are_ordine_assoluto nvarchar (250) NULL ," + vbCrLf + _
				"	are_external_id nvarchar (50) NULL ," + vbCrLf + _
				"	are_tipologia_padre_base int NULL ," + vbCrLf + _
				"	are_visibile bit NULL ," + vbCrLf + _
				"	are_albero_visibile bit NULL ," + vbCrLf + _
				"	are_external_source nvarchar (250) NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_eventi_categorie (" + vbCrLf + _
				"	evc_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	evc_nome_it nvarchar (250) NULL ," + vbCrLf + _
				"	evc_nome_en nvarchar (250) NULL ," + vbCrLf + _
				"	evc_nome_fr nvarchar (250) NULL ," + vbCrLf + _
				"	evc_nome_es nvarchar (250) NULL ," + vbCrLf + _
				"	evc_nome_de nvarchar (250) NULL ," + vbCrLf + _
				"	evc_foto nvarchar (255) NULL ," + vbCrLf + _
				"	evc_codice nvarchar (50) NULL ," + vbCrLf + _
				"	evc_descr_it ntext NULL ," + vbCrLf + _
				"	evc_descr_en ntext NULL ," + vbCrLf + _
				"	evc_descr_fr ntext NULL ," + vbCrLf + _
				"	evc_descr_es ntext NULL ," + vbCrLf + _
				"	evc_descr_de ntext NULL ," + vbCrLf + _
				"	evc_foglia bit NULL ," + vbCrLf + _
				"	evc_livello int NULL ," + vbCrLf + _
				"	evc_padre_id int NULL ," + vbCrLf + _
				"	evc_ordine int NULL ," + vbCrLf + _
				"	evc_ordine_assoluto nvarchar (250) NULL ," + vbCrLf + _
				"	evc_external_id nvarchar (50) NULL ," + vbCrLf + _
				"	evc_tipologia_padre_base int NULL ," + vbCrLf + _
				"	evc_visibile bit NULL ," + vbCrLf + _
				"	evc_albero_visibile bit NULL ," + vbCrLf + _
				"	evc_external_source nvarchar (250) NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_eventi_descrittori (" + vbCrLf + _
				"	evd_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	evd_nome_it nvarchar (255) NULL ," + vbCrLf + _
				"	evd_nome_en nvarchar (255) NULL ," + vbCrLf + _
				"	evd_nome_fr nvarchar (255) NULL ," + vbCrLf + _
				"	evd_nome_es nvarchar (255) NULL ," + vbCrLf + _
				"	evd_nome_de nvarchar (255) NULL ," + vbCrLf + _
				"	evd_unita_it nvarchar (50) NULL ," + vbCrLf + _
				"	evd_unita_en nvarchar (50) NULL ," + vbCrLf + _
				"	evd_unita_fr nvarchar (50) NULL ," + vbCrLf + _
				"	evd_unita_es nvarchar (50) NULL ," + vbCrLf + _
				"	evd_unita_de nvarchar (50) NULL ," + vbCrLf + _
				"	evd_tipo int NULL ," + vbCrLf + _
				"	evd_principale bit NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_eventi_tipologie (" + vbCrLf + _
				"	evt_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	evt_nome_it nvarchar (250) NULL ," + vbCrLf + _
				"	evt_nome_en nvarchar (250) NULL ," + vbCrLf + _
				"	evt_nome_fr nvarchar (250) NULL ," + vbCrLf + _
				"	evt_nome_es nvarchar (250) NULL ," + vbCrLf + _
				"	evt_nome_de nvarchar (250) NULL ," + vbCrLf + _
				"	evt_visibile bit NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.ilog_admin (" + vbCrLf + _
				"	ilog_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	ilog_admin_id int NULL ," + vbCrLf + _
				"	ilog_data datetime NULL ," + vbCrLf + _
				"	ilog_tabella nvarchar (100) NULL ," + vbCrLf + _
				"	ilog_chiave_id int NULL ," + vbCrLf + _
				"	ilog_descr nvarchar (250) NULL ," + vbCrLf + _
				"	ilog_azione_cod int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_admin (" + vbCrLf + _
				"	adm_id int NOT NULL ," + vbCrLf + _
				"	adm_area_id int NULL ," + vbCrLf + _
				"	adm_permesso int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_evCategorie_descrittori (" + vbCrLf + _
				"	rcd_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rcd_categoria_id int NULL ," + vbCrLf + _
				"	rcd_descrittore_id int NULL ," + vbCrLf + _
				"	rcd_ordine int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_anagrafiche_descrittori (" + vbCrLf + _
				"	and_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	and_raggruppamento_id int NULL ," + vbCrLf + _
				"	and_nome_it nvarchar (255) NULL ," + vbCrLf + _
				"	and_nome_en nvarchar (255) NULL ," + vbCrLf + _
				"	and_nome_fr nvarchar (255) NULL ," + vbCrLf + _
				"	and_nome_es nvarchar (255) NULL ," + vbCrLf + _
				"	and_nome_de nvarchar (255) NULL ," + vbCrLf + _
				"	and_unita_it nvarchar (50) NULL ," + vbCrLf + _
				"	and_unita_en nvarchar (50) NULL ," + vbCrLf + _
				"	and_unita_fr nvarchar (50) NULL ," + vbCrLf + _
				"	and_unita_es nvarchar (50) NULL ," + vbCrLf + _
				"	and_unita_de nvarchar (50) NULL ," + vbCrLf + _
				"	and_tipo int NULL ," + vbCrLf + _
				"	and_principale bit NULL ," + vbCrLf + _
				"	and_img nvarchar (250) NULL ," + vbCrLf + _
				"	and_external_id nvarchar (250) NULL ," + vbCrLf + _
				"	and_external_source nvarchar (250) NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_eventi (" + vbCrLf + _
				"	eve_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	eve_categoria_id int NULL ," + vbCrLf + _
				"	eve_alt_categoria_id int NULL ," + vbCrLf + _
				"	eve_tipologia_id int NULL ," + vbCrLf + _
				"	eve_titolo_it nvarchar (250) NULL ," + vbCrLf + _
				"	eve_titolo_en nvarchar (250) NULL ," + vbCrLf + _
				"	eve_titolo_fr nvarchar (250) NULL ," + vbCrLf + _
				"	eve_titolo_es nvarchar (250) NULL ," + vbCrLf + _
				"	eve_titolo_de nvarchar (250) NULL ," + vbCrLf + _
				"	eve_descr_it ntext NULL ," + vbCrLf + _
				"	eve_descr_en ntext NULL ," + vbCrLf + _
				"	eve_descr_fr ntext NULL ," + vbCrLf + _
				"	eve_descr_es ntext NULL ," + vbCrLf + _
				"	eve_descr_de ntext NULL ," + vbCrLf + _
				"	eve_ingresso_intero nvarchar (100) NULL ," + vbCrLf + _
				"	eve_ingresso_ridotto nvarchar (100) NULL ," + vbCrLf + _
				"	eve_ridotto_it ntext NULL ," + vbCrLf + _
				"	eve_ridotto_en ntext NULL ," + vbCrLf + _
				"	eve_ridotto_fr ntext NULL ," + vbCrLf + _
				"	eve_ridotto_es ntext NULL ," + vbCrLf + _
				"	eve_ridotto_de ntext NULL ," + vbCrLf + _
				"	eve_info_it ntext NULL ," + vbCrLf + _
				"	eve_info_en ntext NULL ," + vbCrLf + _
				"	eve_info_fr ntext NULL ," + vbCrLf + _
				"	eve_info_es ntext NULL ," + vbCrLf + _
				"	eve_info_de ntext NULL ," + vbCrLf + _
				"	eve_codice nvarchar (250) NULL ," + vbCrLf + _
				"	eve_telefono nvarchar (250) NULL ," + vbCrLf + _
				"	eve_insData datetime NULL ," + vbCrLf + _
				"	eve_insAdmin_id int NULL ," + vbCrLf + _
				"	eve_modData datetime NULL ," + vbCrLf + _
				"	eve_modAdmin_id int NULL ," + vbCrLf + _
				"	eve_visibile bit NULL ," + vbCrLf + _
				"	eve_censurato bit NULL ," + vbCrLf + _
				"	eve_ranking int NULL ," + vbCrLf + _
				"	eve_pubblData datetime NULL ," + vbCrLf + _
				"	eve_censurato_perche ntext NULL ," + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_anTipi_descrittori (" + vbCrLf + _
				"	rtd_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rtd_tipologia_id int NULL ," + vbCrLf + _
				"	rtd_descrittore_id int NULL ," + vbCrLf + _
				"	rtd_ordine int NULL ," + vbCrLf + _
				"	rtd_locked bit NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_eventi_collegati (" + vbCrLf + _
				"	rec_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rec_evento_id int NULL ," + vbCrLf + _
				"	rec_eventoCol_id int NULL ," + vbCrLf + _
				"	rec_paginaSito_id int NULL ," + vbCrLf + _
				"	rec_link_it nvarchar (250) NULL ," + vbCrLf + _
				"	rec_link_en nvarchar (250) NULL ," + vbCrLf + _
				"	rec_link_fr nvarchar (250) NULL ," + vbCrLf + _
				"	rec_link_es nvarchar (250) NULL ," + vbCrLf + _
				"	rec_link_de nvarchar (250) NULL ," + vbCrLf + _
				"	rec_nome_it nvarchar (250) NULL ," + vbCrLf + _
				"	rec_nome_en nvarchar (250) NULL ," + vbCrLf + _
				"	rec_nome_fr nvarchar (250) NULL ," + vbCrLf + _
				"	rec_nome_es nvarchar (250) NULL ," + vbCrLf + _
				"	rec_nome_de nvarchar (250) NULL ," + vbCrLf + _
				"	rec_descr_it ntext NULL ," + vbCrLf + _
				"	rec_descr_en ntext NULL ," + vbCrLf + _
				"	rec_descr_fr ntext NULL ," + vbCrLf + _
				"	rec_descr_es ntext NULL ," + vbCrLf + _
				"	rec_descr_de ntext NULL ," + vbCrLf + _
				"	rec_ordine int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_eventi_descrCat (" + vbCrLf + _
				"	red_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	red_evento_id int NULL ," + vbCrLf + _
				"	red_descrittore_id int NULL ," + vbCrLf + _
				"	red_valore_it nvarchar (250) NULL ," + vbCrLf + _
				"	red_valore_en nvarchar (250) NULL ," + vbCrLf + _
				"	red_valore_fr nvarchar (250) NULL ," + vbCrLf + _
				"	red_valore_es nvarchar (250) NULL ," + vbCrLf + _
				"	red_valore_de nvarchar (250) NULL ," + vbCrLf + _
				"	red_memo_it ntext NULL ," + vbCrLf + _
				"	red_memo_en ntext NULL ," + vbCrLf + _
				"	red_memo_fr ntext NULL ," + vbCrLf + _
				"	red_memo_es ntext NULL ," + vbCrLf + _
				"	red_memo_de ntext NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_eventi_img (" + vbCrLf + _
				"	evi_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	evi_didascalia_it ntext NULL ," + vbCrLf + _
				"	evi_didascalia_en ntext NULL ," + vbCrLf + _
				"	evi_didascalia_fr ntext NULL ," + vbCrLf + _
				"	evi_didascalia_es ntext NULL ," + vbCrLf + _
				"	evi_didascalia_de ntext NULL ," + vbCrLf + _
				"	evi_pubblicazione datetime NULL ," + vbCrLf + _
				"	evi_visibile bit NULL ," + vbCrLf + _
				"	evi_numero int NULL ," + vbCrLf + _
				"	evi_thumb nvarchar (250) NULL ," + vbCrLf + _
				"	evi_zoom nvarchar (250) NULL ," + vbCrLf + _
				"	evi_ordine int NULL ," + vbCrLf + _
				"	evi_evento_id int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.itb_anagrafiche (" + vbCrLf + _
				"	ana_id int NOT NULL ," + vbCrLf + _
				"   ana_codice nvarchar (250) NULL, " + vbCrLf + _
				"	ana_tipo_id int NULL ," + vbCrLf + _
				"	ana_alt_tipo_id int NULL ," + vbCrLf + _
				"	ana_area_id int NULL ," + vbCrLf + _
				"	ana_insData datetime NULL ," + vbCrLf + _
				"	ana_insAdmin_id int NULL ," + vbCrLf + _
				"	ana_modData datetime NULL ," + vbCrLf + _
				"	ana_modAdmin_id int NULL ," + vbCrLf + _
				"	ana_descr_it ntext NULL ," + vbCrLf + _
				"	ana_descr_en ntext NULL ," + vbCrLf + _
				"	ana_descr_fr ntext NULL ," + vbCrLf + _
				"	ana_descr_es ntext NULL ," + vbCrLf + _
				"	ana_descr_de ntext NULL ," + vbCrLf + _
				"	ana_visibile bit NULL ," + vbCrLf + _
				"	ana_link_attivi bit NULL ," + vbCrLf + _
				"	ana_censurato bit NULL ," + vbCrLf + _
				"	ana_ranking int NULL ," + vbCrLf + _
				"	ana_censurato_perche ntext NULL ," + vbCrLf + _
				"	ana_classificazione nvarchar (50) NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_anagrafiche_descrTipi (" + vbCrLf + _
				"	rad_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rad_anagrafica_id int NULL ," + vbCrLf + _
				"	rad_descrittore_id int NULL ," + vbCrLf + _
				"	rad_valore_it nvarchar (255) NULL ," + vbCrLf + _
				"	rad_valore_en nvarchar (255) NULL ," + vbCrLf + _
				"	rad_valore_fr nvarchar (255) NULL ," + vbCrLf + _
				"	rad_valore_es nvarchar (255) NULL ," + vbCrLf + _
				"	rad_valore_de nvarchar (255) NULL ," + vbCrLf + _
				"	rad_memo_it ntext NULL ," + vbCrLf + _
				"	rad_memo_en ntext NULL ," + vbCrLf + _
				"	rad_memo_fr ntext NULL ," + vbCrLf + _
				"	rad_memo_es ntext NULL ," + vbCrLf + _
				"	rad_memo_de ntext NULL ," + vbCrLf + _ 
                "   rad_scadenza SMALLDATETIME NULL " + vbCrLf + _ 
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_anagrafiche_img (" + vbCrLf + _
				"	ani_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	ani_didascalia_it ntext NULL ," + vbCrLf + _
				"	ani_didascalia_en ntext NULL ," + vbCrLf + _
				"	ani_didascalia_fr ntext NULL ," + vbCrLf + _
				"	ani_didascalia_es ntext NULL ," + vbCrLf + _
				"	ani_didascalia_de ntext NULL ," + vbCrLf + _
				"	ani_pubblicazione datetime NULL ," + vbCrLf + _
				"	ani_visibile bit NULL ," + vbCrLf + _
				"	ani_numero int NULL ," + vbCrLf + _
				"	ani_thumb nvarchar (250) NULL ," + vbCrLf + _
				"	ani_zoom nvarchar (250) NULL ," + vbCrLf + _
				"	ani_ordine int NULL ," + vbCrLf + _
				"	ani_anagrafica_id int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_luoghi (" + vbCrLf + _
				"	rlu_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rlu_evento_id int NULL ," + vbCrLf + _
				"	rlu_anagrafica_id int NULL ," + vbCrLf + _
				"	rlu_descr_it ntext NULL ," + vbCrLf + _
				"	rlu_descr_en ntext NULL ," + vbCrLf + _
				"	rlu_descr_fr ntext NULL ," + vbCrLf + _
				"	rlu_descr_es ntext NULL ," + vbCrLf + _
				"	rlu_descr_de ntext NULL ," + vbCrLf + _
				"	rlu_area_id int NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"CREATE TABLE dbo.irel_periodi (" + vbCrLf + _
				"	rpe_id int IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
				"	rpe_luogo_id int NULL ," + vbCrLf + _
				"	rpe_descr_it ntext NULL ," + vbCrLf + _
				"	rpe_descr_en ntext NULL ," + vbCrLf + _
				"	rpe_descr_fr ntext NULL ," + vbCrLf + _
				"	rpe_descr_es ntext NULL ," + vbCrLf + _
				"	rpe_descr_de ntext NULL ," + vbCrLf + _
				"	rpe_dal datetime NULL ," + vbCrLf + _
				"	rpe_al datetime NULL " + vbCrLf + _
				" ); " + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_descrRag ADD CONSTRAINT PK_itb_anagrafiche_descrRag " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( adr_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_tipi ADD CONSTRAINT PK_itb_anagrafiche_tipi " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( ant_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_aree ADD CONSTRAINT PK_itb_aree " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( are_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi_categorie ADD CONSTRAINT PK_itb_eventi_categorie " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( evc_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi_descrittori ADD CONSTRAINT PK_itb_eventi_descrittori " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( evd_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi_tipologie ADD CONSTRAINT PK_itb_eventi_tipologie " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( evt_id ) ;  " + vbCrLf + _
				"ALTER TABLE dbo.ilog_admin ADD CONSTRAINT PK_ilog_admin " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( ilog_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_admin ADD CONSTRAINT PK_irel_admin " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( adm_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_evCategorie_descrittori ADD CONSTRAINT PK_irel_evCategorie_descrittori " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( rcd_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_descrittori ADD CONSTRAINT PK_itb_anagrafiche_descrittori " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( and_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi ADD CONSTRAINT PK_itb_eventi " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( eve_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_anTipi_descrittori ADD CONSTRAINT PK_irel_anTipi_descrittori " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( rtd_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_collegati ADD CONSTRAINT PK_irel_eventi_collegati " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( rec_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_descrCat ADD CONSTRAINT PK_irel_eventi_descrCat " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( red_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_img ADD CONSTRAINT PK_irel_eventi_img " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( evi_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche ADD CONSTRAINT PK_itb_anagrafiche " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( ana_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_anagrafiche_descrTipi ADD CONSTRAINT PK_irel_anagrafiche_descrTipi " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( rad_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_anagrafiche_img ADD CONSTRAINT PK_irel_anagrafiche_img " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( ani_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_luoghi ADD CONSTRAINT PK_irel_luoghi " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( rlu_id) ;  " + vbCrLf + _
				"ALTER TABLE dbo.irel_periodi ADD CONSTRAINT PK_irel_periodi " + vbCrLf + _
				"	PRIMARY KEY  CLUSTERED ( rpe_id ) ; " + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_tipi ADD" + vbCrLf + _
				"	CONSTRAINT FK_itb_anagrafiche_tipi_itb_anagrafiche_tipi__padre_base" + vbCrLf + _
				"		FOREIGN KEY (ant_tipologia_padre_base)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_tipi ( ant_id )," + vbCrLf + _
				"	CONSTRAINT FK_itb_anagrafiche_tipi_itb_anagrafiche_tipi__padre" + vbCrLf + _
				"		FOREIGN KEY (ant_padre_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_tipi ( ant_id ) ;" + vbCrLf + _
				"ALTER TABLE dbo.itb_aree ADD" + vbCrLf + _
				"	CONSTRAINT FK_itb_aree_itb_aree__padre" + vbCrLf + _
				"		FOREIGN KEY (are_padre_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_aree ( are_id )," + vbCrLf + _
				"	CONSTRAINT FK_itb_aree_itb_aree__padre_base" + vbCrLf + _
				"		FOREIGN KEY (are_tipologia_padre_base)" + vbCrLf + _
				"		REFERENCES dbo.itb_aree ( are_id ) ;" + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi_categorie ADD" + vbCrLf + _
				"	CONSTRAINT FK_itb_eventi_categorie_itb_eventi_categorie__padre" + vbCrLf + _
				"		FOREIGN KEY (evc_padre_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_categorie ( evc_id )," + vbCrLf + _
				"	CONSTRAINT FK_itb_eventi_categorie_itb_eventi_categorie__padre_base" + vbCrLf + _
				"		FOREIGN KEY (evc_tipologia_padre_base)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_categorie ( evc_id ) ;" + vbCrLf + _
				"ALTER TABLE dbo.ilog_admin ADD" + vbCrLf + _
				"	CONSTRAINT FK_ilog_admin_tb_admin" + vbCrLf + _
				"		FOREIGN KEY (ilog_admin_id)" + vbCrLf + _
				"		REFERENCES dbo.tb_admin ( id_admin )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_admin ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_admin_tb_admin" + vbCrLf + _
				"		FOREIGN KEY (adm_id)" + vbCrLf + _
				"		REFERENCES dbo.tb_admin ( id_admin )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_evCategorie_descrittori ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_evCategorie_descrittori_itb_eventi_categorie" + vbCrLf + _
				"		FOREIGN KEY (rcd_categoria_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_categorie ( evc_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_evCategorie_descrittori_itb_eventi_descrittori" + vbCrLf + _
				"		FOREIGN KEY (rcd_descrittore_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_descrittori ( evd_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_descrittori ADD" + vbCrLf + _
				"	CONSTRAINT FK_itb_anagrafiche_descrittori_itb_anagrafiche_descrRag" + vbCrLf + _
				"		FOREIGN KEY (and_raggruppamento_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_descrRag ( adr_id ) ;" + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi ADD" + vbCrLf + _
				"	CONSTRAINT FK_itb_eventi_itb_eventi_categorie" + vbCrLf + _
				"		FOREIGN KEY (eve_categoria_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_categorie ( evc_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"   CONSTRAINT FK_itb_eventi_itb_eventi_categorie__alt" + vbCrLf + _
				"		FOREIGN KEY (eve_alt_categoria_id) " + vbCrLf + _
				"		REFERENCES itb_eventi_categorie (evc_id), " + vbCrLf + _
				"	CONSTRAINT FK_itb_eventi_itb_eventi_tipologie" + vbCrLf + _
				"		FOREIGN KEY (eve_tipologia_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_tipologie ( evt_id ) ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_anTipi_descrittori ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_anTipi_descrittori_itb_anagrafiche_descrittori" + vbCrLf + _
				"		FOREIGN KEY (rtd_descrittore_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_descrittori ( and_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_anTipi_descrittori_itb_anagrafiche_tipi" + vbCrLf + _
				"		FOREIGN KEY (rtd_tipologia_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_tipi ( ant_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_collegati ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_eventi_collegati_itb_eventi" + vbCrLf + _
				"		FOREIGN KEY (rec_evento_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi ( eve_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_eventi_collegati_itb_eventi_col" + vbCrLf + _
				"		FOREIGN KEY (rec_eventoCol_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi ( eve_id ) ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_descrCat ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_eventi_descrCat_itb_eventi" + vbCrLf + _
				"		FOREIGN KEY (red_evento_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi ( eve_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_eventi_descrCat_itb_eventi_descrittori" + vbCrLf + _
				"		FOREIGN KEY (red_descrittore_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi_descrittori ( evd_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_img ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_eventi_img_itb_eventi" + vbCrLf + _
				"		FOREIGN KEY (evi_evento_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi ( eve_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche ADD" + vbCrLf + _
				"	CONSTRAINT FK_itb_anagrafiche_itb_anagrafiche_tipi" + vbCrLf + _
				"		FOREIGN KEY (ana_tipo_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_tipi ( ant_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"   CONSTRAINT FK_itb_anagrafiche_itb_anagrafiche_tipi_alt " + vbCrLf + _
				"		FOREIGN KEY (ana_alt_tipo_id) " + vbCrLf + _
				"		REFERENCES itb_anagrafiche_tipi (ant_id)," + vbCrLf + _
				"	CONSTRAINT FK_itb_anagrafiche_itb_aree" + vbCrLf + _
				"		FOREIGN KEY (ana_area_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_aree ( are_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_itb_anagrafiche_tb_Indirizzario" + vbCrLf + _
				"		FOREIGN KEY (ana_id)" + vbCrLf + _
				"		REFERENCES dbo.tb_Indirizzario ( IDElencoIndirizzi ) ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_anagrafiche_descrTipi ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_anagrafiche_descrTipi_itb_anagrafiche" + vbCrLf + _
				"		FOREIGN KEY (rad_anagrafica_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche ( ana_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_anagrafiche_descrTipi_itb_anagrafiche_descrittori" + vbCrLf + _
				"		FOREIGN KEY (rad_descrittore_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche_descrittori ( and_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_anagrafiche_img ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_anagrafiche_img_itb_anagrafiche" + vbCrLf + _
				"		FOREIGN KEY (ani_anagrafica_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche ( ana_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_luoghi ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_luoghi_itb_anagrafiche" + vbCrLf + _
				"		FOREIGN KEY (rlu_anagrafica_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_anagrafiche ( ana_id )," + vbCrLf + _
				"	CONSTRAINT FK_irel_luoghi_itb_aree" + vbCrLf + _
				"		FOREIGN KEY (rlu_area_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_aree ( are_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ," + vbCrLf + _
				"	CONSTRAINT FK_irel_luoghi_itb_eventi" + vbCrLf + _
				"		FOREIGN KEY (rlu_evento_id)" + vbCrLf + _
				"		REFERENCES dbo.itb_eventi ( eve_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_periodi ADD" + vbCrLf + _
				"	CONSTRAINT FK_irel_periodi_irel_luoghi" + vbCrLf + _
				"		FOREIGN KEY (rpe_luogo_id)" + vbCrLf + _
				"		REFERENCES dbo.irel_luoghi ( rlu_id )" + vbCrLf + _
				"		ON DELETE CASCADE  ON UPDATE CASCADE ;" + vbCrLf + _
				"ALTER TABLE dbo.irel_luoghi" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_irel_luoghi_itb_anagrafiche" + vbCrLf + _
				"ALTER TABLE dbo.irel_eventi_collegati" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_irel_eventi_collegati_itb_eventi_col" + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_eventi_itb_eventi_tipologie" + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_descrittori" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_anagrafiche_descrittori_itb_anagrafiche_descrRag" + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi_categorie" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_eventi_categorie_itb_eventi_categorie__padre" + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi_categorie" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_eventi_categorie_itb_eventi_categorie__padre_base" + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_tipi" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_anagrafiche_tipi_itb_anagrafiche_tipi__padre" + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche_tipi" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_anagrafiche_tipi_itb_anagrafiche_tipi__padre_base" + vbCrLf + _
				"ALTER TABLE dbo.itb_eventi" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_eventi_itb_eventi_categorie__alt" + vbCrLf + _
				"ALTER TABLE dbo.itb_anagrafiche" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_anagrafiche_itb_anagrafiche_tipi_alt" + vbCrLf + _
				"ALTER TABLE dbo.itb_aree" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_aree_itb_aree__padre" + vbCrLf + _
				"ALTER TABLE dbo.itb_aree" + vbCrLf + _
				"	NOCHECK CONSTRAINT FK_itb_aree_itb_aree__padre_base; " + vbCrLf + _
				" CREATE PROCEDURE dbo.spws_AreeIncluse ( " + vbCrLf + _
				"     @id int = 0, " + vbCrLf + _
				"     @are_list varchar(400) OUT " + vbCrLf + _
				" ) AS " + vbCrLf + _
				"     SET NOCOUNT ON " + vbCrLf + _
				"     DECLARE @are_id AS int " + vbCrLf + _
				"     DECLARE @temp_out AS varchar(200) " + vbCrLf + _
				vbCrLf + _
				"     SET @temp_out = CAST(@id AS VARCHAR) " + vbCrLf + _
				"     SET @are_id = (SELECT MIN(are_id) FROM itb_aree WHERE (are_padre_id = @id)) " + vbCrLf + _
				"     WHILE @are_id IS NOT NULL " + vbCrLf + _
				"     BEGIN " + vbCrLf + _
				"         SET @are_list = ISNULL(@are_list,@temp_out) + ',' + ISNULL(CAST(@are_id AS varchar),'') " + vbCrLf + _
				"         -- PRINT @ant_list " + vbCrLf + _
				"         EXEC dbo.spws_AreeIncluse @are_id, @are_list " + vbCrLf + _
				"         SET @are_id = (SELECT MIN(are_id) a FROM itb_aree WHERE (are_padre_id = @id) AND (are_id > @are_id) ) " + vbCrLf + _
				"     END " + vbCrLf + _
				"     SET @are_list = ISNULL(@are_list,@temp_out) " + vbCrLf + _
				"     RETURN ; " + vbCrLf + _
				" CREATE PROCEDURE dbo.spws_CategorieEventiInclusi( " + vbCrLf + _
				"     @id int = 0, " + vbCrLf + _
				"     @evc_list varchar(400) OUT " + vbCrLf + _
				" ) AS " + vbCrLf + _
				"     SET NOCOUNT ON " + vbCrLf + _
				"     DECLARE @evc_id AS int " + vbCrLf + _
				"     DECLARE @temp_out AS varchar(200) " + vbCrLf + _
				"     SET @temp_out = CAST(@id AS VARCHAR) " + vbCrLf + _
				"     SET @evc_id = (SELECT MIN(evc_id) FROM itb_eventi_categorie WHERE (evc_padre_id = @id)) " + vbCrLf + _
				"     WHILE @evc_id IS NOT NULL " + vbCrLf + _
				"     BEGIN " + vbCrLf + _
				"        SET @evc_list = ISNULL(@evc_list,@temp_out) + ',' + ISNULL(CAST(@evc_id AS varchar),'') " + vbCrLf + _
				"        -- PRINT @ant_list " + vbCrLf + _
				"        EXEC dbo.spws_CategorieEventiInclusi @evc_id, @evc_list " + vbCrLf + _
				"        SET @evc_id = (SELECT MIN(evc_id) a FROM itb_eventi_categorie WHERE (evc_padre_id = @id) AND (evc_id > @evc_id) ) " + vbCrLf + _
				"     END " + vbCrLf + _
				"     SET @evc_list = ISNULL(@evc_list,@temp_out) " + vbCrLf + _
				"     RETURN ; " + _
				" CREATE PROCEDURE dbo.spws_CategorieIncluse( " + vbCrLf + _
				"     @id int = 0, " + vbCrLf + _
				"     @ant_list varchar(400) OUT " + vbCrLf + _
				" ) AS " + vbCrLf + _
				"     SET NOCOUNT ON " + vbCrLf + _
				"     DECLARE @ant_id AS int " + vbCrLf + _
				"     DECLARE @temp_out AS varchar(200) " + vbCrLf + _
				"     SET @temp_out = CAST(@id AS VARCHAR) " + vbCrLf + _
				"     SET @ant_id = (SELECT MIN(ant_id) FROM itb_anagrafiche_tipi WHERE (ant_padre_id = @id)) " + vbCrLf + _
				"     WHILE @ant_id IS NOT NULL " + vbCrLf + _
				"     BEGIN " + vbCrLf + _
				"         SET @ant_list = ISNULL(@ant_list,@temp_out) + ',' + ISNULL(CAST(@ant_id AS varchar),'') " + vbCrLf + _
				"         -- PRINT @ant_list " + vbCrLf + _
				"         EXEC dbo.spws_CategorieIncluse @ant_id, @ant_list " + vbCrLf + _
				"         SET @ant_id = (SELECT MIN(ant_id) a FROM itb_anagrafiche_tipi WHERE (ant_padre_id = @id) AND (ant_id > @ant_id) ) " + vbCrLf + _
				"     END " + vbCrLf + _
				"     SET @ant_list = ISNULL(@ant_list,@temp_out) " + vbCrLf + _
				"     RETURN "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'ATTIVAZIONE NEXT-INFO CON RELATIVI PARAMETRI
'...........................................................................................
function Activate_NEXTINFO(conn)
    Activate_NEXTINFO = _
                " INSERT INTO tb_siti(sito_nome,sito_dir, sito_p1, sito_amministrazione, id_sito, sito_prmEsterni_Admin, sito_prmesterni_sito ) " + _
				"     VALUES('NEXT-info [gestione informazioni ed eventi]', 'NEXTinfo',	'INFO_USER', 1, " & NEXTINFO & ", '../NEXTinfo/PassportAdmin.asp', '../NEXTinfo/PassportSito.asp'); " + vbCrLf + _
                " INSERT INTO tb_rubriche (nome_rubrica, locked_rubrica, rubrica_esterna, note_rubrica) " + _
                "     VALUES('Anagrafiche - elenco completo', 1, 0, 'Utilizzata da NEXT-INFO'); " + _
                " INSERT INTO tb_siti_parametri (par_key, par_value, par_sito_id ) " + _
                "     SELECT 'RUBRICA_ANAGRAFICHE', CAST(id_rubrica AS nvarchar(20)), " & NEXTINFO & _
                "       FROM tb_rubriche " + _
                "       WHERE nome_rubrica LIKE 'Anagrafiche%' AND note_rubrica LIKE '%NEXT-INFO%' ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO  1
'...........................................................................................
'aggiunge codice dei raggruppamenti di descrittori
'...........................................................................................
function Aggiornamento__INFO__1(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento__INFO__1 = _
				" ALTER TABLE itb_anagrafiche_descrRag ADD " + _
				"	adr_codice nvarchar(50) NULL, " + _
				"	adr_note ntext NULL ; "
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO  2
'...........................................................................................
'aggiunge campi per indicazione "Aree" alternative
'...........................................................................................
function Aggiornamento__INFO__2(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento__INFO__2 = _
				" ALTER TABLE itb_anagrafiche ADD " + _
				"	ana_alt_area_id INT NULL; " + _
				" ALTER TABLE irel_luoghi ADD " + _
				"	rlu_alt_area_id INT NULL; " + _
				" ALTER TABLE itb_anagrafiche ADD CONSTRAINT FK_itb_anagrafiche_itb_aree_alt " + _
				"	FOREIGN KEY (ana_alt_area_id) REFERENCES itb_aree (are_id) ; " + _
				" ALTER TABLE itb_anagrafiche NOCHECK CONSTRAINT FK_itb_anagrafiche_itb_aree_alt ; " + _
				" ALTER TABLE irel_luoghi ADD CONSTRAINT FK_irel_luoghi_itb_aree_alt " + _
				"	FOREIGN KEY (rlu_alt_area_id) REFERENCES itb_aree (are_id) ; " + _
				" ALTER TABLE irel_luoghi NOCHECK CONSTRAINT FK_irel_luoghi_itb_aree_alt ; "
	end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO  3
'...........................................................................................
'aggiunge campo degli eventi per indicare "evento di particolare interesse"
'...........................................................................................
function Aggiornamento__INFO__3(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento__INFO__3 = _
				" ALTER TABLE itb_eventi ADD " + _
                "   eve_interessante BIT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO  4
'...........................................................................................
'ClassCategorie: aggiunge il campo per la gestione della lista degli IDs dei padri
'...........................................................................................
function AggiornamentoSpeciale__INFO__4(DB, rs, version)
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "itb_eventi_categorie", "evc")
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "itb_anagrafiche_tipi", "ant")
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "itb_aree", "are")
    AggiornamentoSpeciale__INFO__4 = "SELECT * FROM AA_versione"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO  5
'...........................................................................................
'aggiunge una tabella per gestione prenotazioni
'...........................................................................................
Function Aggiornamento__INFO__5(conn)
    Aggiornamento__INFO__5 = _
		" CREATE TABLE dbo.itb_anagrafiche_prenotazioni (" + vbCrLf + _
		"	anp_id INT IDENTITY (1, 1) NOT NULL ," + vbCrLf + _
		"	anp_ana_id INT NOT NULL ," + vbCrLf + _
		"	anp_contatto_id INT NOT NULL ," + vbCrLf + _
		"	anp_dataArrivo DATETIME NULL ," + vbCrLf + _
		"	anp_dataPartenza DATETIME NULL ," + vbCrLf + _
		"	anp_adulti INT NULL ," + vbCrLf + _
		"	anp_bambini INT NULL ," + vbCrLf + _
		"	anp_cameraTipo NVARCHAR(255) ," + vbCrLf + _
		"	anp_cameraNumero INT NULL, " + vbCrLf + _
		" 	anp_data DATETIME NULL" + vbCrLf + _
		" ); " + vbCrLf + _
		" ALTER TABLE dbo.itb_anagrafiche_prenotazioni ADD CONSTRAINT FK_itb_anagrafiche_prenotazioni_itb_anagrafiche " + _
		" FOREIGN KEY (anp_ana_id) REFERENCES itb_anagrafiche (ana_id); " + vbCrLf + _
		" ALTER TABLE dbo.itb_anagrafiche_prenotazioni ADD CONSTRAINT FK_itb_anagrafiche_prenotazioni_tb_indirizzario " + _
		" FOREIGN KEY (anp_contatto_id) REFERENCES tb_indirizzario (idElencoIndirizzi) " + vbCrLf + _
		" ON DELETE CASCADE  ON UPDATE CASCADE"
End Function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 6
'...........................................................................................
'   aggiunge campi alle categorie di eventi, di anagrafiche ed alle aree per suddivisioni
'   in principali ed alternative
'...........................................................................................
function Aggiornamento__INFO__6(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento__INFO__6 = _
				" ALTER TABLE itb_anagrafiche_tipi ADD " + _
                "       ant_alternativa BIT NULL, " + _
                "       ant_principale BIT NULL ; " + _
                " ALTER TABLE itb_eventi_categorie ADD " + _
                "       evc_alternativa BIT NULL, " + _
                "       evc_principale BIT NULL ; " + _
                " ALTER TABLE itb_aree ADD " + _
                "       are_alternativa BIT NULL, " + _
                "       are_principale BIT NULL ; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 7
'...........................................................................................
'   aggiunge campi periodo di massima agli eventi (ricalcano automaticamente il limite minimo
'   ed il limite massimo dei periodi degli eventi)
'...........................................................................................
function Aggiornamento__INFO__7(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__7 = _
                " ALTER TABLE itb_eventi ADD " + _
                "   eve_min_dal SMALLDATETIME NULL, " + _
                "   eve_max_al SMALLDATETIME NULL ;" + _
                " UPDATE itb_eventi SET eve_min_dal = (SELECT MIN(rpe_dal) FROM irel_periodi INNER JOIN irel_luoghi ON irel_periodi.rpe_luogo_id = irel_luoghi.rlu_id WHERE irel_luoghi.rlu_evento_id = itb_eventi.eve_id) ; " + _
                " UPDATE itb_eventi SET eve_max_al = (SELECT MAX(rpe_al) FROM irel_periodi INNER JOIN irel_luoghi ON irel_periodi.rpe_luogo_id = irel_luoghi.rlu_id WHERE irel_luoghi.rlu_evento_id = itb_eventi.eve_id) ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 8
'...........................................................................................
'   crea vista per composizione query eventi con colonna calcolata per visibilita'
'...........................................................................................
function Aggiornamento__INFO__8(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__8 = _
                " CREATE VIEW dbo.iv_eventi AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          ( CASE WHEN ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "                      CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS eve_visibile_assoluto, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 9
'...........................................................................................
'   crea vista eventi visibili
'...........................................................................................
function Aggiornamento__INFO__9(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__9 = _
                " CREATE VIEW dbo.iv_eventi_visibili AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "         CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLf + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 10
'...........................................................................................
'   crea vista anagrafiche con campo calcolato che indica la visibilita'
'...........................................................................................
function Aggiornamento__INFO__10(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__10 = _
                " CREATE VIEW dbo.iv_anagrafiche AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          ( CASE WHEN ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS ana_visibile_assoluto, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 11
'...........................................................................................
'   crea vista anagrafiche visibili
'...........................................................................................
function Aggiornamento__INFO__11(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__11 = _
                " CREATE VIEW dbo.iv_anagrafiche_visibili AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 12
'...........................................................................................
'   modifica permessi di accesso degli amministratori per agginuta di un nuovo profilo 
'...........................................................................................
function Aggiornamento__INFO__12(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__12 = _
                " UPDATE irel_admin SET adm_permesso = IsNull(adm_permesso, 0) + 1 "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 13
'...........................................................................................
'   modifica vista angrafiche per aggiunta campi indirizzazio
'...........................................................................................
function Aggiornamento__INFO__13(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__13 = _
                DropObject(conn, "iv_anagrafiche", "VIEW") + _
                DropObject(conn, "iv_anagrafiche_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_anagrafiche AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          ( CASE WHEN ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS ana_visibile_assoluto, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.NomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, tb_Indirizzario.CittaElencoIndirizzi, " + vbCrLf + _
                "          tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.lingua " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_anagrafiche_visibili AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.NomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, tb_Indirizzario.CittaElencoIndirizzi, " + vbCrLf + _
                "          tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.lingua " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 14
'...........................................................................................
'   corregge errore viste su eventi
'...........................................................................................
function Aggiornamento__INFO__14(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__14 = _
                DropObject(conn, "iv_eventi", "VIEW") + _
                DropObject(conn, "iv_eventi_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_eventi AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          ( CASE WHEN ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "                      CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS eve_visibile_assoluto, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie_alt.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_eventi_visibili AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie_alt.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "         CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLf + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 15
'...........................................................................................
'   aggiunge colonne per stampa calendario eventi
'...........................................................................................
function Aggiornamento__INFO__15(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__15 = _
                " ALTER TABLE itb_eventi ADD"+ _
                "   eve_descr_calendario_it NTEXT NULL,"+ _
                "   eve_descr_calendario_en NTEXT NULL,"+ _
                "   eve_descr_calendario_fr NTEXT NULL,"+ _
                "   eve_descr_calendario_de NTEXT NULL,"+ _
                "   eve_descr_calendario_es NTEXT NULL"
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 16
'...........................................................................................
'   aggiorna vista anagrafiche ed eventi perche' erano state generate con codice sql troppo 
'   lungo: crea problemi nella rigenerazione automatica delle viste da codice aggiornamento.
'...........................................................................................
function Aggiornamento__INFO__16(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__16 = _
                DropObject(conn, "iv_anagrafiche", "VIEW") + _
                DropObject(conn, "iv_anagrafiche_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_anagrafiche AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          ( CASE WHEN ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS ana_visibile_assoluto, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.NomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, tb_Indirizzario.CittaElencoIndirizzi, " + vbCrLf + _
                "          tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.lingua " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_anagrafiche_visibili AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.NomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, tb_Indirizzario.CittaElencoIndirizzi, " + vbCrLf + _
                "          tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.lingua " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                " ; " + _
                DropObject(conn, "iv_eventi", "VIEW") + _
                DropObject(conn, "iv_eventi_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_eventi AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          ( CASE WHEN ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "                      CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS eve_visibile_assoluto, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie_alt.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_eventi_visibili AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie_alt.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "         CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLf + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 17
'...........................................................................................
'   corregge registrazione script per gestione permessi aggiuntivi degli utenti dal next-passport
'...........................................................................................
function Aggiornamento__INFO__17(conn)
    Aggiornamento__INFO__17 = _
        " UPDATE tb_siti SET " + _
            " sito_prmEsterni_admin = '../NEXTinfo/PassportAdmin.asp', " + _
            " sito_prmEsterni_sito = '../NEXTinfo/PassportSito.asp' " + _
        " WHERE id_sito=" & NEXTINFO
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 18
'...........................................................................................
'   aggiorna le viste delle anagrafiche per i campi google maps
'...........................................................................................
function Aggiornamento__INFO__18(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__18 = _
                DropObject(conn, "iv_anagrafiche", "VIEW") + _
                DropObject(conn, "iv_anagrafiche_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_anagrafiche AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          ( CASE WHEN ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS ana_visibile_assoluto, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.NomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, tb_Indirizzario.CittaElencoIndirizzi, " + vbCrLf + _
                "          tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.lingua, " + vbCrLF + _
                "          tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_anagrafiche_visibili AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.NomeElencoIndirizzi, tb_Indirizzario.CognomeElencoIndirizzi, tb_Indirizzario.TitoloElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.NomeOrganizzazioneElencoIndirizzi, tb_Indirizzario.IndirizzoElencoIndirizzi, tb_Indirizzario.CittaElencoIndirizzi, " + vbCrLf + _
                "          tb_Indirizzario.StatoProvElencoIndirizzi, tb_Indirizzario.ZonaElencoIndirizzi, tb_Indirizzario.CAPElencoIndirizzi, tb_Indirizzario.CountryElencoIndirizzi, " + vbCrLF + _
                "          tb_Indirizzario.isSocieta, tb_Indirizzario.ModoRegistra, tb_Indirizzario.lingua, " + vbCrLF + _
                "          tb_Indirizzario.google_maps_latitudine, tb_Indirizzario.google_maps_longitudine " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 19
'...........................................................................................
'   modifica il campo tipo camera delle prenotazioni delle anagrafiche
'...........................................................................................
function Aggiornamento__INFO__19(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__19 = _
                " ALTER TABLE itb_anagrafiche_prenotazioni DROP COLUMN anp_cameraNumero;" + vbCrLF + _
				" ALTER TABLE itb_anagrafiche_prenotazioni DROP COLUMN anp_cameraTipo;" + vbCrLF + _
				" ALTER TABLE itb_anagrafiche_prenotazioni ADD" + vbCrLF + _
				"	anp_testo " + SQL_CharField(Conn, 0) + " NULL"
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 20
'...........................................................................................
'   aggiunge campo per conteggio click sul sito web dell'anagrafica
'...........................................................................................
function Aggiornamento__INFO__20(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__20 = _
                " ALTER TABLE itb_anagrafiche ADD" + vbCrLF + _
				"	ana_web_click INT NULL," + vbCrLF + _
				"	ana_web_reset DATETIME NULL;" + vbCrLF + _
				" UPDATE itb_anagrafiche SET ana_web_reset = GETDATE()"
    end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 21
'...........................................................................................
' NICOLA - 31/05/2010
'...........................................................................................
'   aggiorna le viste delle anagrafiche per aggiungere i campi mancanti della tabella indirizzario
'...........................................................................................
function Aggiornamento__INFO__21(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__21 = _
                DropObject(conn, "iv_anagrafiche", "VIEW") + _
                DropObject(conn, "iv_anagrafiche_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_anagrafiche AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          ( CASE WHEN ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS ana_visibile_assoluto, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.* " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_anagrafiche_visibili AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.* " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 22
'...........................................................................................
' Giacomo 22/04/2011
'...........................................................................................
' aggiunge parametri per inibire la cancellazione degli eventi
'...........................................................................................
function Aggiornamento__INFO__22(conn)
	Aggiornamento__INFO__22 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFO__22(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTINFO)) <> "" then
		CALL AddParametroSito(conn, "INFO_INIBISCI_CANCELLAZIONE_EVENTI", _
									0, _
									"Flag per impedire la cancellazione degli eventi", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTINFO, _
									null, null, null, null, null)
		AggiornamentoSpeciale__INFO__22 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 23
'...........................................................................................
'Giacomo 29/04/2010
'aggiunta parametri copiandoli da parametri old
'...........................................................................................
function Aggiornamento__INFO__23(conn)
	Aggiornamento__INFO__23 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__INFO__23(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTINFO)) <> "" then
		sql = "SELECT * FROM tb_siti_parametri WHERE par_sito_id=" & NEXTINFO
		rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
		while not rs.eof 
			CALL AddParametroSito(conn, rs("par_key"), _
										null, _
										"(PARAMETRI OLD)", _
										"", _
										adVarChar, _
										0, _
										"", _
										1, _
										1, _
										NEXTINFO, _
										rs("par_value"), null, null, null, null)
			rs.moveNext
		wend
		rs.close
		sql = ""
		AggiornamentoSpeciale__INFO__23 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 24
'...........................................................................................
'Giacomo 08/07/2011
'aggiunta parametri copiandoli da parametri old
'...........................................................................................
function Aggiornamento__INFO__24(conn, lingua_abbr)
	Aggiornamento__INFO__24 = _
		  " ALTER TABLE irel_anagrafiche_descrTipi ADD " + vbCrLf + _
		  " 	rad_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	rad_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE irel_anagrafiche_img ADD " + vbCrLf + _
		  " 	ani_didascalia_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE irel_eventi_collegati ADD " + vbCrLf + _
		  " 	rec_link_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	rec_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	rec_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE irel_eventi_descrCat ADD " + vbCrLf + _
		  " 	red_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	red_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _	  
		  " ALTER TABLE irel_eventi_img ADD " + vbCrLf + _
		  " 	evi_didascalia_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _		  
		  " ALTER TABLE irel_luoghi ADD " + vbCrLf + _
		  " 	rlu_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _			  
		  " ALTER TABLE irel_periodi ADD " + vbCrLf + _
		  " 	rpe_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _		  
		  " ALTER TABLE itb_anagrafiche ADD " + vbCrLf + _
		  " 	ana_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _			  
		  " ALTER TABLE itb_anagrafiche_descrittori ADD " + vbCrLf + _
		  " 	and_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	and_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 50) + " NULL;" + vbCrLf + _	  
		  " ALTER TABLE itb_anagrafiche_descrRag ADD " + vbCrLf + _
		  " 	adr_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;" + vbCrLf + _	
		  " ALTER TABLE itb_anagrafiche_tipi ADD " + vbCrLf + _
		  " 	ant_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	ant_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _	  
		  " ALTER TABLE itb_aree ADD " + vbCrLf + _
		  " 	are_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	are_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _		  
		  " ALTER TABLE itb_eventi ADD " + vbCrLf + _
		  " 	eve_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	eve_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _	
		  " 	eve_ridotto_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _	
		  " 	eve_info_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _	
		  " 	eve_descr_calendario_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _	
		  " ALTER TABLE itb_eventi_categorie ADD " + vbCrLf + _
		  " 	evc_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	evc_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE itb_eventi_descrittori ADD " + vbCrLf + _
		  " 	evd_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	evd_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 50) + " NULL;" + vbCrLf + _	
		  " ALTER TABLE itb_eventi_tipologie ADD " + vbCrLf + _
		  " 	evt_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL;"
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 25
'...........................................................................................
' Giacomo - 08/07/2011
'...........................................................................................
' aggiorna le viste iv_anagrafiche e iv_anagrafiche_visibili per aggiungere i campi mancanti nelle nuove lingue
'...........................................................................................
function Aggiornamento__INFO__25(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__25 = _
                DropObject(conn, "iv_anagrafiche", "VIEW") + _
                DropObject(conn, "iv_anagrafiche_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_anagrafiche AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          ( CASE WHEN ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS ana_visibile_assoluto, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
				"		   itb_anagrafiche_tipi.ant_nome_ru, itb_anagrafiche_tipi.ant_nome_pt, itb_anagrafiche_tipi.ant_nome_cn, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
				"		   itb_anagrafiche_tipi_alt.ant_nome_ru AS alt_ant_nome_ru, itb_anagrafiche_tipi_alt.ant_nome_pt AS alt_ant_nome_pt, itb_anagrafiche_tipi_alt.ant_nome_cn AS alt_ant_nome_cn, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
				"		   itb_aree.are_nome_ru, itb_aree.are_nome_pt, itb_aree.are_nome_cn, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
				"		   itb_aree_alt.are_nome_ru AS alt_are_nome_ru, itb_aree_alt.are_nome_pt AS alt_are_nome_pt, itb_aree_alt.are_nome_cn AS alt_are_nome_cn, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.* " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_anagrafiche_visibili AS " + vbCrLf + _
                "   SELECT itb_anagrafiche.*, " + vbCrLf + _
                "          itb_anagrafiche_tipi.ant_nome_it, itb_anagrafiche_tipi.ant_nome_en, itb_anagrafiche_tipi.ant_nome_fr, itb_anagrafiche_tipi.ant_nome_es, itb_anagrafiche_tipi.ant_nome_de, " + vbCrLF + _
				"		   itb_anagrafiche_tipi.ant_nome_ru, itb_anagrafiche_tipi.ant_nome_pt, itb_anagrafiche_tipi.ant_nome_cn, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_codice, itb_anagrafiche_tipi.ant_padre_id, itb_anagrafiche_tipi.ant_tipologia_padre_base, itb_anagrafiche_tipi.ant_tipologie_padre_lista, " + vbCrLF + _
                "          itb_anagrafiche_tipi.ant_ordine, itb_anagrafiche_tipi.ant_ordine_assoluto, itb_anagrafiche_tipi.ant_visibile, itb_anagrafiche_tipi.ant_albero_visibile, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_it AS alt_ant_nome_it, itb_anagrafiche_tipi_alt.ant_nome_en AS alt_ant_nome_en, itb_anagrafiche_tipi_alt.ant_nome_fr AS alt_ant_nome_fr, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_nome_es AS alt_ant_nome_es, itb_anagrafiche_tipi_alt.ant_nome_de AS alt_ant_nome_de, " + vbCrLF + _
				"		   itb_anagrafiche_tipi_alt.ant_nome_ru AS alt_ant_nome_ru, itb_anagrafiche_tipi_alt.ant_nome_pt AS alt_ant_nome_pt, itb_anagrafiche_tipi_alt.ant_nome_cn AS alt_ant_nome_cn, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_codice AS alt_ant_codice, itb_anagrafiche_tipi_alt.ant_padre_id AS alt_ant_padre_id, itb_anagrafiche_tipi_alt.ant_tipologia_padre_base AS alt_ant_tipologia_padre_base, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_tipologie_padre_lista AS alt_ant_tipologie_padre_lista, itb_anagrafiche_tipi_alt.ant_ordine AS alt_ant_ordine, " + vbCrLF + _
                "          itb_anagrafiche_tipi_alt.ant_ordine_assoluto AS alt_ant_ordine_assoluto, itb_anagrafiche_tipi_alt.ant_visibile AS alt_ant_visibile, itb_anagrafiche_tipi_alt.ant_albero_visibile AS alt_ant_albero_visibile, " + vbCrLF + _
                "          itb_aree.are_nome_it, itb_aree.are_nome_en, itb_aree.are_nome_fr, itb_aree.are_nome_es, itb_aree.are_nome_de, " + vbCrLF + _
				"		   itb_aree.are_nome_ru, itb_aree.are_nome_pt, itb_aree.are_nome_cn, " + vbCrLF + _
                "          itb_aree.are_codice, itb_aree.are_padre_id, itb_aree.are_tipologia_padre_base, itb_aree.are_tipologie_padre_lista, " + vbCrLF + _
                "          itb_aree.are_ordine, itb_aree.are_ordine_assoluto, itb_aree.are_visibile, itb_aree.are_albero_visibile, " + vbCrLF + _
                "          itb_aree_alt.are_nome_it AS alt_are_nome_it, itb_aree_alt.are_nome_en AS alt_are_nome_en, itb_aree_alt.are_nome_fr AS alt_are_nome_fr, " + vbCrLF + _
                "          itb_aree_alt.are_nome_es AS alt_are_nome_es, itb_aree_alt.are_nome_de AS alt_are_nome_de, " + vbCrLF + _
				"		   itb_aree_alt.are_nome_ru AS alt_are_nome_ru, itb_aree_alt.are_nome_pt AS alt_are_nome_pt, itb_aree_alt.are_nome_cn AS alt_are_nome_cn, " + vbCrLF + _
                "          itb_aree_alt.are_codice AS alt_are_codice, itb_aree_alt.are_padre_id AS alt_are_padre_id, itb_aree_alt.are_tipologia_padre_base AS alt_are_tipologia_padre_base, " + vbCrLF + _
                "          itb_aree_alt.are_tipologie_padre_lista AS alt_are_tipologie_padre_lista, itb_aree_alt.are_ordine AS alt_are_ordine, itb_aree_alt.are_ordine_assoluto AS alt_are_ordine_assoluto, " + vbCrLF + _
                "          itb_aree_alt.are_visibile AS alt_are_visibile, itb_aree_alt.are_albero_visibile AS alt_are_albero_visibile, " + vbCrLF + _
                "          tb_Indirizzario.* " + vbCrLF + _
                "   FROM itb_anagrafiche INNER JOIN itb_anagrafiche_tipi ON itb_anagrafiche.ana_tipo_id = itb_anagrafiche_tipi.ant_id " + vbCrLF + _
                "   INNER JOIN tb_Indirizzario ON itb_anagrafiche.ana_id = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
                "   INNER JOIN itb_aree ON itb_anagrafiche.ana_area_id = itb_aree.are_id " + vbCrLf + _
                "   LEFT JOIN itb_anagrafiche_tipi itb_anagrafiche_tipi_alt ON itb_anagrafiche.ana_alt_tipo_id = itb_anagrafiche_tipi_alt.ant_id " + vbCrLF + _
                "   LEFT JOIN itb_aree itb_aree_alt ON itb_anagrafiche.ana_alt_area_id = itb_aree_alt.are_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_anagrafiche_tipi.ant_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_anagrafiche_tipi.ant_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_anagrafiche_tipi_alt.ant_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_anagrafiche.ana_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_anagrafiche.ana_censurato, 0) = 0 " + vbCrLF + _
                " ; "
    end select
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO NEXT-INFO 26
'...........................................................................................
' Giacomo - 08/07/2011
'...........................................................................................
' aggiorna le viste iv_eventi e iv_eventi_visibili per aggiungere i campi mancanti nelle nuove lingue
'...........................................................................................
function Aggiornamento__INFO__26(conn)
	Select case DB_Type(conn)
        case DB_SQL
            Aggiornamento__INFO__26 = _
                DropObject(conn, "iv_eventi", "VIEW") + _
                DropObject(conn, "iv_eventi_visibili", "VIEW") + _
                " CREATE VIEW dbo.iv_eventi AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          ( CASE WHEN ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "                        IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "                      IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "                      CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "                      IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLF + _
                "                 THEN 1 ELSE 0 END ) AS eve_visibile_assoluto, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
				"		   itb_eventi_categorie.evc_nome_ru, itb_eventi_categorie.evc_nome_pt, itb_eventi_categorie.evc_nome_cn, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
				"		   itb_eventi_categorie_alt.evc_nome_ru AS alt_evc_nome_ru, itb_eventi_categorie_alt.evc_nome_pt AS alt_evc_nome_pt, itb_eventi_categorie_alt.evc_nome_cn AS alt_evc_nome_cn, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie_alt.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                " ; " + _
                " CREATE VIEW dbo.iv_eventi_visibili AS " + vbCrLf + _
                "   SELECT itb_eventi.*, itb_eventi_tipologie.*, " + vbCrLF + _
                "          itb_eventi_categorie.evc_nome_it, itb_eventi_categorie.evc_nome_en, itb_eventi_categorie.evc_nome_fr, itb_eventi_categorie.evc_nome_es, itb_eventi_categorie.evc_nome_de, " + vbCrLF + _
				"		   itb_eventi_categorie.evc_nome_ru, itb_eventi_categorie.evc_nome_pt, itb_eventi_categorie.evc_nome_cn, " + vbCrLF + _
                "          itb_eventi_categorie.evc_codice, itb_eventi_categorie.evc_padre_id, itb_eventi_categorie.evc_tipologia_padre_base, itb_eventi_categorie.evc_tipologie_padre_lista, " + vbCrLF + _
                "          itb_eventi_categorie.evc_ordine, itb_eventi_categorie.evc_ordine_assoluto, itb_eventi_categorie.evc_visibile, itb_eventi_categorie.evc_albero_visibile, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_it AS alt_evc_nome_it, itb_eventi_categorie_alt.evc_nome_en AS alt_evc_nome_en, itb_eventi_categorie_alt.evc_nome_fr AS alt_evc_nome_fr, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_nome_es AS alt_evc_nome_es, itb_eventi_categorie_alt.evc_nome_de AS alt_evc_nome_de, " + vbCrLF + _
				"		   itb_eventi_categorie_alt.evc_nome_ru AS alt_evc_nome_ru, itb_eventi_categorie_alt.evc_nome_pt AS alt_evc_nome_pt, itb_eventi_categorie_alt.evc_nome_cn AS alt_evc_nome_cn, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_codice AS alt_evc_codice, itb_eventi_categorie_alt.evc_padre_id AS alt_evc_padre_id, itb_eventi_categorie_alt.evc_tipologia_padre_base AS alt_evc_tipologia_padre_base, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_tipologie_padre_lista AS alt_evc_tipologie_padre_lista, itb_eventi_categorie_alt.evc_ordine AS alt_evc_ordine, " + vbCrLF + _
                "          itb_eventi_categorie_alt.evc_ordine_assoluto AS alt_evc_ordine_assoluto, itb_eventi_categorie_alt.evc_visibile AS alt_evc_visibile, itb_eventi_categorie_alt.evc_albero_visibile AS alt_evc_albero_visibile " + vbCrLF + _
                "   FROM itb_eventi INNER JOIN itb_eventi_categorie ON itb_eventi.eve_categoria_id = itb_eventi_categorie.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_categorie itb_eventi_categorie_alt ON itb_eventi.eve_alt_categoria_id = itb_eventi_categorie_alt.evc_id " + vbCrLF + _
                "   LEFT JOIN itb_eventi_tipologie ON itb_eventi.eve_tipologia_id = itb_eventi_tipologie.evt_id " + vbCrLF + _
                "   WHERE ( IsNull(itb_eventi_categorie.evc_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_visibile, 0)=1 ) AND " + vbCrLF + _
                "         ( IsNull(itb_eventi_categorie.evc_albero_visibile, 0)=1 OR " + vbCrLF + _
                "           IsNull(itb_eventi_categorie_alt.evc_albero_visibile, 0)=1 ) AND " + vbCrLF + _
                "         IsNull(itb_eventi.eve_visibile, 0) = 1 AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_censurato, 0) = 0 AND " + vbCrLf + _
                "         CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) <= IsNull(itb_eventi.eve_max_al, GETDATE() - 1) AND " + vbCrLf + _
                "         IsNull(itb_eventi.eve_pubblData, GETDATE() + 1) <= CONVERT(DATETIME, CONVERT(nvarchar(10), GETDATE(), 103), 103) " + vbCrLf + _
                " ; "
    end select
end function
'*******************************************************************************************






%>


