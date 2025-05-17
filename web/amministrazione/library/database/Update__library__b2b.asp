<% 
'...........................................................................................
'...........................................................................................
'libreria di funzioni che contiene gli aggiornamenti SQL delle istanze 
'SQL Server del NEXT-B2B 
'...........................................................................................
'...........................................................................................

<!--#INCLUDE FILE="../Tools.asp" -->

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'FUNZIONI PER L'INSTALLAZIONE DEL B2B

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

function Install_B2B__version_79()
    Install_B2B__version_79 = _
        "CREATE TABLE gItb_articoli ( " + vbCrLF + _ 
        "       Iart_id int IDENTITY (1, 1) NOT NULL , " + vbCrLF + _
        "   	Iart_nome_IT nvarchar (250) NULL , " + vbCrLF + _
        "   	IArt_stato_articolo nvarchar (500) NULL , " + vbCrLF + _
        "   	Iart_Marca nvarchar (500) NULL , " + vbCrLF + _
        "   	Iart_Tipologia nvarchar (1000) NULL , " + vbCrLF + _
        "   	Iart_scorta_minima int NULL , " + vbCrLF + _
        "   	Iart_lotto_minimo int NULL , " + vbCrLF + _
        "   	Iart_lotto_riordino int NULL , " + vbCrLF + _
        "   	Iart_Data_Ins_articolo smalldatetime NULL , " + vbCrLF + _
        "   	Iart_Data_Upd_articolo smalldatetime NULL , " + vbCrLF + _
        "   	IArt_Data_Ins_Import smalldatetime NULL , " + vbCrLF + _
        "   	IArt_Data_Upd_Import smalldatetime NULL , " + vbCrLF + _
        "   	Iart_x_cod_int nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_x_cod_alt nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_x_cod_pro nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_x_ID int NULL , " + vbCrLF + _
        "   	Iart_b2b_ID int NULL , " + vbCrLF + _
        "   	IArt_x_SourceCode nvarchar (10) NULL , " + vbCrLF + _
        "   	IArt_User_UPD nvarchar (50) NULL , " + vbCrLF + _
        "   	IArt_Prezzo money NULL , " + vbCrLF + _
        "   	IArt_b2b_Update_price bit NULL , " + vbCrLF + _
        "   	IArt_b2b_Update_state bit NULL , " + vbCrLF + _
        "   	IArt_aliquota_iva nvarchar (100) NULL , " + vbCrLF + _
        "   	IArt_b2b_field nvarchar (50) NULL , " + vbCrLF + _
        "   	IArt_stato_catalogo_Articolo bit NULL , " + vbCrLF + _
        "   	Iart_Marca_codice nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_tipologia_id nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_x_cod_fornitore nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_x_fornitore nvarchar (50) NULL , " + vbCrLF + _
        "   	Iart_Varianti bit NULL , " + vbCrLF + _
        "   	IArt_prezzo_var_euro real NULL , " + vbCrLF + _
        "   	IArt_prezzo_var_sconto real NULL , " + vbCrLF + _
        "   	Iart_descrizione_IT ntext NULL , " + vbCrLF + _
        "   	Iart_Note ntext NULL ) " + vbCrLF + _
        ";" + vbCrLF + _
        "CREATE TABLE dbo.glog_ordini ( " + vbCrLf + _
        "    log_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    log_data smalldatetime NULL , " + vbCrLf + _
        "    log_ordine_id int NULL , " + vbCrLf + _
        "    log_operatore nvarchar (50) NULL , " + vbCrLf + _
        "    log_operatore_area nvarchar (20) NULL , " + vbCrLf + _
        "    log_operazione_id int NULL , " + vbCrLf + _
        "    log_operazione_extra_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.glog_ordini_operazioni ( " + vbCrLf + _
        "    op_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    op_descrizione nvarchar (250) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_art_acc ( " + vbCrLf + _
        "    aa_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    aa_art_id int NOT NULL , " + vbCrLf + _
        "    aa_acc_id int NOT NULL , " + vbCrLf + _
        "    aa_tipo_id int NOT NULL , " + vbCrLf + _
        "    aa_ordine int NULL , " + vbCrLf + _
        "    aa_note_it nvarchar (250) NULL , " + vbCrLf + _
        "    aa_note_en nvarchar (250) NULL , " + vbCrLf + _
        "    aa_note_fr nvarchar (250) NULL , " + vbCrLf + _
        "    aa_note_es nvarchar (250) NULL , " + vbCrLf + _
        "    aa_note_de nvarchar (250) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_art_ctech ( " + vbCrLf + _
        "    rel_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rel_art_id int NULL , " + vbCrLf + _
        "    rel_ctech_id int NULL , " + vbCrLf + _
        "    rel_ctech_de ntext NULL , " + vbCrLf + _
        "    rel_ctech_en ntext NULL , " + vbCrLf + _
        "    rel_ctech_es ntext NULL , " + vbCrLf + _
        "    rel_ctech_fr ntext NULL , " + vbCrLf + _
        "    rel_ctech_it ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_art_valori ( " + vbCrLf + _
        "    rel_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rel_art_id int NULL , " + vbCrLf + _
        "    rel_prezzo money NULL , " + vbCrLf + _
        "    rel_giacenza_min int NULL , " + vbCrLf + _
        "    rel_lotto_riordino int NULL , " + vbCrLf + _
        "    rel_qta_min_ord int NULL , " + vbCrLf + _
        "    rel_cod_int nvarchar (50) NULL , " + vbCrLf + _
        "    rel_cod_pro nvarchar (50) NULL , " + vbCrLf + _
        "    rel_cod_alt nvarchar (50) NULL , " + vbCrLf + _
        "    rel_disabilitato bit NULL , " + vbCrLf + _
        "    rel_scontoQ_id int NULL , " + vbCrLf + _
        "    rel_ordine nvarchar (100) NULL , " + vbCrLf + _
        "    rel_external_id int NULL , " + vbCrLf + _
        "    rel_var_euro real NULL , " + vbCrLf + _
        "    rel_var_sconto real NULL , " + vbCrLf + _
        "    rel_prezzo_indipendente bit NULL , " + vbCrLf + _
        "    rel_foto_id int NULL , " + vbCrLf + _
        "    rel_descr_de ntext NULL , " + vbCrLf + _
        "    rel_descr_en ntext NULL , " + vbCrLf + _
        "    rel_descr_es ntext NULL , " + vbCrLf + _
        "    rel_descr_fr ntext NULL , " + vbCrLf + _
        "    rel_descr_it ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_art_vv ( " + vbCrLf + _
        "    rvv_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rvv_art_var_id int NULL , " + vbCrLf + _
        "    rvv_val_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_carichi_var ( " + vbCrLf + _
        "    rcv_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rcv_car_id int NULL , " + vbCrLf + _
        "    rcv_art_var_id int NULL , " + vbCrLf + _
        "    rcv_qta int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_giacenze ( " + vbCrLf + _
        "    gia_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    gia_magazzino_id int NULL , " + vbCrLf + _
        "    gia_art_var_id int NULL , " + vbCrLf + _
        "    gia_qta int NULL , " + vbCrLf + _
        "    gia_impegnato int NULL , " + vbCrLf + _
        "    gia_ordinato int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.grel_mov_var ( " + vbCrLf + _
        "    rmv_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rmv_mov_id int NULL , " + vbCrLf + _
        "    rmv_art_var_id int NULL , " + vbCrLf + _
        "    rmv_qta_richiesta int NULL , " + vbCrLf + _
        "    rmv_qta_spedita int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_accessori_tipo ( " + vbCrLf + _
        "    at_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    at_nome_it nvarchar (250) NULL , " + vbCrLf + _
        "    at_nome_en nvarchar (250) NULL , " + vbCrLf + _
        "    at_nome_fr nvarchar (250) NULL , " + vbCrLf + _
        "    at_nome_es nvarchar (250) NULL , " + vbCrLf + _
        "    at_nome_de nvarchar (250) NULL , " + vbCrLf + _
        "    at_ordine int NOT NULL , " + vbCrLf + _
        "    at_vincolo_vendita bit NOT NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_agenti ( " + vbCrLf + _
        "    ag_id int NOT NULL , " + vbCrLf + _
        "    ag_admin_id int NULL , " + vbCrLf + _
        "    ag_gruppo_id int NULL , " + vbCrLf + _
        "    ag_commissione real NULL , " + vbCrLf + _
        "    ag_codice nvarchar (20) NULL , " + vbCrLf + _
        "    ag_range_sconto_massimo int NULL , " + vbCrLf + _
        "    ag_supervisore bit NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_art_foto ( " + vbCrLf + _
        "    fo_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    fo_articolo_id int NULL , " + vbCrLf + _
        "    fo_thumb nvarchar (255) NULL , " + vbCrLf + _
        "    fo_zoom nvarchar (255) NULL , " + vbCrLf + _
        "    fo_numero nvarchar (10) NULL , " + vbCrLf + _
        "    fo_ordine int NULL , " + vbCrLf + _
        "    fo_didascalia_it nvarchar (255) NULL , " + vbCrLf + _
        "    fo_didascalia_en nvarchar (255) NULL , " + vbCrLf + _
        "    fo_didascalia_fr nvarchar (255) NULL , " + vbCrLf + _
        "    fo_didascalia_es nvarchar (255) NULL , " + vbCrLf + _
        "    fo_didascalia_de nvarchar (255) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_articoli ( " + vbCrLf + _
        "    art_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    art_nome_it nvarchar (250) NULL , " + vbCrLf + _
        "    art_nome_en nvarchar (250) NULL , " + vbCrLf + _
        "    art_nome_fr nvarchar (250) NULL , " + vbCrLf + _
        "    art_nome_es nvarchar (250) NULL , " + vbCrLf + _
        "    art_nome_de nvarchar (250) NULL , " + vbCrLf + _
        "    art_cod_int nvarchar (50) NULL , " + vbCrLf + _
        "    art_cod_pro nvarchar (50) NULL , " + vbCrLf + _
        "    art_cod_alt nvarchar (50) NULL , " + vbCrLf + _
        "    art_prezzo_base money NULL , " + vbCrLf + _
        "    art_scontoQ_id int NULL , " + vbCrLf + _
        "    art_giacenza_min int NULL , " + vbCrLf + _
        "    art_lotto_riordino int NULL , " + vbCrLf + _
        "    art_qta_min_ord int NULL , " + vbCrLf + _
        "    art_NovenSingola bit NOT NULL , " + vbCrLf + _
        "    art_se_accessorio bit NOT NULL , " + vbCrLf + _
        "    art_ha_accessori bit NULL , " + vbCrLf + _
        "    art_se_bundle bit NULL , " + vbCrLf + _
        "    art_se_confezione bit NULL , " + vbCrLf + _
        "    art_in_bundle bit NULL , " + vbCrLf + _
        "    art_in_confezione bit NULL , " + vbCrLf + _
        "    art_varianti bit NULL , " + vbCrLf + _
        "    art_data_insert smalldatetime NULL , " + vbCrLf + _
        "    art_data_update smalldatetime NULL , " + vbCrLf + _
        "    art_disabilitato bit NULL , " + vbCrLf + _
        "    art_tipologia_id int NULL , " + vbCrLf + _
        "    art_marca_id int NULL , " + vbCrLf + _
        "    art_iva_id int NULL , " + vbCrLf + _
        "    art_external_id int NULL , " + vbCrLf + _
        "    art_raggruppamento_id int NULL , " + vbCrLf + _
        "    art_accessori_note_de ntext NULL , " + vbCrLf + _
        "    art_accessori_note_en ntext NULL , " + vbCrLf + _
        "    art_accessori_note_es ntext NULL , " + vbCrLf + _
        "    art_accessori_note_fr ntext NULL , " + vbCrLf + _
        "    art_accessori_note_it ntext NULL , " + vbCrLf + _
        "    art_composizione_note_de ntext NULL , " + vbCrLf + _
        "    art_composizione_note_en ntext NULL , " + vbCrLf + _
        "    art_composizione_note_es ntext NULL , " + vbCrLf + _
        "    art_composizione_note_fr ntext NULL , " + vbCrLf + _
        "    art_composizione_note_it ntext NULL , " + vbCrLf + _
        "    art_descr_de ntext NULL , " + vbCrLf + _
        "    art_descr_en ntext NULL , " + vbCrLf + _
        "    art_descr_es ntext NULL , " + vbCrLf + _
        "    art_descr_fr ntext NULL , " + vbCrLf + _
        "    art_descr_it ntext NULL , " + vbCrLf + _
        "    art_note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_articoli_cod_fornitori ( " + vbCrLf + _
        "    cod_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    cod_Iart_id int NULL , " + vbCrLf + _
        "    cod_codice_articolo nvarchar (50) NULL , " + vbCrLf + _
        "    cod_codice_fornitore nvarchar (50) NULL , " + vbCrLf + _
        "    cod_fornitore_preferenziale bit NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_articoli_ordinati ( " + vbCrLf + _
        "    ao_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    ao_ut_id int NULL , " + vbCrLf + _
        "    ao_variante_id int NULL , " + vbCrLf + _
        "    ao_ranking int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_bundle ( " + vbCrLf + _
        "    bun_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    bun_quantita int NULL , " + vbCrLf + _
        "    bun_bundle_id int NULL , " + vbCrLf + _
        "    bun_articolo_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_carattech ( " + vbCrLf + _
        "    ct_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    ct_nome_it nvarchar (255) NULL , " + vbCrLf + _
        "    ct_nome_en nvarchar (255) NULL , " + vbCrLf + _
        "    ct_nome_fr nvarchar (255) NULL , " + vbCrLf + _
        "    ct_nome_es nvarchar (255) NULL , " + vbCrLf + _
        "    ct_nome_de nvarchar (255) NULL , " + vbCrLf + _
        "    ct_unita_it nvarchar (50) NULL , " + vbCrLf + _
        "    ct_unita_en nvarchar (50) NULL , " + vbCrLf + _
        "    ct_unita_fr nvarchar (50) NULL , " + vbCrLf + _
        "    ct_unita_es nvarchar (50) NULL , " + vbCrLf + _
        "    ct_unita_de nvarchar (50) NULL , " + vbCrLf + _
        "    ct_tipo int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_carichi ( " + vbCrLf + _
        "    car_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    car_magazzino_id int NULL , " + vbCrLf + _
        "    car_admin_id int NULL , " + vbCrLf + _
        "    car_fornitore nvarchar (100) NULL , " + vbCrLf + _
        "    car_fornitore_cod nvarchar (50) NULL , " + vbCrLf + _
        "    car_data smalldatetime NULL , " + vbCrLf + _
        "    car_movimentato bit NULL , " + vbCrLf + _
        "    car_note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_codici ( " + vbCrLf + _
        "    Cod_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    Cod_Codice nvarchar (50) NULL , " + vbCrLf + _
        "    Cod_variante_id int NULL , " + vbCrLf + _
        "    Cod_lista_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_criteri ( " + vbCrLf + _
        "    cri_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    cri_nome nvarchar (50) NULL , " + vbCrLf + _
        "    cri_valore nvarchar (50) NULL , " + vbCrLf + _
        "    cri_descrizione nvarchar (150) NULL , " + vbCrLf + _
        "    cri_tipo nvarchar (50) NULL , " + vbCrLf + _
        "    cri_statistica_id int NULL , " + vbCrLf + _
        "    cri_operatore nvarchar (3) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_dett_cart ( " + vbCrLf + _
        "    dett_ID int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    dett_art_var_id int NULL , " + vbCrLf + _
        "    dett_cart_id int NULL , " + vbCrLf + _
        "    dett_qta int NULL , " + vbCrLf + _
        "    dett_prezzo_unitario money NULL , " + vbCrLf + _
        "    dett_iva_id int NULL , " + vbCrLf + _
        "    dett_prezzo_listino money NULL , " + vbCrLf + _
        "    dett_sconto real NULL , " + vbCrLf + _
        "    dett_descr_IT nvarchar (500) NULL , " + vbCrLf + _
        "    dett_descr_EN nvarchar (500) NULL , " + vbCrLf + _
        "    dett_descr_FR nvarchar (500) NULL , " + vbCrLf + _
        "    dett_descr_DE nvarchar (500) NULL , " + vbCrLf + _
        "    dett_descr_ES nvarchar (500) NULL , " + vbCrLf + _
        "    dett_note nvarchar (500) NULL , " + vbCrLf + _
        "    dett_cod_promozione nvarchar (50) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_dett_cart_dest ( " + vbCrLf + _
        "    dd_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    dd_qta int NULL , " + vbCrLf + _
        "    dd_dett_id int NULL , " + vbCrLf + _
        "    dd_ind_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_dett_cart_proposte ( " + vbCrLf + _
        "    dp_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    dp_qta int NULL , " + vbCrLf + _
        "    dp_dett_id int NULL , " + vbCrLf + _
        "    dp_ut_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_dettagli_ord ( " + vbCrLf + _
        "    Det_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    det_ord_id int NULL , " + vbCrLf + _
        "    det_ind_id int NULL , " + vbCrLf + _
        "    det_art_var_id int NULL , " + vbCrLf + _
        "    det_qta float NULL , " + vbCrLf + _
        "    det_prezzo_unitario money NULL , " + vbCrLf + _
        "    det_iva real NULL , " + vbCrLf + _
        "    det_prezzo_listino money NULL , " + vbCrLf + _
        "    det_sconto real NULL , " + vbCrLf + _
        "    det_descr_IT nvarchar (500) NULL , " + vbCrLf + _
        "    det_descr_EN nvarchar (500) NULL , " + vbCrLf + _
        "    det_descr_FR nvarchar (500) NULL , " + vbCrLf + _
        "    det_descr_DE nvarchar (500) NULL , " + vbCrLf + _
        "    det_descr_ES nvarchar (500) NULL , " + vbCrLf + _
        "    det_note nvarchar (500) NULL , " + vbCrLf + _
        "    det_cod_promozione nvarchar (50) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_dettagli_ord_utenti ( " + vbCrLf + _
        "    du_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    du_qta int NULL , " + vbCrLf + _
        "    du_det_id int NULL , " + vbCrLf + _
        "    du_ut_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_iva ( " + vbCrLf + _
        "    iva_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    iva_nome nvarchar (250) NULL , " + vbCrLf + _
        "    iva_valore real NULL , " + vbCrLf + _
        "    iva_ordine int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_lista_codici ( " + vbCrLf + _
        "    lstCod_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    LstCod_nome nvarchar (50) NULL , " + vbCrLf + _
        "    LstCod_cod nvarchar (50) NULL , " + vbCrLf + _
        "    lstCod_sistema bit NULL , " + vbCrLf + _
        "    lstCod_note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_listini ( " + vbCrLf + _
        "    listino_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    listino_codice nvarchar (50) NULL , " + vbCrLf + _
        "    listino_datacreazione smalldatetime NULL , " + vbCrLf + _
        "    listino_datascadenza smalldatetime NULL , " + vbCrLf + _
        "    listino_B2C bit NULL , " + vbCrLf + _
        "    listino_offerte bit NULL , " + vbCrLf + _
        "    listino_base bit NULL , " + vbCrLf + _
        "    listino_base_attuale bit NULL , " + vbCrLf + _
        "    listino_ancestor_id int NULL , " + vbCrLf + _
        "    listino_with_child bit NOT NULL , " + vbCrLf + _
        "    listino_note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_magazzini ( " + vbCrLf + _
        "    mag_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    mag_nome nvarchar (50) NULL , " + vbCrLf + _
        "    mag_vendita_pubblico bit NULL , " + vbCrLf + _
        "    mag_disponibilita bit NULL , " + vbCrLf + _
        "    mag_codice nvarchar (50) NULL , " + vbCrLf + _
        "    mag_note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_marche ( " + vbCrLf + _
        "    mar_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    mar_nome_it nvarchar (50) NULL , " + vbCrLf + _
        "    mar_nome_en nvarchar (50) NULL , " + vbCrLf + _
        "    mar_nome_fr nvarchar (50) NULL , " + vbCrLf + _
        "    mar_nome_es nvarchar (50) NULL , " + vbCrLf + _
        "    mar_nome_de nvarchar (50) NULL , " + vbCrLf + _
        "    mar_logo nvarchar (255) NULL , " + vbCrLf + _
        "    mar_sito nvarchar (255) NULL , " + vbCrLf + _
        "    mar_codice nvarchar (20) NULL , " + vbCrLf + _
        "    mar_generica bit NULL , " + vbCrLf + _
        "    mar_descr_de ntext NULL , " + vbCrLf + _
        "    mar_descr_en ntext NULL , " + vbCrLf + _
        "    mar_descr_es ntext NULL , " + vbCrLf + _
        "    mar_descr_fr ntext NULL , " + vbCrLf + _
        "    mar_descr_it ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_movimenti ( " + vbCrLf + _
        "    mov_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    mov_sorg_id int NULL , " + vbCrLf + _
        "    mov_dest_id int NULL , " + vbCrLf + _
        "    mov_admin_id int NULL , " + vbCrLf + _
        "    mov_data smalldatetime NULL , " + vbCrLf + _
        "    mov_data_evasione smalldatetime NULL , " + vbCrLf + _
        "    mov_codice nvarchar (50) NULL , " + vbCrLf + _
        "    mov_note_evasione ntext NULL , " + vbCrLf + _
        "    mov_note_richiesta ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_ordini ( " + vbCrLf + _
        "    ord_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    ord_riv_id int NULL , " + vbCrLf + _
        "    ord_data smalldatetime NULL , " + vbCrLf + _
        "    ord_stato_id int NULL , " + vbCrLf + _
        "    ord_magazzino_id int NULL , " + vbCrLf + _
        "    ord_data_ins smalldatetime NULL , " + vbCrLf + _
        "    ord_data_ultima_mod smalldatetime NULL , " + vbCrLf + _
        "    ord_cod nvarchar (50) NULL , " + vbCrLf + _
        "    ord_movimenta bit NULL , " + vbCrLf + _
        "    ord_impegna bit NULL , " + vbCrLf + _
        "    ord_archiviato bit NULL , " + vbCrLf + _
        "    ord_Note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_prezzi ( " + vbCrLf + _
        "    prz_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    prz_prezzo money NULL , " + vbCrLf + _
        "    prz_visibile bit NOT NULL , " + vbCrLf + _
        "    prz_promozione bit NULL , " + vbCrLf + _
        "    prz_listino_id int NULL , " + vbCrLf + _
        "    prz_variante_id int NULL , " + vbCrLf + _
        "    prz_scontoQ_id int NULL , " + vbCrLf + _
        "    prz_iva_id int NULL , " + vbCrLf + _
        "    prz_var_euro real NULL , " + vbCrLf + _
        "    prz_var_sconto real NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_rivenditori ( " + vbCrLf + _
        "    riv_id int NOT NULL , " + vbCrLf + _
        "    riv_listino_id int NULL , " + vbCrLf + _
        "    riv_lstcod_id int NULL , " + vbCrLf + _
        "    riv_valuta_id int NULL , " + vbCrLf + _
        "    riv_agente_id int NULL , " + vbCrLf + _
        "    riv_codice nvarchar (20) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_scontiQ ( " + vbCrLf + _
        "    sco_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    sco_qta_da int NULL , " + vbCrLf + _
        "    sco_sconto real NULL , " + vbCrLf + _
        "    sco_classe_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_scontiQ_classi ( " + vbCrLf + _
        "    scc_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    scc_nome nvarchar (255) NULL , " + vbCrLf + _
        "    scc_note ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_shopping_cart ( " + vbCrLf + _
        "    sc_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    sc_date_cart smalldatetime NULL , " + vbCrLf + _
        "    sc_ut_id int NOT NULL , " + vbCrLf + _
        "    sc_riv_id int NOT NULL , " + vbCrLf + _
        "    sc_Sospeso bit NOT NULL , " + vbCrLf + _
        "    sc_NomeCart nvarchar (50) NULL , " + vbCrLf + _
        "    sc_completato bit NULL , " + vbCrLf + _
        "    sc_giorni_validita_preventivo int NULL , " + vbCrLf + _
        "    sc_NoteCart ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_stati_ordine ( " + vbCrLf + _
        "    so_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    so_nome nvarchar (100) NULL , " + vbCrLf + _
        "    so_ordine int NULL , " + vbCrLf + _
        "    so_stato_ordini int NULL , " + vbCrLf + _
        "    so_internet bit NULL , " + vbCrLf + _
        "    so_descrizione ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_statistiche ( " + vbCrLf + _
        "    sta_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    sta_nome nvarchar (50) NULL , " + vbCrLf + _
        "    sta_dataC datetime NULL , " + vbCrLf + _
        "    sta_temp bit NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_tip_ctech ( " + vbCrLf + _
        "    rct_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rct_ctech_id int NULL , " + vbCrLf + _
        "    rct_ordine int NULL , " + vbCrLf + _
        "    rct_tipologia_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_tipologie ( " + vbCrLf + _
        "    tip_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    tip_nome_it nvarchar (250) NULL , " + vbCrLf + _
        "    tip_nome_en nvarchar (250) NULL , " + vbCrLf + _
        "    tip_nome_fr nvarchar (250) NULL , " + vbCrLf + _
        "    tip_nome_es nvarchar (250) NULL , " + vbCrLf + _
        "    tip_nome_de nvarchar (250) NULL , " + vbCrLf + _
        "    tip_logo nvarchar (255) NULL , " + vbCrLf + _
        "    tip_foto nvarchar (255) NULL , " + vbCrLf + _
        "    tip_codice nvarchar (50) NULL , " + vbCrLf + _
        "    tip_foglia bit NULL , " + vbCrLf + _
        "    tip_livello int NULL , " + vbCrLf + _
        "    tip_padre_id int NULL , " + vbCrLf + _
        "    tip_ordine int NULL , " + vbCrLf + _
        "    tip_ordine_assoluto nvarchar (250) NULL , " + vbCrLf + _
        "    tip_external_id nvarchar (50) NULL , " + vbCrLf + _
        "    tip_tipologia_padre_base int NULL , " + vbCrLf + _
        "    tip_visibile bit NOT NULL , " + vbCrLf + _
        "    tip_albero_visibile bit NOT NULL , " + vbCrLf + _
        "    tip_descr_de ntext NULL , " + vbCrLf + _
        "    tip_descr_en ntext NULL , " + vbCrLf + _
        "    tip_descr_es ntext NULL , " + vbCrLf + _
        "    tip_descr_fr ntext NULL , " + vbCrLf + _
        "    tip_descr_it ntext NULL , " + vbCrLf + _
        "    tip_tipologie_padre_lista nvarchar (255) NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_tipologie_raggruppamenti ( " + vbCrLf + _
        "    rag_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    rag_nome_it nvarchar (250) NULL , " + vbCrLf + _
        "    rag_nome_en nvarchar (250) NULL , " + vbCrLf + _
        "    rag_nome_fr nvarchar (250) NULL , " + vbCrLf + _
        "    rag_nome_es nvarchar (250) NULL , " + vbCrLf + _
        "    rag_nome_de nvarchar (250) NULL , " + vbCrLf + _
        "    rag_foto nvarchar (255) NULL , " + vbCrLf + _
        "    rag_ordine int NOT NULL , " + vbCrLf + _
        "    rag_tipologia_id int NULL , " + vbCrLf + _
        "    rag_visibile bit NOT NULL , " + vbCrLf + _
        "    rag_descr_de ntext NULL , " + vbCrLf + _
        "    rag_descr_en ntext NULL , " + vbCrLf + _
        "    rag_descr_es ntext NULL , " + vbCrLf + _
        "    rag_descr_fr ntext NULL , " + vbCrLf + _
        "    rag_descr_it ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_valori ( " + vbCrLf + _
        "    val_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    val_nome_it nvarchar (50) NULL , " + vbCrLf + _
        "    val_nome_en nvarchar (50) NULL , " + vbCrLf + _
        "    val_nome_fr nvarchar (50) NULL , " + vbCrLf + _
        "    val_nome_es nvarchar (50) NULL , " + vbCrLf + _
        "    val_nome_de nvarchar (50) NULL , " + vbCrLf + _
        "    val_icona nvarchar (255) NULL , " + vbCrLf + _
        "    val_cod_int nvarchar (50) NULL , " + vbCrLf + _
        "    val_cod_pro nvarchar (50) NULL , " + vbCrLf + _
        "    val_cod_alt nvarchar (50) NULL , " + vbCrLf + _
        "    val_var_id int NULL , " + vbCrLf + _
        "    val_ordine int NULL , " + vbCrLf + _
        "    val_descr_de ntext NULL , " + vbCrLf + _
        "    val_descr_en ntext NULL , " + vbCrLf + _
        "    val_descr_es ntext NULL , " + vbCrLf + _
        "    val_descr_fr ntext NULL , " + vbCrLf + _
        "    val_descr_it ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_valute ( " + vbCrLf + _
        "    valu_ID int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    Valu_nome nvarchar (50) NULL , " + vbCrLf + _
        "    valu_codice nvarchar (3) NULL , " + vbCrLf + _
        "    Valu_Cambio money NULL , " + vbCrLf + _
        "    valu_simbolo nvarchar (5) NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_varianti ( " + vbCrLf + _
        "    var_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    var_nome_it nvarchar (250) NULL , " + vbCrLf + _
        "    var_nome_en nvarchar (250) NULL , " + vbCrLf + _
        "    var_nome_fr nvarchar (250) NULL , " + vbCrLf + _
        "    var_nome_es nvarchar (250) NULL , " + vbCrLf + _
        "    var_nome_de nvarchar (250) NULL , " + vbCrLf + _
        "    var_ordine int NULL , " + vbCrLf + _
        "    var_descr_de ntext NULL , " + vbCrLf + _
        "    var_descr_en ntext NULL , " + vbCrLf + _
        "    var_descr_es ntext NULL , " + vbCrLf + _
        "    var_descr_fr ntext NULL , " + vbCrLf + _
        "    var_descr_it ntext NULL )          " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE TABLE dbo.gtb_wish_list ( " + vbCrLf + _
        "    wish_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
        "    wish_ut_id int NULL , " + vbCrLf + _
        "    wish_variante_id int NULL ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gItb_articoli WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gItb_articoli PRIMARY KEY ( Iart_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.glog_ordini WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_glog_ordini PRIMARY KEY ( log_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.glog_ordini_operazioni WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_glog_ordini_operazioni PRIMARY KEY ( op_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_acc WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_art_acc PRIMARY KEY ( aa_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_ctech WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_art_ctech PRIMARY KEY ( rel_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_valori WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_art_valori PRIMARY KEY ( rel_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_vv WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_art_vv PRIMARY KEY ( rvv_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_carichi_var WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_carichi_var PRIMARY KEY ( rcv_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_giacenze WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_giacenze PRIMARY KEY ( gia_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_mov_var WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_mov_var PRIMARY KEY ( rmv_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_accessori_tipo WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_accessori_tipo PRIMARY KEY ( at_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_agenti WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_agenti PRIMARY KEY ( ag_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_art_foto WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_grel_art_foto PRIMARY KEY ( fo_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_articoli WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_articoli PRIMARY KEY ( art_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_articoli_cod_fornitori WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gItb_articoli_cod_fornitori PRIMARY KEY ( cod_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_articoli_ordinati WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_articoli_ordinati PRIMARY KEY ( ao_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_bundle WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_bundle PRIMARY KEY ( bun_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_carattech WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_carattech PRIMARY KEY ( ct_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_carichi WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_carichi PRIMARY KEY ( car_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_codici WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_codici PRIMARY KEY ( Cod_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_criteri WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_criteri PRIMARY KEY ( cri_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dett_cart WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_dett_cart PRIMARY KEY ( dett_ID ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dett_cart_proposte WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_dett_cart_proposte PRIMARY KEY ( dp_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dettagli_ord WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_dettagli_ord PRIMARY KEY ( Det_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dettagli_ord_utenti WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_dettagli_ord_utenti PRIMARY KEY ( du_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_iva WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_iva PRIMARY KEY ( iva_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_lista_codici WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_lista_codici PRIMARY KEY ( lstCod_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_listini WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_listini PRIMARY KEY ( listino_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_magazzini WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_magazzini PRIMARY KEY ( mag_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_marche WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_marche PRIMARY KEY ( mar_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_movimenti WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_movimenti PRIMARY KEY ( mov_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_ordini WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_ordini PRIMARY KEY ( ord_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_prezzi WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_prezzi PRIMARY KEY ( prz_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_rivenditori WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_rivenditori PRIMARY KEY ( riv_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_scontiQ WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_scontiQ PRIMARY KEY ( sco_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_scontiQ_classi WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_scontiQ_classi PRIMARY KEY ( scc_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_shopping_cart WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_shopping_cart PRIMARY KEY ( sc_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_stati_ordine WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_tipi_ordine PRIMARY KEY ( so_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_statistiche WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_statistiche PRIMARY KEY ( sta_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_tipologie WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_linee PRIMARY KEY ( tip_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_tipologie_raggruppamenti WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_tipologie_raggruppamenti PRIMARY KEY ( rag_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_valori WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_valori PRIMARY KEY ( val_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_valute WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_valute PRIMARY KEY ( valu_ID ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_varianti WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_tb_varianti PRIMARY KEY ( var_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_wish_list WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_wish_list PRIMARY KEY ( wish_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__grel_art_valori__rel_art_id ON dbo.grel_art_valori(rel_art_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__grel_art_vv__rvv_art_var_id ON dbo.grel_art_vv(rvv_art_var_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__grel_giacenze__gia_art_var_id ON dbo.grel_giacenze(gia_art_var_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__gtb_articoli__art_tipologia_id ON dbo.gtb_articoli(art_tipologia_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__gtb_articoli_ordinati ON dbo.gtb_articoli_ordinati(ao_ut_id, ao_variante_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__gtb_codici ON dbo.gtb_codici(Cod_lista_id, Cod_variante_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__gtb_prezzi ON dbo.gtb_prezzi(prz_listino_id, prz_variante_id) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE  INDEX IDX__gtb_wish_list ON dbo.gtb_wish_list(wish_ut_id, wish_variante_id) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.glog_ordini ADD  " + vbCrLf + _
        "    CONSTRAINT FK_glog_ordini__glog_ordini_operazioni FOREIGN KEY ( log_operazione_id ) " + vbCrLf + _
        "    REFERENCES dbo.glog_ordini_operazioni ( op_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_glog_ordini_gtb_ordini FOREIGN KEY ( log_ordine_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_ordini ( ord_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_acc ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_acc_gtb_accessori_tipo FOREIGN KEY ( aa_tipo_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_accessori_tipo ( at_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_acc_gtb_articoli FOREIGN KEY ( aa_art_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_articoli ( art_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_acc_gtb_articoli1 FOREIGN KEY ( aa_acc_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_articoli ( art_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.grel_art_acc nocheck constraint FK_grel_art_acc_gtb_articoli1 " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_ctech ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_ctech_gtb_articoli FOREIGN KEY ( rel_art_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_articoli ( art_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_ctech_gtb_carattech FOREIGN KEY ( rel_ctech_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_carattech ( ct_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_valori ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_valori__gtb_art_foto FOREIGN KEY ( rel_foto_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_art_foto ( fo_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_valori_gItb_articoli FOREIGN KEY ( rel_external_id ) " + vbCrLf + _
        "    REFERENCES dbo.gItb_articoli ( Iart_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_valori_gtb_articoli FOREIGN KEY ( rel_art_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_articoli ( art_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_valori_gtb_scontiQ_classi FOREIGN KEY ( rel_scontoQ_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_scontiQ_classi ( scc_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.grel_art_valori nocheck constraint FK_grel_art_valori__gtb_art_foto " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.grel_art_valori nocheck constraint FK_grel_art_valori_gItb_articoli " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.grel_art_valori nocheck constraint FK_grel_art_valori_gtb_scontiQ_classi " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_art_vv ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_vv_grel_art_valori FOREIGN KEY ( rvv_art_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_vv_gtb_valori FOREIGN KEY ( rvv_val_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_valori ( val_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_carichi_var ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_carichi_var_grel_art_valori FOREIGN KEY ( rcv_art_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_carichi_var_gtb_carichi FOREIGN KEY ( rcv_car_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_carichi ( car_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_giacenze ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_giacenze_grel_art_valori FOREIGN KEY ( gia_art_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_giacenze_gtb_magazzini FOREIGN KEY ( gia_magazzino_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_magazzini ( mag_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.grel_mov_var ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_mov_var_grel_art_valori FOREIGN KEY ( rmv_art_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_grel_mov_var_gtb_movimenti FOREIGN KEY ( rmv_mov_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_movimenti ( mov_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_agenti ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_agenti_tb_admin FOREIGN KEY ( ag_admin_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_admin ( ID_admin ), " + vbCrLf + _
        "    CONSTRAINT FK_gtb_agenti_tb_gruppi FOREIGN KEY ( ag_gruppo_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_gruppi ( id_Gruppo ), " + vbCrLf + _
        "    CONSTRAINT FK_gtb_agenti_tb_Utenti FOREIGN KEY ( ag_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Utenti ( ut_ID ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_agenti nocheck constraint FK_gtb_agenti_tb_admin " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_agenti nocheck constraint FK_gtb_agenti_tb_gruppi " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_art_foto ADD  " + vbCrLf + _
        "    CONSTRAINT FK_grel_art_foto_gtb_articoli1 FOREIGN KEY ( fo_articolo_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_articoli ( art_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_articoli ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_gItb_articoli FOREIGN KEY ( art_external_id ) " + vbCrLf + _
        "    REFERENCES dbo.gItb_articoli ( Iart_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_gtb_iva FOREIGN KEY ( art_iva_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_iva ( iva_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_gtb_marche FOREIGN KEY ( art_marca_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_marche ( mar_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_gtb_scontiQ_classi FOREIGN KEY ( art_scontoQ_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_scontiQ_classi ( scc_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_gtb_tipologie FOREIGN KEY ( art_tipologia_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_tipologie ( tip_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_gtb_tipologie_raggruppamenti FOREIGN KEY ( art_raggruppamento_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_tipologie_raggruppamenti ( rag_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_articoli nocheck constraint FK_gtb_articoli_gItb_articoli " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_articoli nocheck constraint FK_gtb_articoli_gtb_scontiQ_classi " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_articoli nocheck constraint FK_gtb_articoli_gtb_tipologie_raggruppamenti " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_articoli_cod_fornitori ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gItb_articoli_cod_fornitori_gItb_articoli FOREIGN KEY ( cod_Iart_id ) " + vbCrLf + _
        "    REFERENCES dbo.gItb_articoli ( Iart_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_articoli_ordinati ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_ordinati_grel_art_valori FOREIGN KEY ( ao_variante_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_articoli_ordinati_tb_Utenti FOREIGN KEY ( ao_ut_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Utenti ( ut_ID ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_bundle ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_bundle_grel_art_valori FOREIGN KEY ( bun_bundle_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_bundle_grel_art_valori1 FOREIGN KEY ( bun_articolo_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_bundle nocheck constraint FK_gtb_bundle_grel_art_valori1 " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_carichi ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_carichi_gtb_magazzini FOREIGN KEY ( car_magazzino_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_magazzini ( mag_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_carichi_tb_admin FOREIGN KEY ( car_admin_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_admin ( ID_admin ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_codici ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_codici_grel_art_valori1 FOREIGN KEY ( Cod_variante_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_codici_gtb_lista_codici1 FOREIGN KEY ( Cod_lista_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_lista_codici ( lstCod_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_criteri ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_criteri_gtb_statistiche FOREIGN KEY ( cri_statistica_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_statistiche ( sta_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dett_cart ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_grel_art_valori FOREIGN KEY ( dett_art_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_gtb_iva FOREIGN KEY ( dett_iva_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_iva ( iva_id ), " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_gtb_shopping_cart FOREIGN KEY ( dett_cart_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_shopping_cart ( sc_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_dett_cart nocheck constraint FK_gtb_dett_cart_grel_art_valori " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dett_cart_dest ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_dest_gtb_dett_cart FOREIGN KEY ( dd_dett_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_dett_cart ( dett_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_dest_tb_Indirizzario FOREIGN KEY ( dd_ind_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Indirizzario ( IDElencoIndirizzi ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dett_cart_proposte ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_proposte_gtb_dett_cart FOREIGN KEY ( dp_dett_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_dett_cart ( dett_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dett_cart_proposte_tb_Utenti FOREIGN KEY ( dp_ut_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Utenti ( ut_ID ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dettagli_ord ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dettagli_ord_grel_art_valori FOREIGN KEY ( det_art_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dettagli_ord_gtb_ordini FOREIGN KEY ( det_ord_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_ordini ( ord_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dettagli_ord_tb_Indirizzario FOREIGN KEY ( det_ind_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Indirizzario ( IDElencoIndirizzi ) " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_dettagli_ord nocheck constraint FK_gtb_dettagli_ord_grel_art_valori " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_dettagli_ord_utenti ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dettagli_ord_utenti_gtb_dettagli_ord FOREIGN KEY ( du_det_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_dettagli_ord ( Det_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_dettagli_ord_utenti_tb_Utenti FOREIGN KEY ( du_ut_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Utenti ( ut_ID ) " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_dettagli_ord_utenti nocheck constraint FK_gtb_dettagli_ord_utenti_tb_Utenti " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_listini ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_listini_gtb_listini_ancestor FOREIGN KEY ( listino_ancestor_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_listini ( listino_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_listini nocheck constraint FK_gtb_listini_gtb_listini_ancestor " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_movimenti ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_movimenti_gtb_magazzini FOREIGN KEY ( mov_sorg_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_magazzini ( mag_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_movimenti_gtb_magazzini1 FOREIGN KEY ( mov_dest_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_magazzini ( mag_id ), " + vbCrLf + _
        "    CONSTRAINT FK_gtb_movimenti_tb_admin FOREIGN KEY ( mov_admin_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_admin ( ID_admin ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_movimenti nocheck constraint FK_gtb_movimenti_gtb_magazzini1 " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_ordini ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_ordini_gtb_magazzini FOREIGN KEY ( ord_magazzino_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_magazzini ( mag_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_ordini_gtb_rivenditori FOREIGN KEY ( ord_riv_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_rivenditori ( riv_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_ordini_gtb_stati_ordine FOREIGN KEY ( ord_stato_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_stati_ordine ( so_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_prezzi ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_prezzi_grel_art_valori FOREIGN KEY ( prz_variante_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_prezzi_gtb_iva FOREIGN KEY ( prz_iva_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_iva ( iva_id ), " + vbCrLf + _
        "    CONSTRAINT FK_gtb_prezzi_gtb_listini FOREIGN KEY ( prz_listino_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_listini ( listino_id ) ON DELETE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_prezzi_gtb_scontiQ_classi FOREIGN KEY ( prz_scontoQ_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_scontiQ_classi ( scc_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_prezzi nocheck constraint FK_gtb_prezzi_gtb_scontiQ_classi " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_rivenditori ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_rivenditori_gtb_agenti FOREIGN KEY ( riv_agente_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_agenti ( ag_id ), " + vbCrLf + _
        "    CONSTRAINT FK_gtb_rivenditori_gtb_lista_codici1 FOREIGN KEY ( riv_lstcod_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_lista_codici ( lstCod_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_rivenditori_gtb_listini1 FOREIGN KEY ( riv_listino_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_listini ( listino_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_rivenditori_gtb_valute1 FOREIGN KEY ( riv_valuta_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_valute ( valu_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_rivenditori_tb_Utenti FOREIGN KEY ( riv_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Utenti ( ut_ID ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_rivenditori nocheck constraint FK_gtb_rivenditori_gtb_agenti " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_rivenditori nocheck constraint FK_gtb_rivenditori_gtb_lista_codici1 " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_scontiQ ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_scontiQ_gtb_scontiQ_classi FOREIGN KEY ( sco_classe_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_scontiQ_classi ( scc_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_shopping_cart ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_shopping_cart__gtb_rivenditori FOREIGN KEY ( sc_riv_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_rivenditori ( riv_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_tip_ctech ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_tip_ctech_gtb_carattech FOREIGN KEY ( rct_ctech_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_carattech ( ct_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_tip_ctech_gtb_tipologie FOREIGN KEY ( rct_tipologia_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_tipologie ( tip_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_tipologie ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_tipologie_gtb_tipologie_padre_base FOREIGN KEY ( tip_tipologia_padre_base ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_tipologie ( tip_id ) NOT FOR REPLICATION , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_tipologie_gtb_tipologie1 FOREIGN KEY ( tip_padre_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_tipologie ( tip_id ) NOT FOR REPLICATION  " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_tipologie nocheck constraint FK_gtb_tipologie_gtb_tipologie_padre_base " + vbCrLf + _
        ";" + vbCrLf + _
        "alter table dbo.gtb_tipologie nocheck constraint FK_gtb_tipologie_gtb_tipologie1 " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_tipologie_raggruppamenti ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_tipologie_raggruppamenti_gtb_tipologie FOREIGN KEY ( rag_tipologia_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_tipologie ( tip_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_valori ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_valori_gtb_varianti1 FOREIGN KEY ( val_var_id ) " + vbCrLf + _
        "    REFERENCES dbo.gtb_varianti ( var_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "ALTER TABLE dbo.gtb_wish_list ADD  " + vbCrLf + _
        "    CONSTRAINT FK_gtb_wish_list_grel_art_valori FOREIGN KEY ( wish_variante_id ) " + vbCrLf + _
        "    REFERENCES dbo.grel_art_valori ( rel_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
        "    CONSTRAINT FK_gtb_wish_list_tb_Utenti FOREIGN KEY ( wish_ut_id ) " + vbCrLf + _
        "    REFERENCES dbo.tb_Utenti ( ut_ID ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_agenti AS  " + vbCrLf + _
        "    SELECT * FROM gtb_agenti  " + vbCrLf + _
        "        INNER JOIN tb_admin ON gtb_agenti.ag_admin_id = tb_admin.ID_admin  " + vbCrLf + _
        "        INNER JOIN tb_Utenti ON gtb_agenti.ag_id = tb_Utenti.ut_ID  " + vbCrLf + _
        "        INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_articoli AS " + vbCrLf + _
        "    SELECT * FROM gtb_articoli " + vbCrLf + _
        "        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLf + _
        "        INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + vbCrLf + _
        "        INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id  " + vbCrLf + _
        "        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_rivenditori AS  " + vbCrLf + _
        "    SELECT * FROM gtb_rivenditori  " + vbCrLf + _
        "        INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_utenti.ut_ID  " + vbCrLf + _
        "        INNER JOIN tb_Indirizzario ON tb_utenti.ut_NextCom_ID = tb_indirizzario.IDElencoIndirizzi  " + vbCrLf + _
        "        INNER JOIN gtb_valute ON gtb_rivenditori.riv_valuta_id = gtb_valute.valu_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_Giacenze_Varianti AS  " + vbCrLf + _
        "    SELECT * FROM grel_giacenze INNER JOIN  " + vbCrLf + _
        "        grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id INNER JOIN  " + vbCrLf + _
        "        grel_carichi_var ON grel_art_valori.rel_id = grel_carichi_var.rcv_art_var_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_articoli_varianti AS  " + vbCrLf + _
        "    SELECT TOP 100 PERCENT gtb_valori.*, grel_art_vv.*, gtb_varianti.*  " + vbCrLf + _
        "        FROM grel_art_vv INNER JOIN gtb_valori ON grel_art_vv.rvv_val_id = gtb_valori.val_id  " + vbCrLf + _
        "        INNER JOIN gtb_varianti ON gtb_valori.val_var_id = gtb_varianti.var_id  " + vbCrLf + _
        "        ORDER BY gtb_varianti.var_ordine, gtb_varianti.var_id, gtb_valori.val_ordine " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_carichi AS  " + vbCrLf + _
        "    SELECT * FROM grel_carichi_var  " + vbCrLf + _
        "        INNER JOIN grel_art_valori ON grel_carichi_var.rcv_art_var_id = grel_art_valori.rel_id " + vbCrLf + _
        "        INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLf + _
        "        INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id " + vbCrLf + _
        "        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_inventario AS  " + vbCrLf + _
        "    SELECT * FROM grel_giacenze  " + vbCrLf + _
        "    INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id  " + vbCrLf + _
        "    INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_listini AS " + vbCrLf + _
        "    SELECT * FROM grel_art_valori " + vbCrLf + _
        "        INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLf + _
        "        INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLf + _
        "        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLf + _
        "        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_listino_offerte AS " + vbCrLf + _
        "    SELECT * FROM gtb_articoli " + vbCrLf + _
        "        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLf + _
        "        INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLf + _
        "        INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLf + _
        "        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLf + _
        "        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLf + _
        "    WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLf + _
        "        AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLf + _
        "        AND tip_visibile=1 " + vbCrLf + _
        "        AND tip_albero_visibile=1 " + vbCrLf + _
        "        AND ISNULL(listino_offerte, 0)=1 " + vbCrLf + _
        "        AND ISNULL(prz_visibile, 0)=1 " + vbCrLf + _
        "        AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_listino_vendita AS  " + vbCrLf + _
        "        SELECT * FROM gtb_articoli  " + vbCrLf + _
        "        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id  " + vbCrLf + _
        "            INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id  " + vbCrLf + _
        "            INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id  " + vbCrLf + _
        "        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLf + _
        "        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id  " + vbCrLf + _
        "            WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0  " + vbCrLf + _
        "                  AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0  " + vbCrLf + _
        "                  AND tip_visibile=1  " + vbCrLf + _
        "                  AND tip_albero_visibile=1  " + vbCrLf + _
        "                  AND prz_visibile=1  " + vbCrLf + _
        "                  AND ( ( ISNULL(listino_offerte, 0)=1  " + vbCrLf + _
        "                            AND ISNULL(prz_visibile, 0)=1  " + vbCrLf + _
        "                          AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)  " + vbCrLf + _
        "                        )  " + vbCrLf + _
        "                        OR (ISNULL(listino_offerte, 0)=0  " + vbCrLf + _
        "                  AND prz_variante_id NOT IN (  " + vbCrLf + _
        "                        SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + vbCrLf + _
        "                        WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1  " + vbCrLf + _
        "                        AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)  " + vbCrLf + _
        "                                             )  " + vbCrLf + _
        "                        )  " + vbCrLf + _
        "                      ) " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_dettagli_ord AS  " + vbCrLf + _
        "    SELECT * FROM gtb_dettagli_ord  " + vbCrLf + _
        "        LEFT JOIN grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id  " + vbCrLf + _
        "        LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLf + _
        ";" + vbCrLf + _
        "CREATE VIEW dbo.gv_CartDetail AS  " + vbCrLf + _
        "    SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST,  " + vbCrLf + _
        "              (SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT  " + vbCrLf + _
        "        FROM gtb_dett_cart LEFT JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id  " + vbCrLf + _
        "        LEFT JOIN grel_art_valori ON gtb_dett_cart.dett_art_var_id = grel_art_valori.rel_id  " + vbCrLf + _
        "        LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id  " + vbCrLf + _
        "        LEFT JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLf + _
        "    WHERE (gtb_dett_cart.dett_art_var_id IS NULL) OR  " + vbCrLf + _
        "        ( ISNULL(gtb_articoli.art_disabilitato, 0)=0 AND  " + vbCrLf + _
        "          ISNULL(grel_art_valori.rel_disabilitato,0)=0 AND  " + vbCrLf + _
        "          ISNULL(gtb_tipologie.tip_albero_visibile, 0) = 1 AND  " + vbCrLf + _
        "          ISNULL(gtb_tipologie.tip_visibile, 0)= 1 ) " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE TRIGGER dbo.TRG_gtb_prezzi_FOR_DELETE  " + vbCrLf + _
        "        ON gtb_prezzi  " + vbCrLf + _
        "        FOR DELETE  " + vbCrLf + _
        "    AS  " + vbCrLf + _
        "        DELETE FROM gtb_prezzi WHERE prz_id IN (  " + vbCrLf + _
        "                SELECT gtb_prezzi.prz_id FROM (gtb_prezzi INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id=gtb_listini.listino_id)  " + vbCrLf + _
        "                INNER JOIN deleted ON (gtb_prezzi.prz_variante_id = deleted.prz_variante_id AND gtb_listini.listino_ancestor_id=deleted.prz_listino_id)  " + vbCrLf + _
        "                WHERE gtb_prezzi.prz_prezzo = deleted.prz_prezzo AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_visibile = deleted.prz_visibile AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_promozione = deleted.prz_promozione AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_variante_id = deleted.prz_variante_id AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_scontoQ_id = deleted.prz_scontoQ_id AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_iva_id = deleted.prz_iva_id AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_var_euro = deleted.prz_var_euro AND  " + vbCrLf + _
        "                      gtb_prezzi.prz_var_sconto = deleted.prz_var_sconto)  " + vbCrLf + _
        ";" + vbCrLf + _
        " CREATE TRIGGER dbo.TRG_gtb_prezzi_FOR_INSERT  " + vbCrLf + _
        "     ON gtb_prezzi  " + vbCrLf + _
        "     FOR INSERT NOT FOR REPLICATION  " + vbCrLf + _
        " AS  " + vbCrLf + _
        "     INSERT INTO gtb_prezzi (prz_prezzo, prz_visibile, prz_promozione, prz_listino_id, prz_variante_id, prz_scontoQ_id, prz_iva_id, prz_var_euro, prz_var_sconto)  " + vbCrLf + _
        "     SELECT inserted.prz_prezzo, inserted.prz_visibile, inserted.prz_promozione, L_child.listino_id, inserted.prz_variante_id,  " + vbCrLf + _
        "            inserted.prz_scontoQ_id, inserted.prz_iva_id, inserted.prz_var_euro, inserted.prz_var_sconto   " + vbCrLf + _
        "         FROM inserted  " + vbCrLf + _
        "         INNER JOIN gtb_listini L_ancestor ON inserted.prz_listino_id = L_ancestor.listino_id  " + vbCrLf + _
        "         INNER JOIN gtb_listini L_child ON L_ancestor.listino_id = L_child.listino_ancestor_id  " + vbCrLf + _
        "         WHERE (SELECT COUNT(*)   " + vbCrLf + _
        "                        FROM gtb_prezzi   " + vbCrLf + _
        "                     WHERE gtb_prezzi.prz_listino_id=L_child.listino_id AND   " + vbCrLf + _
        "                           gtb_prezzi.prz_variante_id=inserted.prz_variante_id)=0  " + vbCrLf + _
        ";" + vbCrLf
        
        
        'codice di inserimento automatico dati delle tabelle necessarie per far funzionare il b2b
        Install_B2B__version_79 = Install_B2B__version_79 + _
            " INSERT INTO tb_siti(id_sito, sito_nome, sito_dir, sito_amministrazione, sito_p1) " + _
			"   VALUES (" & NEXTB2B & ", 'NEXT-b2b [gestione prodotti, magazzino e vendita]', 'NextB2b', 1, 'B2B_ADMIN'); " + _
            " INSERT INTO gtb_accessori_tipo (at_nome_it, at_ordine, at_vincolo_vendita) VALUES ('Accessori', 1, 1) ; " + _
            " INSERT INTO gtb_accessori_tipo (at_nome_it, at_ordine, at_vincolo_vendita) VALUES ('Prodotti correlati', 2, 0) ; " + _
            " INSERT INTO gtb_iva (iva_nome, iva_valore, iva_ordine) VALUES ('20%', 20, 1) ; " + _
            " INSERT INTO gtb_iva (iva_nome, iva_valore, iva_ordine) VALUES ('10%', 10, 2) ; " + _
            " INSERT INTO gtb_iva (iva_nome, iva_valore, iva_ordine) VALUES ('4%', 4, 3) ; " + _
            " INSERT INTO gtb_iva (iva_nome, iva_valore, iva_ordine) VALUES ('Esente', 0, 4) ; " + _
            " INSERT INTO gtb_stati_ordine (so_nome, so_ordine, so_stato_ordini, so_internet) VALUES ('ordine internet da verificare', 0, 0, 1); " + _
            " INSERT INTO gtb_stati_ordine (so_nome, so_ordine, so_stato_ordini, so_internet) VALUES ('ordine internet confermato', 0, 1, 0); " + _
            " INSERT INTO gtb_stati_ordine (so_nome, so_ordine, so_stato_ordini, so_internet) VALUES ('ordine evaso', 0, 2, 0); " + _
            " INSERT INTO gtb_stati_ordine (so_nome, so_ordine, so_stato_ordini, so_internet) VALUES ('ordine archiviato', 0, 3, 0); " + _
            " INSERT INTO gtb_valute (valu_nome, valu_codice, valu_cambio, valu_simbolo) VALUES ('Euro', 'EUR', 1, CHAR(128)) "
end function


function Indexing_B2B()
        Indexing_B2B = _
            " INSERT INTO tb_siti_tabelle (tab_sito_id, tab_titolo, tab_name, tab_field_chiave, tab_field_titolo, tab_field_descrizione, tab_field_foto_thumb, tab_field_foto_zoom, tab_field_visibile, tab_field_ordine, tab_parametro, tab_colore, tab_from_sql ) " + _
                " VALUES (" & NEXTB2B & ", 'Articoli - categorie', 'gtb_tipologie', 'tip_id', 'tip_nome_', 'tip_descr_', 'tip_foto', '', '(CASE WHEN tip_visibile AND tip_albero_visibile THEN 1 ELSE 0 END)', 'tip_ordine_assoluto', 'ID', '#FF6633', 'gtb_tipologie' ) ;" + _
            " INSERT INTO tb_siti_tabelle (tab_sito_id, tab_titolo, tab_name, tab_field_chiave, tab_field_titolo, tab_field_descrizione, tab_field_foto_thumb, tab_field_foto_zoom, tab_field_visibile, tab_field_ordine, tab_parametro, tab_colore, tab_from_sql ) " + _
                " VALUES (" & NEXTB2B & ", 'Articoli', 'gtb_articoli', 'art_id', 'art_nome_', 'art_descr_', " + _
                        " '(SELECT TOP 1 fo_thumb FROM gtb_art_foto WHERE fo_thumb<>'' AND fo_articolo_id=gtb_articoli.art_id ORDER BY fo_ordine)', '(SELECT TOP 1 fo_zoom FROM gtb_art_foto WHERE fo_thumb<>'' AND fo_articolo_id=gtb_articoli.art_id ORDER BY fo_ordine)', " + _
                        " '(CASE WHEN art_disabilitato THEN 0 ELSE 1 END)', '', 'ID', '#FF9933', 'gtb_articoli' ) "
end function


'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'AGGIORNAMENTI DI DATABASE

'**************************************************************************************************************************************************************************************
'**************************************************************************************************************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 1
'...........................................................................................
'aggiunge campi per la descrizione della composizione del bundle / confezione ed accessori
'...........................................................................................
function Aggiornamento__B2B__1(conn)
	Aggiornamento__B2B__1 = _
		" ALTER TABLE gtb_articoli ADD " + vbCrlf + _
		"		art_composizione_note_it ntext NULL, " + vbCrlf + _
		"		art_composizione_note_en ntext NULL, " + vbCrlf + _
		"		art_composizione_note_fr ntext NULL, " + vbCrlf + _
		"		art_composizione_note_es ntext NULL, " + vbCrlf + _
		"		art_composizione_note_de ntext NULL, " + vbCrlf + _
		"		art_accessori_note_it ntext NULL, " + vbCrlf + _
		"		art_accessori_note_en ntext NULL, " + vbCrlf + _
		"		art_accessori_note_fr ntext NULL, " + vbCrlf + _
		"		art_accessori_note_es ntext NULL, " + vbCrlf + _
		"		art_accessori_note_de ntext NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 2
'...........................................................................................
'aggiunta tabella per gestione "prodotti piu' ordinati"
'...........................................................................................
function Aggiornamento__B2B__2(conn)
	Aggiornamento__B2B__2 = _
		" CREATE TABLE dbo.gtb_articoli_ordinati ( " + vbCrlf + _
		"		ao_id int IDENTITY (1, 1) NOT NULL , " + vbCrlf + _
		"		ao_ut_id int NULL ," + vbCrlf + _
		"		ao_variante_id int NULL ," + vbCrlf + _
		"		ao_ranking int NULL " + vbCrlf + _
		"		); " + vbCrlf + _
		"	ALTER TABLE dbo.gtb_articoli_ordinati ADD CONSTRAINT PK_gtb_articoli_ordinati PRIMARY KEY  CLUSTERED " + vbCrlf + _
		"		( ao_id ) ;" + vbCrlf + _
		"	ALTER TABLE dbo.gtb_articoli_ordinati ADD CONSTRAINT FK_gtb_articoli_ordinati_tb_Utenti FOREIGN KEY " + vbCrlf + _
		"		( ao_ut_id ) " + vbCrlf + _
		"		REFERENCES tb_Utenti ( ut_ID) ON DELETE CASCADE  ON UPDATE CASCADE ; " + vbCrlf + _
		"	ALTER TABLE dbo.gtb_articoli_ordinati ADD CONSTRAINT FK_gtb_articoli_ordinati_grel_art_valori FOREIGN KEY " + vbCrlf + _
		"		( ao_variante_id ) REFERENCES grel_art_valori (	rel_id)  ON DELETE CASCADE  ON UPDATE CASCADE "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 3
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__3(conn)
	Aggiornamento__B2B__3 = _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_articoli_varianti ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_articoli_varianti AS " + vbCrlf + _
		" 	SELECT TOP 100 PERCENT dbo.gtb_valori.*, dbo.grel_art_vv.*, dbo.gtb_varianti.* " + vbCrlf + _
		" 		FROM dbo.grel_art_vv INNER JOIN dbo.gtb_valori ON dbo.grel_art_vv.rvv_val_id = dbo.gtb_valori.val_id INNER JOIN " + vbCrlf + _
		" 		dbo.gtb_varianti ON dbo.gtb_valori.val_var_id = dbo.gtb_varianti.var_id " + vbCrlf + _
		" 		ORDER BY dbo.gtb_varianti.var_ordine, dbo.gtb_varianti.var_id, dbo.gtb_valori.val_ordine ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CodificheArticoli ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_CodificheArticoli AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_codici ON dbo.grel_art_valori.rel_id = dbo.gtb_codici.Cod_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 4
'...........................................................................................
' crea la tabella per l'import dati ed il collegamento con relative relazioni
'...........................................................................................
function Aggiornamento__B2B__4(conn)
	Aggiornamento__B2B__4 = _
		" CREATE TABLE dbo.gItb_articoli ( " + _
		" 	Iart_id int IDENTITY (1, 1) NOT NULL ," + _
		" 	Iart_nome_IT nvarchar (250) NULL ," + _
		" 	Iart_descrizione_IT ntext NULL ," + _
		" 	IArt_stato_articolo nvarchar (500) NULL ," + _
		" 	Iart_Marca nvarchar (500) NULL ," + _
		" 	Iart_Tipologia nvarchar (1000) NULL ," + _
		" 	Iart_scorta_minima int NULL ," + _
		" 	Iart_lotto_minimo int NULL ," + _
		" 	Iart_lotto_riordino int NULL ," + _
		" 	Iart_Data_Ins_articolo smalldatetime NULL ," + _
		" 	Iart_Data_Upd_articolo smalldatetime NULL ," + _
		" 	IArt_Data_Ins_Import smalldatetime NULL ," + _
		" 	IArt_Data_Upd_Import smalldatetime NULL ," + _
		" 	Iart_Note ntext NULL ," + _
		" 	Iart_x_cod_int nvarchar (50) NULL ," + _
		" 	Iart_x_cod_alt nvarchar (50) NULL ," + _
		" 	Iart_x_cod_pro nvarchar (50) NULL ," + _
		" 	Iart_x_ID int NULL ," + _
		" 	Iart_b2b_ID int NULL ," + _
		" 	IArt_x_SourceCode nvarchar (10) NULL ," + _
		" 	IArt_User_UPD nvarchar (50) NULL ," + _
		" 	IArt_Prezzo money NULL ," + _
		" 	IArt_b2b_Update_price bit NULL ," + _
		" 	IArt_b2b_Update_state bit NULL ," + _
		" 	IArt_aliquota_iva nvarchar (100) NULL ," + _
		" 	IArt_b2b_field nvarchar (50) NULL ," + _
		" 	IArt_stato_catalogo_Articolo bit NULL ," + _
		" 	CONSTRAINT PK_gItb_articoli PRIMARY KEY CLUSTERED ( Iart_id ) " + _
		" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 5
'...........................................................................................
' aggiunge campi che indicano se l'articolo o la variante sono collegati ad un articolo esterno
'...........................................................................................
function Aggiornamento__B2B__5(conn)
	Aggiornamento__B2B__5 = _
		" ALTER TABLE gtb_articoli ADD " + vbCrLf + _
		"		art_external_id int NULL ;" + _
		" ALTER TABLE gtb_articoli ADD CONSTRAINT FK_gtb_articoli_gItb_articoli " + vbCrLF + _
		"		FOREIGN KEY (art_external_id) REFERENCES gItb_articoli (Iart_id) NOT FOR REPLICATION ; " + vbCrLf + _
		" ALTER TABLE gtb_articoli NOCHECK CONSTRAINT FK_gtb_articoli_gItb_articoli ; " + vbCrLF + _
		" ALTER TABLE grel_art_Valori ADD " + vbCrLf + _
		"		rel_external_id int NULL ;" + _
		" ALTER TABLE grel_art_valori ADD CONSTRAINT FK_grel_art_valori_gItb_articoli " + vbCrLF + _
		"		FOREIGN KEY (rel_external_id) REFERENCES gItb_articoli (Iart_id) NOT FOR REPLICATION ; " + vbCrLf + _
		" ALTER TABLE grel_art_valori NOCHECK CONSTRAINT FK_grel_art_valori_gItb_articoli"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 6
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__6(conn)
	Aggiornamento__B2B__6 = _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CodificheArticoli ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_CodificheArticoli AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_codici ON dbo.grel_art_valori.rel_id = dbo.gtb_codici.Cod_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 7
'...........................................................................................
' aggiunge i campi per la gestione dei prezzi delle varianti con sconti / variazioni in &euro;
' su prezzo base articolo
'...........................................................................................
function Aggiornamento__B2B__7(conn)
	Aggiornamento__B2B__7 = _
		" ALTER TABLE grel_Art_valori ADD " + _
		" 	rel_var_euro real NULL, " + _
		"		rel_var_sconto real  NULL, " + _
		"		rel_prezzo_indipendente bit NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 8
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__8(conn)
	Aggiornamento__B2B__8 = _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CodificheArticoli ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_CodificheArticoli AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_codici ON dbo.grel_art_valori.rel_id = dbo.gtb_codici.Cod_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 9
'...........................................................................................
' aggiunge il campo externalID per l'import delle tipologie
'...........................................................................................
function Aggiornamento__B2B__9(conn)
	Aggiornamento__B2B__9 = _
		" ALTER TABLE gtb_tipologie ADD " + _
		" 	tip_external_cod int"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 10
'...........................................................................................
' cambia il campo externalID per l'import delle tipologie
'...........................................................................................
function Aggiornamento__B2B__10(conn)
	Aggiornamento__B2B__10 = _
		" ALTER TABLE gtb_tipologie DROP COLUMN " + _
		" 	tip_external_cod;" + _
		" ALTER TABLE gtb_tipologie ADD " + _
		" 	tip_external_id int;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 11
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__11(conn)
	Aggiornamento__B2B__11 = _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CodificheArticoli ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_CodificheArticoli AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_codici ON dbo.grel_art_valori.rel_id = dbo.gtb_codici.Cod_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 12
'...........................................................................................
' aggiunge campo a tabella rivenditori per codice esterno cliente
'...........................................................................................
function Aggiornamento__B2B__12(conn)
	Aggiornamento__B2B__12 = _
		" ALTER TABLE gtb_rivenditori ADD riv_codice nvarchar(20) NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 13
'...........................................................................................
' aggiorna la vista sui rivenditori per variazione struttura tabella
'...........................................................................................
function Aggiornamento__B2B__13(conn)
	Aggiornamento__B2B__13 = _
		" DROP VIEW dbo.gv_rivenditori ; " + VbCrLf + _ 
		" CREATE VIEW dbo.gv_rivenditori AS " + VbCrLf + _ 
		"		SELECT * FROM gtb_rivenditori " + VbCrLf + _ 
		"		INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_utenti.ut_ID " + VbCrLf + _ 
		"		INNER JOIN tb_Indirizzario ON tb_utenti.ut_NextCom_ID = tb_indirizzario.IDElencoIndirizzi " + VbCrLf + _ 
		"		INNER JOIN gtb_valute ON gtb_rivenditori.riv_valuta_id = gtb_valute.valu_id "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 14
'...........................................................................................
' aggiunge campo a tabella rivenditori per codice esterno cliente
'...........................................................................................
function Aggiornamento__B2B__14(conn)
	Aggiornamento__B2B__14 = _
		" ALTER TABLE gtb_agenti ADD ag_codice nvarchar(20) NULL; " + _
		" DROP VIEW dbo.gv_agenti ; " + VbCrLf + _ 
		" CREATE VIEW dbo.gv_agenti AS " + VbCrLf + _
		" 	SELECT * FROM dbo.gtb_agenti " + VbCrLf + _
		"		INNER JOIN dbo.tb_admin ON dbo.gtb_agenti.ag_admin_id = dbo.tb_admin.ID_admin " + VbCrLf + _
		"		INNER JOIN dbo.tb_Utenti ON dbo.gtb_agenti.ag_id = dbo.tb_Utenti.ut_ID " + VbCrLf + _
		"		INNER JOIN dbo.tb_Indirizzario ON dbo.tb_Utenti.ut_NextCom_ID = dbo.tb_Indirizzario.IDElencoIndirizzi "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 15
'...........................................................................................
' aggiunge tabelle per gestione statistiche per mailing list
'...........................................................................................
function Aggiornamento__B2B__15(conn)
	Aggiornamento__B2B__15 = _
		"CREATE TABLE dbo.gtb_statistiche ( " + VbCrLf + _
		" 	sta_id int IDENTITY (1, 1) NOT NULL ,  " + VbCrLf + _
		" 	sta_nome nvarchar (50) NULL ,  " + VbCrLf + _
		" 	sta_dataC datetime NULL , " + VbCrLf + _
		" 	sta_temp bit NULL ,  " + VbCrLf + _
		" 	CONSTRAINT PK_gtb_statistiche PRIMARY KEY  CLUSTERED ( sta_id )  " + VbCrLf + _
		" ) ;  " + VbCrLf + _
		"CREATE TABLE dbo.gtb_criteri (  " + VbCrLf + _
		" 	cri_id int  IDENTITY (1, 1) NOT NULL , " + VbCrLf + _
		" 	cri_nome nvarchar (50) NULL ,  " + VbCrLf + _
		" 	cri_valore nvarchar (50) NULL , " + VbCrLf + _
		" 	cri_descrizione nvarchar (150) NULL , " + VbCrLf + _
		" 	cri_tipo nvarchar (50) NULL , " + VbCrLf + _
		" 	cri_statistica_id int NULL , " + VbCrLf + _
		" 	cri_operatore nvarchar (3) NULL , " + VbCrLf + _
		" 	CONSTRAINT PK_gtb_criteri PRIMARY KEY  CLUSTERED ( cri_id ) , " + VbCrLf + _
		" 	CONSTRAINT FK_gtb_criteri_gtb_statistiche FOREIGN KEY ( cri_statistica_id)  " + VbCrLf + _
		" 	REFERENCES gtb_statistiche ( sta_id	) ON DELETE CASCADE  ON UPDATE CASCADE  " + VbCrLf + _
		" ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 16
'...........................................................................................
' aggiunge campo su tabella marche per inserire il codice di collegamento esterno ed aggiorna
' le viste collegate
'...........................................................................................
function Aggiornamento__B2B__16(conn)
	Aggiornamento__B2B__16 = _
		" ALTER TABLE gtb_marche ADD mar_codice nvarchar(20) NULL; " + _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 17
'...........................................................................................
' aggiunge campi tabella di import articoli per gestione marche e tipologie
'...........................................................................................
function Aggiornamento__B2B__17(conn)
	Aggiornamento__B2B__17 = _
		" ALTER TABLE gItb_articoli ADD " + _
		"				Iart_Marca_codice nvarchar(50) NULL, " + _
		"				Iart_tipologia_id int NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 18
'...........................................................................................
' aggiunge campi a tabella di import articoli per collegamento con codici fornitori.
'...........................................................................................
function Aggiornamento__B2B__18(conn)
	Aggiornamento__B2B__18 = _
		" ALTER TABLE gItb_articoli ADD " + _
		"				Iart_x_cod_fornitore nvarchar(50) NULL, " + _
		"				Iart_x_fornitore nvarchar(50) NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 19
'...........................................................................................
' aggiunge campi a tabella di import articoli per collegamento con codici fornitori.
'...........................................................................................
function Aggiornamento__B2B__19(conn)
	Aggiornamento__B2B__19 = _
		" ALTER TABLE gItb_articoli ADD " + _
		"				Iart_Varianti bit NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 20
'...........................................................................................
' aggiunge campo foto a variante
'...........................................................................................
function Aggiornamento__B2B__20(conn)
	Aggiornamento__B2B__20 = _
		" ALTER TABLE grel_art_valori ADD "+ vbCrlf + _
		"		rel_foto_id INTEGER NULL; "+ vbCrlf + _
		"	ALTER TABLE grel_art_valori WITH NOCHECK ADD CONSTRAINT FK_grel_art_valori__gtb_art_foto FOREIGN KEY "+ vbCrlf + _
		"		( rel_foto_id ) REFERENCES gtb_art_foto ( fo_id ) NOT FOR REPLICATION; "+ vbCrlf + _
		" ALTER TABLE grel_art_valori NOCHECK CONSTRAINT FK_grel_art_valori__gtb_art_foto"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 21
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__21(conn)
	Aggiornamento__B2B__21 = _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CodificheArticoli ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_CodificheArticoli AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_codici ON dbo.grel_art_valori.rel_id = dbo.gtb_codici.Cod_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 22
'...........................................................................................
' rivoluzione codici: aggiunta flag listino di sistema, inserisco record solo per codici che cambiano,
'					  cancello vista gv_codificheArticoli
'...........................................................................................
function Aggiornamento__B2B__22(conn)
	Aggiornamento__B2B__22 = _
		" ALTER TABLE gtb_lista_codici ADD "+ vbCrlf + _
		"		lstCod_sistema BIT NULL; "+ vbCrlf + _
		" UPDATE gtb_lista_codici SET lstCod_sistema = 1; "+ vbCrlf + _
		" DELETE FROM gtb_codici WHERE cod_codice = (SELECT rel_cod_int FROM grel_art_valori WHERE rel_id=cod_variante_id); "+ vbCrlf + _
		" DROP VIEW dbo.gv_CodificheArticoli ; " + VbCrLf + _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 23
'...........................................................................................
' aggiunge tabella per la registrazione dei codici fornitori esterne.
'...........................................................................................
function Aggiornamento__B2B__23(conn)
	Aggiornamento__B2B__23 = _
		" CREATE TABLE dbo.gItb_articoli_cod_fornitori ( " + VbCrLf + _
		"		cod_id int IDENTITY (1, 1) NOT NULL , " + VbCrLf + _
		"		cod_Iart_id int NULL , " + VbCrLf + _
		"		cod_codice_articolo nvarchar (50) NULL , " + VbCrLf + _
		"		cod_codice_fornitore nvarchar (50) NULL , " + VbCrLf + _
		"		cod_fornitore_preferenziale bit NULL , " + VbCrLf + _
		"		CONSTRAINT PK_gItb_articoli_cod_fornitori PRIMARY KEY  CLUSTERED  " + VbCrLf + _
		"		(cod_id), " + VbCrLf + _
		"		CONSTRAINT FK_gItb_articoli_cod_fornitori_gItb_articoli " + VbCrLf + _
		"		FOREIGN KEY (cod_Iart_id) REFERENCES gItb_articoli (Iart_id)  " + VbCrLf + _
		"		ON DELETE CASCADE  ON UPDATE CASCADE )"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 24
'...........................................................................................
' aggiunge tabella per la registrazione dei codici fornitori esterne.
'...........................................................................................
function Aggiornamento__B2B__24(conn)
	Aggiornamento__B2B__24 = _
		" ALTER TABLE gtb_magazzini ADD mag_codice nvarchar(50) NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 25
'...........................................................................................
' aggiunge campi per indicazione marca di default
'...........................................................................................
function Aggiornamento__B2B__25(conn)
	Aggiornamento__B2B__25 = _
		" ALTER TABLE gtb_marche ADD " + _
		" 	mar_generica BIT NULL; " + _
		" UPDATE gtb_marche SET mar_generica=0 "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 26
'...........................................................................................
'aggiunge tabella per gestione gerarchia categorie foglie
'...........................................................................................
function Aggiornamento__B2B__26(conn)
	Aggiornamento__B2B__26 = _
		" ALTER TABLE gtb_tipologie ADD " + _
		"		tip_tipologia_padre_base INT NULL; " + _
		"	ALTER TABLE gtb_tipologie ADD CONSTRAINT FK_gtb_tipologie_gtb_tipologie_padre_base " + _
		"		FOREIGN KEY ( tip_tipologia_padre_base ) " + vbCrlf + _
		"		REFERENCES gtb_tipologie ( tip_ID) NOT FOR REPLICATION; " + _
		" ALTER TABLE gtb_tipologie NOCHECK CONSTRAINT FK_gtb_tipologie_gtb_tipologie_padre_base; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 27
'...........................................................................................
' aggiorna dati delle tipologie e relativi padri
'...........................................................................................
function Aggiornamento__B2B__27(conn, rs)
	dim sql, level
	sql = "SELECT tip_livello FROM gtb_tipologie GROUP BY tip_livello"
	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	sql = ""
	while not rs.eof
		sql = sql & "SELECT (Tip_L" & rs("tip_livello") & ".tip_id) AS tip_padre_id, " +_
							"(Tip_L" & rs("tip_livello") & ".tip_nome_it) AS tip_padre_nome, " +_
							"TIP_L0.tip_id, TIP_L0.tip_codice, (" 
		for level = rs("tip_livello") to 1 step -1 
			sql = sql & "TIP_L" & level & ".tip_nome_it " & SQL_concat(conn) & " ' - ' " & SQL_concat(conn)
		next
		sql = sql & " TIP_L0.tip_nome_it) AS NAME FROM " & IIF(cInteger(rs("tip_livello"))>0, String(rs("tip_livello"), "("), "") & " gtb_tipologie TIP_L0 " 
		for level = 1 to rs("tip_livello")
			sql = sql & " INNER JOIN gtb_tipologie TIP_L" & level & " ON TIP_L" & (level-1) & ".tip_padre_id = TIP_L" & level & ".tip_id ) "
		next
		sql = sql & "WHERE TIP_L0.tip_livello=" & rs("tip_livello")
		rs.movenext
		if not rs.eof then
			sql = sql & " UNION "
		end if
	wend
	rs.close
	if sql <> "" then
		sql = sql & " ORDER BY NAME"
	else
		sql = "SELECT * FROM gtb_Tipologie ORDER BY tip_nome_it"
	end if
	
	rs.open sql, DB.objConn, adOpenStatic, adLockOptimistic, adCmdText
	sql = ""
	while not rs.eof
		sql = sql + " UPDATE gtb_tipologie SET tip_tipologia_padre_base=" & rs("tip_padre_id") & _
			  		" WHERE tip_id=" & rs("tip_id") & ";" + vbCrLf
		rs.movenext
	wend
	rs.close
	Aggiornamento__B2B__27 = sql
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 28
'...........................................................................................
' aggiunge gestione raggruppamenti
'...........................................................................................
function Aggiornamento__B2B__28(conn)
	Aggiornamento__B2B__28 = _
		"CREATE TABLE dbo.gtb_tipologie_raggruppamenti ( " + _
		" rag_id int IDENTITY (1, 1) NOT NULL , " + _
		" rag_nome_it nvarchar (250) NULL , " + _
		" rag_nome_en nvarchar (250) NULL , " + _
		" rag_nome_fr nvarchar (250) NULL , " + _
		" rag_nome_es nvarchar (250) NULL , " + _
		" rag_nome_de nvarchar (250) NULL , " + _
		" rag_foto nvarchar (255) NULL , " + _
		" rag_descr_it ntext NULL , " + _
		" rag_descr_en ntext NULL , " + _
		" rag_descr_fr ntext NULL , " + _
		" rag_descr_es ntext NULL , " + _
		" rag_descr_de ntext NULL , " + _
		" rag_ordine int NOT NULL , " + _
		" rag_tipologia_id int NULL , " + _
		" CONSTRAINT PK_gtb_tipologie_raggruppamenti " + _
		" 	PRIMARY KEY  CLUSTERED ( rag_id ), " + _
		" CONSTRAINT FK_gtb_tipologie_raggruppamenti_gtb_tipologie " + _
		" 	FOREIGN KEY ( rag_tipologia_id ) REFERENCES gtb_tipologie ( tip_id ) ON DELETE CASCADE  ON UPDATE CASCADE )"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 29
'...........................................................................................
' aggiunge relazione raggruppamenti articoli
'...........................................................................................
function Aggiornamento__B2B__29(conn)
	Aggiornamento__B2B__29 = _
		" ALTER TABLE gtb_articoli ADD art_raggruppamento_id INT null; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 30
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__30(conn)
	Aggiornamento__B2B__30 = _
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 31
'...........................................................................................
'corregge relazione tra articoli e raggruppamenti
'...........................................................................................
function Aggiornamento__B2B__31(conn)
	Aggiornamento__B2B__31 = _
		" ALTER TABLE gtb_articoli ADD CONSTRAINT FK_gtb_articoli_gtb_tipologie_raggruppamenti" + _
		" FOREIGN KEY (art_raggruppamento_id) REFERENCES gtb_tipologie_raggruppamenti (rag_id)" + _ 
		" NOT FOR REPLICATION; " + _
		" ALTER TABLE dbo.gtb_articoli NOCHECK CONSTRAINT FK_gtb_articoli_gtb_tipologie_raggruppamenti"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 32
'...........................................................................................
'aggiunge campi per descrizione varianti e valori di variante
'...........................................................................................
function Aggiornamento__B2B__32(conn)
	Aggiornamento__B2B__32 = _
		" ALTER TABLE gtb_varianti ADD " + _
		" 	var_descr_it ntext NULL , " + _
		" 	var_descr_en ntext NULL , " + _
		" 	var_descr_fr ntext NULL , " + _
		" 	var_descr_es ntext NULL , " + _
		" 	var_descr_de ntext NULL ;" + _
		" ALTER TABLE gtb_valori ADD " + _
		" 	val_descr_it ntext NULL , " + _
		" 	val_descr_en ntext NULL , " + _
		" 	val_descr_fr ntext NULL , " + _
		" 	val_descr_es ntext NULL , " + _
		" 	val_descr_de ntext NULL ;" + _
		" DROP VIEW gv_articoli_varianti; " + _
		" CREATE VIEW dbo.gv_articoli_varianti AS " + VbCrLf + _
		" 	SELECT TOP 100 PERCENT dbo.gtb_valori.*, dbo.grel_art_vv.*, dbo.gtb_varianti.* " + VbCrLf + _
		" 	FROM dbo.grel_art_vv INNER JOIN dbo.gtb_valori ON dbo.grel_art_vv.rvv_val_id = dbo.gtb_valori.val_id " + VbCrLf + _
		"		INNER JOIN dbo.gtb_varianti ON dbo.gtb_valori.val_var_id = dbo.gtb_varianti.var_id " + VbCrLf + _
		"		ORDER BY dbo.gtb_varianti.var_ordine, dbo.gtb_varianti.var_id, dbo.gtb_valori.val_ordine; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 33
'...........................................................................................
'aggiunge campi per descrizione delle varianti
'...........................................................................................
function Aggiornamento__B2B__33(conn)
	Aggiornamento__B2B__33 = _
		" ALTER TABLE grel_art_valori ADD " + _
		" 	rel_descr_it ntext NULL , " + _
		" 	rel_descr_en ntext NULL , " + _
		" 	rel_descr_fr ntext NULL , " + _
		" 	rel_descr_es ntext NULL , " + _
		" 	rel_descr_de ntext NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 34
'...........................................................................................
'aggiornamento viste per variazione struttura tabelle articoli.
'...........................................................................................
function Aggiornamento__B2B__34(conn)
	Aggiornamento__B2B__34 = _ 
		" DROP VIEW dbo.gv_articoli ;" + _
		" CREATE VIEW dbo.gv_articoli AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_carichi ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_carichi AS " + vbCrlf + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrlf + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrlf + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_CartDetail ; " + vbCrlf + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrlf + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrlf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrlf + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrlf + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrlf + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id ; " + vbCrlf + _
		" DROP VIEW dbo.gv_dettagli_ord ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + VbCrLf + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_Giacenze_Varianti ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_Giacenze_Varianti AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id ; " + VbCrLf + _
		"	DROP VIEW dbo.gv_inventario ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_inventario AS " + VbCrLf + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + VbCrLf + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listini ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listini AS " + VbCrLf + _
		" 	SELECT * FROM dbo.grel_art_valori INNER JOIN " + VbCrLf + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + VbCrLf + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_offerte ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE gtb_articoli.art_disabilitato = 0 AND " + VbCrLf + _
		"				grel_art_valori.rel_disabilitato=0 AND " + VbCrLf + _
		"				listino_offerte=1 AND " + VbCrLf + _
		"				prz_visibile=1 AND " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1) ; " + VbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita ; " + VbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS  " + VbCrLf + _
		"		SELECT * FROM gtb_articoli " + VbCrLf + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + VbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + VbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + VbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + VbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND  " + VbCrLf + _
		"				ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + VbCrLf + _
		"				((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR  " + VbCrLf + _
		"				(ISNULL(listino_offerte, 0)=0 AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id  " + VbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND  " + VbCrLf + _
		"				(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 35
'...........................................................................................
'aggiunge tabella per tipo relazione "accessori" (prodotti correlati ed altro)
'...........................................................................................
function Aggiornamento__B2B__35(conn)
	Aggiornamento__B2B__35 = _
		" CREATE TABLE dbo.gtb_accessori_tipo ( " + _
		" 	at_id int IDENTITY (1, 1) NOT NULL , " + _
		" 	at_nome_it nvarchar (250) NULL , " + _
		" 	at_nome_en nvarchar (250) NULL , " + _
		" 	at_nome_fr nvarchar (250) NULL , " + _
		" 	at_nome_es nvarchar (250) NULL , " + _
		" 	at_nome_de nvarchar (250) NULL , " + _
		" 	at_ordine int NOT NULL, " + _
		" 	CONSTRAINT PK_gtb_accessori_tipo " + _
		"		 	PRIMARY KEY  CLUSTERED ( at_id )" + _
		" );" + _
		" ALTER TABLE grel_art_acc ADD " + _
		"		aa_tipo_id INT NULL; " + _
		" UPDATE grel_art_acc SET aa_tipo_id=1; " + _
		" ALTER TABLE grel_art_acc ALTER COLUMN aa_tipo_id INT NOT NULL; " + _
		" INSERT INTO gtb_accessori_tipo (at_nome_it, at_ordine) VALUES ('Accessori', 2); " + _
		" INSERT INTO gtb_accessori_tipo (at_nome_it, at_ordine) VALUES ('Prodotti correlati', 1); " + _
		" ALTER TABLE grel_art_acc ADD CONSTRAINT FK_grel_art_acc_gtb_accessori_tipo " + _
		" 	FOREIGN KEY (aa_tipo_id) REFERENCES gtb_accessori_tipo(at_id) " + _
		"		ON DELETE CASCADE ON UPDATE CASCADE "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 36
'...........................................................................................
'aggiunge campi a relazione accessori / prodotti correlati
'...........................................................................................
function Aggiornamento__B2B__36(conn)
	Aggiornamento__B2B__36 = _
		" ALTER TABLE grel_art_acc ADD " + _
		"		aa_ordine INT NULL, " + _
		" 	aa_note_it nvarchar (250) NULL , " + _
		" 	aa_note_en nvarchar (250) NULL , " + _
		" 	aa_note_fr nvarchar (250) NULL , " + _
		" 	aa_note_es nvarchar (250) NULL , " + _
		" 	aa_note_de nvarchar (250) NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 37
'...........................................................................................
'corregge valori a 0 su raggruppamento articolo
'cancella agenti non pi&ugrave; validi
'corregge valori a 0 su id foto della variante.
'corregge valori classi di sconto
'...........................................................................................
function Aggiornamento__B2B__37(conn)
	Aggiornamento__B2B__37 = _
		" UPDATE gtb_articoli SET art_raggruppamento_id=NULL WHERE IsNull(art_raggruppamento_id, 0)=0; " + _
		" DELETE FROM gtb_agenti WHERE ag_admin_id NOT IN (SELECT ID_admin FROM tb_admin); " + _
		" UPDATE grel_art_valori SET rel_foto_id=NULL WHERE IsNull(rel_foto_id, 0)=0; " + _
		" UPDATE grel_art_valori SET rel_scontoQ_id=NULL " + _
		"			WHERE IsNULL(rel_scontoQ_id,0)=0 OR " + _
		"				  (IsNULL(rel_scontoQ_id,0)<>0 AND rel_scontoQ_id NOT IN (SELECT scc_id FROM gtb_scontiQ_classi) );" + _
		" UPDATE gtb_Articoli SET art_scontoQ_id=NULL " + _
		"			WHERE IsNULL(art_scontoQ_id,0)=0 OR " + _
		"				  (IsNULL(art_scontoQ_id,0)<>0 AND art_scontoQ_id NOT IN (SELECT scc_id FROM gtb_scontiQ_classi) );" + _
		" UPDATE gtb_prezzi SET prz_scontoQ_id=NULL " + _
		"			WHERE IsNULL(prz_scontoQ_id,0)=0 OR " + _
		"				  (IsNULL(prz_scontoQ_id,0)<>0 AND prz_scontoQ_id NOT IN (SELECT scc_id FROM gtb_scontiQ_classi) );" + _
		" UPDATE gtb_rivenditori SET riv_agente_id=NULL " + _
		"			WHERE IsNUll(riv_agente_id, 0)=0 OR " + _
		"				  (IsNull(riv_agente_id, 0)<>0 AND riv_agente_id NOT IN (SELECT ag_id FROM gtb_agenti) );"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 38
'...........................................................................................
'aggiunge relazinone non vincolante mancante
'...........................................................................................
function Aggiornamento__B2B__38(conn)
	Aggiornamento__B2B__38 = _
		" ALTER TABLE gtb_articoli WITH NOCHECK ADD CONSTRAINT FK_gtb_articoli_gtb_scontiQ_classi " + _
		"		FOREIGN KEY (art_scontoQ_id) REFERENCES gtb_scontiQ_classi(scc_id)" + _
		"		NOT FOR REPLICATION; " + _
		" ALTER TABLE gtb_articoli NOCHECK CONSTRAINT FK_gtb_articoli_gtb_scontiQ_classi "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 39
'...........................................................................................
'aggiorna stato ordine relazione tra caratteristiche tecniche e categorie per riconoscere 
'che &egrave; presente
'...........................................................................................
function Aggiornamento__B2B__39(conn)
	Aggiornamento__B2B__39 = _
		" UPDATE gtb_tip_ctech SET rct_ordine=0;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 40
'...........................................................................................
'aggiunge campo alle categorie per indicazione visibilita'
'...........................................................................................
function Aggiornamento__B2B__40(conn)
	Aggiornamento__B2B__40 = _
		" ALTER TABLE gtb_tipologie ADD " + _
		"	tip_visibile BIT NULL; " + _
		" UPDATE gtb_tipologie SET tip_visibile=1; " + _
		" ALTER TABLE gtb_tipologie ALTER COLUMN tip_visibile BIT NOT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 41
'...........................................................................................
'aggiunge campo alla struttura degli articoli correlati per vincolo in fase di vendita
'...........................................................................................
function Aggiornamento__B2B__41(conn)
	Aggiornamento__B2B__41 = _
		" ALTER TABLE gtb_accessori_tipo ADD " + _
		"	at_vincolo_vendita BIT NULL; " + _
		" UPDATE gtb_accessori_tipo SET at_vincolo_vendita=0; " + _
		" ALTER TABLE gtb_accessori_tipo ALTER COLUMN at_vincolo_vendita BIT NOT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 42
'...........................................................................................
'aggiornamento delle viste per variazione struttura campi
'...........................................................................................
function Aggiornamento__B2B__42(conn)
	Aggiornamento__B2B__42 = _
		"DROP VIEW gv_agenti; " + _
		"DROP VIEW gv_articoli; " + _
		"DROP VIEW gv_articoli_varianti; " + _
		"DROP VIEW gv_carichi; " + _
		"DROP VIEW gv_CartDetail; " + _
		"DROP VIEW gv_dettagli_ord; " + _
		"DROP VIEW gv_Giacenze_Varianti; " + _
		"DROP VIEW gv_inventario; " + _
		"DROP VIEW gv_listini; " + _
		"DROP VIEW gv_listino_offerte; " + _
		"DROP VIEW gv_listino_vendita; " + _
		"DROP VIEW gv_rivenditori; " + _
		"CREATE VIEW dbo.gv_agenti AS " + vbCrLF + _
		"		SELECT * FROM dbo.gtb_agenti " + vbCrLF + _
		"			INNER JOIN dbo.tb_admin ON dbo.gtb_agenti.ag_admin_id = dbo.tb_admin.ID_admin " + vbCrLF + _
		"			INNER JOIN dbo.tb_Utenti ON dbo.gtb_agenti.ag_id = dbo.tb_Utenti.ut_ID " + vbCrLF + _
		"			INNER JOIN dbo.tb_Indirizzario ON dbo.tb_Utenti.ut_NextCom_ID = dbo.tb_Indirizzario.IDElencoIndirizzi; " + vbCrLF + _
		"CREATE VIEW dbo.gv_articoli AS " + vbCrLF + _
		"		SELECT * FROM dbo.gtb_articoli INNER JOIN " + vbCrLF + _
		"			dbo.grel_art_valori ON dbo.gtb_articoli.art_id = dbo.grel_art_valori.rel_art_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_iva ON dbo.gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_articoli_varianti AS " + vbCrLF + _
		"		SELECT TOP 100 PERCENT dbo.gtb_valori.*, dbo.grel_art_vv.*, dbo.gtb_varianti.* " + vbCrLF + _
		"			FROM dbo.grel_art_vv INNER JOIN dbo.gtb_valori ON dbo.grel_art_vv.rvv_val_id = dbo.gtb_valori.val_id " + vbCrLF + _
		"			INNER JOIN dbo.gtb_varianti ON dbo.gtb_valori.val_var_id = dbo.gtb_varianti.var_id " + vbCrLF + _
		"			ORDER BY dbo.gtb_varianti.var_ordine, dbo.gtb_varianti.var_id, dbo.gtb_valori.val_ordine; " + vbCrLF + _
		"CREATE VIEW dbo.gv_carichi AS " + vbCrLF + _
		"		SELECT * FROM dbo.grel_carichi_var INNER JOIN " + vbCrLF + _
		"			dbo.grel_art_valori ON dbo.grel_carichi_var.rcv_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_marche ON dbo.gtb_articoli.art_marca_id = dbo.gtb_marche.mar_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_CartDetail AS " + vbCrLF + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST, " + vbCrLF + _
		"		          (SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrLF + _
		"		FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id " + vbCrLF + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLF + _
		"		INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_dettagli_ord AS " + vbCrLF + _
		"		SELECT * FROM dbo.gtb_dettagli_ord INNER JOIN " + vbCrLF + _
		"			dbo.grel_art_valori ON dbo.gtb_dettagli_ord.det_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_Giacenze_Varianti AS " + vbCrLF + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + vbCrLF + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrLF + _
		"			dbo.grel_carichi_var ON dbo.grel_art_valori.rel_id = dbo.grel_carichi_var.rcv_art_var_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_inventario AS " + vbCrLF + _
		"		SELECT * FROM dbo.grel_giacenze INNER JOIN " + vbCrLF + _
		"			dbo.grel_art_valori ON dbo.grel_giacenze.gia_art_var_id = dbo.grel_art_valori.rel_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_listini AS " + vbCrLF + _
		"		SELECT * FROM dbo.grel_art_valori INNER JOIN " + vbCrLF + _
		"			dbo.gtb_prezzi ON dbo.grel_art_valori.rel_id = dbo.gtb_prezzi.prz_variante_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_articoli ON dbo.grel_art_valori.rel_art_id = dbo.gtb_articoli.art_id INNER JOIN " + vbCrLF + _
		"			dbo.gtb_iva ON dbo.gtb_prezzi.prz_iva_id = dbo.gtb_iva.iva_id; " + vbCrLF + _
		"CREATE VIEW dbo.gv_listino_offerte AS " + vbCrLF + _
		"		SELECT * FROM gtb_articoli " + vbCrLF + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLF + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLF + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"		WHERE gtb_articoli.art_disabilitato = 0 AND " + vbCrLF + _
		"			grel_art_valori.rel_disabilitato=0 AND listino_offerte=1 AND " + vbCrLF + _
		"			prz_visibile=1 AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1); " + vbCrLF + _
		"CREATE VIEW dbo.gv_listino_vendita AS " + vbCrLF + _
		"		SELECT * FROM gtb_articoli " + vbCrLF + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLF + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLF + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"		WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 AND " + vbCrLF + _
		"			ISNULL(grel_art_valori.rel_disabilitato, 0)=0 AND " + vbCrLF + _
		"			((ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND " + vbCrLF + _
		"			(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1)) OR " + vbCrLF + _
		"			(ISNULL(listino_offerte, 0)=0  " + vbCrLF + _
		"			AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id " + vbCrLF + _
		"			                            WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 AND " + vbCrLF + _
		"			(GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))); " + vbCrLF + _
		"CREATE VIEW dbo.gv_rivenditori AS " + vbCrLF + _
		"		SELECT * FROM gtb_rivenditori " + vbCrLF + _
		"		INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_utenti.ut_ID " + vbCrLF + _
		"		INNER JOIN tb_Indirizzario ON tb_utenti.ut_NextCom_ID = tb_indirizzario.IDElencoIndirizzi " + vbCrLF + _
		"		INNER JOIN gtb_valute ON gtb_rivenditori.riv_valuta_id = gtb_valute.valu_id; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 43
'...........................................................................................
'aggiunge campo alle categorie per indicazione visibilita' rispetto alle categorie padre
'...........................................................................................
function Aggiornamento__B2B__43(conn)
	Aggiornamento__B2B__43 = _
		" ALTER TABLE gtb_tipologie ADD " + _
		"	tip_albero_visibile BIT NULL; " + _
		" UPDATE gtb_tipologie SET tip_albero_visibile=1, tip_visibile=1; " + _
		" ALTER TABLE gtb_tipologie ALTER COLUMN tip_albero_visibile BIT NOT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 44
'...........................................................................................
'aggiornamento delle viste per variazione struttura campi ed aggiunta condiizoni 
'per visibilita' tipologie e relative join
'...........................................................................................
function Aggiornamento__B2B__44(conn)
	Aggiornamento__B2B__44 = _
		"DROP VIEW gv_articoli; " + _
		"CREATE VIEW dbo.gv_articoli AS" + vbCrLF + _
		"		SELECT * FROM gtb_articoli" + vbCrLF + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLF + _
		"			INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = dbo.gtb_iva.iva_id " + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id; " + _
		"DROP VIEW gv_carichi; " + _
		"CREATE VIEW dbo.gv_carichi AS " + vbCrLF + _
		"		SELECT * FROM grel_carichi_var " + vbCrLF + _
		"			INNER JOIN grel_art_valori ON grel_carichi_var.rcv_art_var_id = grel_art_valori.rel_id" + vbCrLF + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLF + _
		"			INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id;" + _
		"DROP VIEW gv_CartDetail; " + _
		"CREATE VIEW dbo.gv_CartDetail AS" + vbCrLF + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST," + vbCrLF + _
		"		          (SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT" + vbCrLF + _
		"		FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id" + vbCrLF + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id; " + _
		"DROP VIEW gv_inventario; " + _
		"CREATE VIEW dbo.gv_inventario AS" + vbCrLF + _
		"		SELECT * FROM grel_giacenze" + vbCrLF + _
		"			INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id" + vbCrLF + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id; " + _
		"DROP VIEW gv_listini; " + _
		"CREATE VIEW dbo.gv_listini AS" + vbCrLF + _
		"		SELECT * FROM grel_art_valori" + vbCrLF + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLF + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id;" + _
		"DROP VIEW gv_listino_offerte; " + _
		"CREATE VIEW dbo.gv_listino_offerte AS" + vbCrLF + _
		"		SELECT * FROM gtb_articoli" + vbCrLF + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLF + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLF + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0" + vbCrLf + _
		"			AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0" + vbCrLf + _
		"			AND tip_visibile=1" + vbCrLf + _
		"			AND tip_albero_visibile=1" + vbCrLf + _
		"			AND ISNULL(listino_offerte, 0)=1" + vbCrLf + _
		"			AND ISNULL(prz_visibile, 0)=1" + vbCrLf + _
		"			AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1); " + _
		"DROP VIEW gv_listino_vendita; " + _
		"CREATE VIEW dbo.gv_listino_vendita AS" + vbCrLF + _
		"		SELECT * FROM gtb_articoli" + vbCrLF + _
		"			INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLF + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLF + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"			INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLF + _
		"		WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0" + vbCrLf + _
		"			AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLf + _
		"			AND tip_visibile=1" + vbCrLf + _
		"			AND tip_albero_visibile=1" + vbCrLf + _
		"			AND ( ( ISNULL(listino_offerte, 0)=1" + vbCrLf + _
		"					AND ISNULL(prz_visibile, 0)=1" + vbCrLf + _
		"					AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) )" + vbCrLf + _
		"					OR (ISNULL(listino_offerte, 0)=0" + vbCrLf + _
		"					AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id" + vbCrLf + _
		"					WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1" + vbCrLf + _
		"					AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 45
'...........................................................................................
'aggiunge campo di visibilita' ai gruppi
'...........................................................................................
function Aggiornamento__B2B__45(conn)
	Aggiornamento__B2B__45 = _
		" ALTER TABLE gtb_tipologie_raggruppamenti ADD " + _
		"	rag_visibile BIT NULL; " + _
		" UPDATE gtb_tipologie_raggruppamenti SET rag_visibile=1; " + _
		" ALTER TABLE gtb_tipologie_raggruppamenti ALTER COLUMN rag_visibile BIT NOT NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 46
'...........................................................................................
'aggiunge campo di visibilita' ai gruppi
'...........................................................................................
function Aggiornamento__B2B__46(conn)
	Aggiornamento__B2B__46 = _
		" ALTER TABLE gItb_articoli ALTER COLUMN Iart_tipologia_id nvarchar(50);"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 47
'...........................................................................................
'modifica dati codice esterno tipologie
'...........................................................................................
function Aggiornamento__B2B__47(conn)
	Aggiornamento__B2B__47 = _
		" ALTER TABLE gtb_tipologie ALTER COLUMN tip_external_id nvarchar(50);"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 48
'...........................................................................................
'creazione nuova vista per gestione disponibilit&agrave; articoli
'...........................................................................................
function Aggiornamento__B2B__48(conn)
	Aggiornamento__B2B__48 = _
		"CREATE VIEW dbo.gv_giacenza_pubblico AS" + vbCrLf + _
		"	SELECT SUM(gia_qta - gia_impegnato) AS giacenza, rel_id, MIN(rel_art_id) AS articolo_id FROM grel_giacenze" + vbCrLf + _
		"		INNER JOIN gtb_magazzini ON grel_giacenze.gia_magazzino_id = gtb_magazzini.mag_id" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		WHERE (gtb_magazzini.mag_vendita_pubblico=1)" + vbCrLf + _
		"		OR ((SELECT COUNT(*) FROM gtb_magazzini WHERE mag_vendita_pubblico=1)=0)" + vbCrLf + _
		"		GROUP BY rel_id ;"
end function
'******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 49
'...........................................................................................
'modifica vista per gestione giacenza
'...........................................................................................
function Aggiornamento__B2B__49(conn)
	Aggiornamento__B2B__49 = _
		"DROP VIEW gv_giacenza_pubblico; " + _
		"CREATE VIEW dbo.gv_giacenza_pubblico AS" + vbCrLf + _
		"	SELECT SUM(gia_qta - gia_impegnato) AS giacenza, rel_id, MIN(rel_art_id) AS articolo_id FROM grel_giacenze" + vbCrLf + _
		"		INNER JOIN gtb_magazzini ON grel_giacenze.gia_magazzino_id = gtb_magazzini.mag_id" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		WHERE (gtb_magazzini.mag_disponibilita=1)" + vbCrLf + _
		"		OR ((SELECT COUNT(*) FROM gtb_magazzini WHERE mag_disponibilita=1)=0)" + vbCrLf + _
		"		GROUP BY rel_id ;"
end function
'******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 50
'...........................................................................................
'aggiorna gestione shopping cart per aggiunta tipologia "preventivi"
'...........................................................................................
function Aggiornamento__B2B__50(conn)
	Aggiornamento__B2B__50 = _
		" ALTER TABLE gtb_shopping_cart ADD sc_giorni_validita_preventivo INT NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 51
'...........................................................................................
'aggiorna struttura viste che gestiscono i dettagli degli ordini.
'...........................................................................................
function Aggiornamento__B2B__51(conn)
	Aggiornamento__B2B__51 = _
		"	DROP VIEW gv_CartDetail; " + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrLF + _
		"		SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST, " + vbCrLF + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT" + vbCrLF + _
		"			FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id" + vbCrLF + _
		"			INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLF + _
		"			INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id" + vbCrLF + _
		"			INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 52
'...........................................................................................
'aggiorna gestione shopping cart per aggiunta tipologia "preventivi"
'...........................................................................................
function Aggiornamento__B2B__52(conn)
	Aggiornamento__B2B__52 = _
		" UPDATE gtb_shopping_cart SET sc_Date_cart=NULL; " + _
		" ALTER TABLE gtb_shopping_cart ALTER COLUMN sc_Date_cart SMALLDATETIME NULL; " + _
		" UPDATE gtb_shopping_cart SET sc_date_cart = GETDATE(); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 53
'...........................................................................................
'corregge problema relazione per trasferimento dati via DTS
'...........................................................................................
function Aggiornamento__B2B__53(conn)
	Aggiornamento__B2B__53 = _
		" UPDATE gtb_articoli SET art_raggruppamento_id=NULL WHERE IsNull(art_raggruppamento_id, 0)=0; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 54
'...........................................................................................
'corregge problema relazione per trasferimento dati via DTS
'...........................................................................................
function Aggiornamento__B2B__54(conn)
	Aggiornamento__B2B__54 = _
		" UPDATE gtb_prezzi SET prz_scontoQ_id=NULL WHERE IsNull(prz_scontoQ_id, 0)=0; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 55
'...........................................................................................
'corregge problema relazione per trasferimento dati via DTS
'...........................................................................................
function Aggiornamento__B2B__55(conn)
	Aggiornamento__B2B__55 = _
		" ALTER TABLE gItb_articoli ADD "  + _
		" 	IArt_prezzo_var_euro REAL NULL, " + _
		"		IArt_prezzo_var_sconto REAL NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 56
'...........................................................................................
'aggiunge campo su listino per derivazione
'...........................................................................................
function Aggiornamento__B2B__56(conn)
	Aggiornamento__B2B__56 = _
		" ALTER TABLE gtb_listini ADD listino_ancestor_id int NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 57
'...........................................................................................
'aggiunge relazione ancestor fra listini
'...........................................................................................
function Aggiornamento__B2B__57(conn)
	Aggiornamento__B2B__57 = _
		" ALTER TABLE gtb_listini ADD CONSTRAINT FK_gtb_listini_gtb_listini_ancestor " + _
		"		FOREIGN KEY (listino_ancestor_id) REFERENCES gtb_listini (listino_id) " + _
		"		NOT FOR REPLICATION ; " + _
		" ALTER TABLE gtb_listini NOCHECK CONSTRAINT FK_gtb_listini_gtb_listini_ancestor "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 58
'...........................................................................................
'rimuove propagazione modifiche delle relazioni (ON UPDATE CASCADE) che 
'interessano la tabella gtb_prezzi per poi inserire il trigger 
'...........................................................................................
function Aggiornamento__B2B__58(conn)
	Aggiornamento__B2B__58 = _
		" ALTER TABLE gtb_prezzi DROP CONSTRAINT FK_gtb_prezzi_grel_art_valori; " + _
		" ALTER TABLE gtb_prezzi DROP CONSTRAINT FK_gtb_prezzi_gtb_listini; " + _
		" ALTER TABLE gtb_prezzi ADD CONSTRAINT FK_gtb_prezzi_grel_art_valori " + _
		"		  FOREIGN KEY (	prz_variante_id	) REFERENCES grel_art_valori (rel_id) ON DELETE CASCADE ; " + _
		" ALTER TABLE gtb_prezzi ADD CONSTRAINT FK_gtb_prezzi_gtb_listini " + _
		"		  FOREIGN KEY ( prz_listino_id ) REFERENCES gtb_listini (listino_id) ON DELETE CASCADE "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 59
'...........................................................................................
'aggiunge trigger in cancellazione a tabella prezzi.
'...........................................................................................
function Aggiornamento__B2B__59(conn)
	Aggiornamento__B2B__59 = _
		" CREATE TRIGGER dbo.TRG_gtb_prezzi_FOR_DELETE " + vbCrLF + _
		"		ON gtb_prezzi " + vbCrLF + _
		"		FOR DELETE " + vbCrLF + _
		"	AS " + vbCrLF + _
		"		DELETE FROM gtb_prezzi WHERE prz_id IN ( " + vbCrLF + _
		"				SELECT gtb_prezzi.prz_id FROM (gtb_prezzi INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id=gtb_listini.listino_id) " + vbCrLF + _
		"				INNER JOIN deleted ON (gtb_prezzi.prz_variante_id = deleted.prz_variante_id AND gtb_listini.listino_ancestor_id=deleted.prz_listino_id) " + vbCrLF + _
		"				WHERE gtb_prezzi.prz_prezzo = deleted.prz_prezzo AND " + vbCrLF + _
		"				      gtb_prezzi.prz_visibile = deleted.prz_visibile AND " + vbCrLF + _
		"				      gtb_prezzi.prz_promozione = deleted.prz_promozione AND " + vbCrLF + _
		"				      gtb_prezzi.prz_variante_id = deleted.prz_variante_id AND " + vbCrLF + _
		"				      gtb_prezzi.prz_scontoQ_id = deleted.prz_scontoQ_id AND " + vbCrLF + _
		"				      gtb_prezzi.prz_iva_id = deleted.prz_iva_id AND " + vbCrLF + _
		"				      gtb_prezzi.prz_var_euro = deleted.prz_var_euro AND " + vbCrLF + _
		"				      gtb_prezzi.prz_var_sconto = deleted.prz_var_sconto) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 60
'...........................................................................................
'aggiunge trigger in inserimento a tabella prezzi.
'...........................................................................................
function Aggiornamento__B2B__60(conn)
	Aggiornamento__B2B__60 = _
		" CREATE TRIGGER dbo.TRG_gtb_prezzi_FOR_INSERT " + vbCrLF + _
		" 	ON gtb_prezzi " + vbCrLF + _
		" 	FOR INSERT NOT FOR REPLICATION " + vbCrLF + _
		" AS " + vbCrLF + _
		" 	INSERT INTO gtb_prezzi (prz_prezzo, prz_visibile, prz_promozione, prz_listino_id, prz_variante_id, prz_scontoQ_id, prz_iva_id, prz_var_euro, prz_var_sconto) " + vbCrLF + _
		" 	SELECT inserted.prz_prezzo, inserted.prz_visibile, inserted.prz_promozione, L_child.listino_id, inserted.prz_variante_id, " + vbCrLF + _
		" 	       inserted.prz_scontoQ_id, inserted.prz_iva_id, inserted.prz_var_euro, inserted.prz_var_sconto  " + vbCrLF + _
		" 		FROM inserted " + vbCrLF + _
		" 		INNER JOIN gtb_listini L_ancestor ON inserted.prz_listino_id = L_ancestor.listino_id " + vbCrLF + _
		" 		INNER JOIN gtb_listini L_child ON L_ancestor.listino_id = L_child.listino_ancestor_id " + vbCrLF + _
		" 		WHERE (SELECT COUNT(*)  " + vbCrLF + _
		" 		       		FROM gtb_prezzi  " + vbCrLF + _
		" 					WHERE gtb_prezzi.prz_listino_id=L_child.listino_id AND  " + vbCrLF + _
		" 					      gtb_prezzi.prz_variante_id=inserted.prz_variante_id)=0 "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 61
'...........................................................................................
'aggiunge campo che indica se il listino corrente ha anche "figli"
'...........................................................................................
function Aggiornamento__B2B__61(conn)
	Aggiornamento__B2B__61 = _
		" ALTER TABLE gtb_listini ADD listino_with_child bit NULL;" + _
		" UPDATE gtb_listini SET listino_with_child=CASE WHEN (SELECT COUNT(*) FROM gtb_listini L_child " + _
		" WHERE L_child.listino_ancestor_id = gtb_listini.listino_id)>0 THEN 1 ELSE 0 END; " + _
		" ALTER TABLE gtb_listini ALTER COLUMN listino_with_child bit NOT NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 62
'...........................................................................................
'rigenera viste per aggiornamento delle strutture
'...........................................................................................
function Aggiornamento__B2B__62(conn)
	Aggiornamento__B2B__62 = _
		"DROP VIEW dbo.gv_agenti ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_agenti AS " + vbCrLf + _
		"	SELECT * FROM gtb_agenti " + vbCrLf + _
		"		INNER JOIN tb_admin ON gtb_agenti.ag_admin_id = tb_admin.ID_admin " + vbCrLf + _
		"		INNER JOIN tb_Utenti ON gtb_agenti.ag_id = tb_Utenti.ut_ID " + vbCrLf + _
		"		INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_articoli ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_articoli AS" + vbCrLf + _
		"	SELECT * FROM gtb_articoli" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLf + _
		"		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id " + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_articoli_varianti ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_articoli_varianti AS " + vbCrLf + _
		"	SELECT TOP 100 PERCENT gtb_valori.*, grel_art_vv.*, gtb_varianti.* " + vbCrLf + _
		"		FROM grel_art_vv INNER JOIN gtb_valori ON grel_art_vv.rvv_val_id = gtb_valori.val_id " + vbCrLf + _
		"		INNER JOIN gtb_varianti ON gtb_valori.val_var_id = gtb_varianti.var_id " + vbCrLf + _
		"		ORDER BY gtb_varianti.var_ordine, gtb_varianti.var_id, gtb_valori.val_ordine" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_carichi ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_carichi AS " + vbCrLf + _
		"	SELECT * FROM grel_carichi_var " + vbCrLf + _
		"		INNER JOIN grel_art_valori ON grel_carichi_var.rcv_art_var_id = grel_art_valori.rel_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_CartDetail ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_CartDetail AS " + vbCrLf + _
		"	SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST, " + vbCrLf + _
		"			(SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT" + vbCrLf + _
		"		FROM grel_art_valori INNER JOIN gtb_dett_cart ON grel_art_valori.rel_id = gtb_dett_cart.dett_art_var_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_dettagli_ord ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_dettagli_ord AS " + vbCrLf + _
		"	SELECT * FROM gtb_dettagli_ord INNER JOIN " + vbCrLf + _
		"		grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id INNER JOIN " + vbCrLf + _
		"		gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_giacenza_pubblico ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_giacenza_pubblico AS" + vbCrLf + _
		"	SELECT SUM(gia_qta - gia_impegnato) AS giacenza, rel_id, MIN(rel_art_id) AS articolo_id FROM grel_giacenze" + vbCrLf + _
		"		INNER JOIN gtb_magazzini ON grel_giacenze.gia_magazzino_id = gtb_magazzini.mag_id" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		WHERE (gtb_magazzini.mag_disponibilita=1)" + vbCrLf + _
		"		OR ((SELECT COUNT(*) FROM gtb_magazzini WHERE mag_disponibilita=1)=0)" + vbCrLf + _
		"		GROUP BY rel_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_Giacenze_Varianti ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_Giacenze_Varianti AS " + vbCrLf + _
		"	SELECT * FROM grel_giacenze INNER JOIN " + vbCrLf + _
		"		grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id INNER JOIN " + vbCrLf + _
		"		grel_carichi_var ON grel_art_valori.rel_id = grel_carichi_var.rcv_art_var_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_inventario ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_inventario AS" + vbCrLf + _
		"	SELECT * FROM grel_giacenze" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_listini ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_listini AS" + vbCrLf + _
		"	SELECT * FROM grel_art_valori" + vbCrLf + _
		"		INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_listino_offerte ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_listino_offerte AS" + vbCrLf + _
		"	SELECT * FROM gtb_articoli" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLf + _
		"		INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLf + _
		"		INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"	WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0" + vbCrLf + _
		"		AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0" + vbCrLf + _
		"		AND tip_visibile=1" + vbCrLf + _
		"		AND tip_albero_visibile=1" + vbCrLf + _
		"		AND ISNULL(listino_offerte, 0)=1" + vbCrLf + _
		"		AND ISNULL(prz_visibile, 0)=1" + vbCrLf + _
		"		AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1)" + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_listino_vendita ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_listino_vendita AS" + vbCrLf + _
		"	SELECT * FROM gtb_articoli" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLf + _
		"		INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLf + _
		"		INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLf + _
		"	WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0" + vbCrLf + _
		"		AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLf + _
		"		AND tip_visibile=1" + vbCrLf + _
		"		AND tip_albero_visibile=1" + vbCrLf + _
		"		AND ( ( ISNULL(listino_offerte, 0)=1" + vbCrLf + _
		"				AND ISNULL(prz_visibile, 0)=1" + vbCrLf + _
		"				AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) )" + vbCrLf + _
		"				OR (ISNULL(listino_offerte, 0)=0" + vbCrLf + _
		"				AND prz_variante_id NOT IN (SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id" + vbCrLf + _
		"				WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1" + vbCrLf + _
		"				AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) ))) " + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_rivenditori ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_rivenditori AS " + vbCrLf + _
		"	SELECT * FROM gtb_rivenditori " + vbCrLf + _
		"		INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_utenti.ut_ID " + vbCrLf + _
		"		INNER JOIN tb_Indirizzario ON tb_utenti.ut_NextCom_ID = tb_indirizzario.IDElencoIndirizzi " + vbCrLf + _
		"		INNER JOIN gtb_valute ON gtb_rivenditori.riv_valuta_id = gtb_valute.valu_id"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 63
'...........................................................................................
'aggiunge chiave primaria a tabella listini ed indice
'...........................................................................................
function Aggiornamento__B2B__63(conn)
	Aggiornamento__B2B__63 = _
		" ALTER TABLE gtb_prezzi ADD CONSTRAINT PK_gtb_prezzi PRIMARY KEY CLUSTERED (prz_id); " + _
		" CREATE INDEX IDX__gtb_prezzi ON gtb_prezzi (prz_listino_id, prz_variante_id) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 64
'...........................................................................................
'aggiunge indici tabelle varie
'...........................................................................................
function Aggiornamento__B2B__64(conn)
	Aggiornamento__B2B__64 = _
		" CREATE INDEX IDX__gtb_articoli_ordinati ON gtb_articoli_ordinati (ao_ut_id, ao_variante_id) ; "  + _
		" CREATE INDEX IDX__gtb_wish_list ON gtb_wish_list (wish_ut_id, wish_variante_id) ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 65
'...........................................................................................
'aggiunge indici tabelle varie
'...........................................................................................
function Aggiornamento__B2B__65(conn)
	Aggiornamento__B2B__65 = _
		" CREATE INDEX IDX__grel_art_valori__rel_art_id ON grel_art_valori (rel_art_id) ; " + _
		" CREATE INDEX IDX__grel_art_vv__rvv_art_var_id ON grel_art_vv (rvv_art_var_id) ; " + _
		" CREATE INDEX IDX__gtb_articoli__art_tipologia_id ON gtb_articoli (art_tipologia_id) ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 66
'...........................................................................................
'aggiunge indici tabelle varie
'...........................................................................................
function Aggiornamento__B2B__66(conn)
	Aggiornamento__B2B__66 = _
		" CREATE INDEX IDX__gtb_codici ON gtb_codici (cod_lista_id, cod_variante_id) ; " + _
		" CREATE INDEX IDX__grel_giacenze__gia_art_var_id ON grel_giacenze (gia_art_var_id) ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 67
'...........................................................................................
'aggiunge dati alla tabella agenti, li imposta ed aggiorna struttura tabella agenti
'...........................................................................................
function Aggiornamento__B2B__67(conn)
	Aggiornamento__B2B__67 = _
		" ALTER TABLE gtb_agenti ADD " + _
		" 	ag_range_sconto_massimo int NULL, " + _
		" 	ag_supervisore bit NULL ; " + _
		" UPDATE gtb_agenti SET ag_supervisore=0; " + _
		" ALTER TABLE gtb_agenti ALTER COLUMN ag_supervisore BIT NULL ; " +_
		" DROP VIEW dbo.gv_agenti ; " + _
		"CREATE VIEW dbo.gv_agenti AS " + vbCrLf + _
		"	SELECT * FROM gtb_agenti " + vbCrLf + _
		"		INNER JOIN tb_admin ON gtb_agenti.ag_admin_id = tb_admin.ID_admin " + vbCrLf + _
		"		INNER JOIN tb_Utenti ON gtb_agenti.ag_id = tb_Utenti.ut_ID " + vbCrLf + _
		"		INNER JOIN tb_Indirizzario ON tb_Utenti.ut_NextCom_ID = tb_Indirizzario.IDElencoIndirizzi" + vbCrLf
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 68
'...........................................................................................
'corregge bug su vista listini di vendita
'...........................................................................................
function Aggiornamento__B2B__68(conn)
	Aggiornamento__B2B__68 = _
		" DROP VIEW dbo.gv_listino_vendita; " + vbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS " + vbCrLf + _
		"		SELECT * FROM gtb_articoli " + vbCrLf + _
		" 		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLf + _
		" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLf + _
		" 		INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLf + _
		"				  AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLf + _
		"				  AND tip_visibile=1 " + vbCrLf + _
		"				  AND tip_albero_visibile=1 " + vbCrLf + _
		"				  AND prz_visibile=1 " + vbCrLf + _
		"				  AND ( ( ISNULL(listino_offerte, 0)=1 " + vbCrLf + _
		"				  		  AND ISNULL(prz_visibile, 0)=1 " + vbCrLf + _
		"						  AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) " + vbCrLf + _
		"						) " + vbCrLf + _
		"						OR (ISNULL(listino_offerte, 0)=0 " + vbCrLf + _
		"				  AND prz_variante_id NOT IN ( " + vbCrLf + _
		"						SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id " + vbCrLf + _
		"						WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 " + vbCrLf + _
		"						AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) " + vbCrLf + _
		"											 ) " + vbCrLf + _
		"						) " + vbCrLf + _
		"					  ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 69
'...........................................................................................
'riorganizza log ordini
'...........................................................................................
function Aggiornamento__B2B__69(conn)
	Aggiornamento__B2B__69 = _
		" ALTER TABLE glog_ordini ADD " + _
		"	log_operazione_id int NULL, " + _
		" 	log_extra_byte int NULL ; " + _
		" CREATE TABLE dbo.glog_ordini_operazioni ( " + _
		"	op_id int  IDENTITY (1, 1) NOT NULL , " + _
		"	op_descrizione nvarchar(250) NULL );" + _
		" INSERT INTO glog_ordini_operazioni (op_descrizione) VALUES('Non codificata') ;"  + _
		" UPDATE glog_ordini SET log_operazione_id=(SELECT TOP 1 op_id FROM glog_ordini_operazioni) ; " + _
		" ALTER TABLE glog_ordini_operazioni ADD CONSTRAINT PK_glog_ordini_operazioni PRIMARY KEY CLUSTERED (op_id); " + _
		" ALTER TABLE glog_ordini ADD CONSTRAINT FK_glog_ordini__glog_ordini_operazioni " + _
		"	FOREIGN KEY ( log_operazione_id ) " + _
		"	REFERENCES glog_ordini_operazioni ( op_ID) ON DELETE CASCADE ON UPDATE CASCADE ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 70
'...........................................................................................
'riorganizza log ordini
'...........................................................................................
function Aggiornamento__B2B__70(conn)
	Aggiornamento__B2B__70 = _
		" ALTER TABLE glog_ordini ADD log_operazione_extra_id int NULL ;" + _
		" ALTER TABLE glog_ordini DROP COLUMN log_extra_byte ; " + _
		" ALTER TABLE glog_ordini DROP COLUMN log_operazione ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 71
'...........................................................................................
'aggiorna collegamento rivenditori, shopping cart
'...........................................................................................
function Aggiornamento__B2B__71(conn)
	Aggiornamento__B2B__71 = _
		" DELETE FROM gtb_shopping_cart WHERE sc_riv_id NOT IN (SELECT riv_id FROM gtb_rivenditori); " + _
		" ALTER TABLE gtb_shopping_cart ADD CONSTRAINT FK_gtb_shopping_cart__gtb_rivenditori " + _
		"	FOREIGN KEY ( sc_riv_id ) " + _
		"	REFERENCES gtb_rivenditori ( riv_ID ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 72
'...........................................................................................
'ottimizza vista per calcolo giacenza
'...........................................................................................
function Aggiornamento__B2B__72(conn)
	Aggiornamento__B2B__72 = _
		" if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[gv_giacenza_pubblico]') and OBJECTPROPERTY(id, N'IsView') = 1) " + vbCrLf + _
		"		DROP VIEW dbo.gv_giacenza_pubblico; " + _
		" CREATE VIEW dbo.gv_giacenza_pubblico AS " + vbCrLF + _
		"	SELECT SUM(gia_qta - gia_impegnato) AS giacenza, rel_id, MIN(rel_art_id) AS articolo_id  " + vbCrLF + _
		"	FROM grel_giacenze INNER JOIN gtb_magazzini ON grel_giacenze.gia_magazzino_id = gtb_magazzini.mag_id  " + vbCrLF + _
		"	INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id  " + vbCrLF + _
		"	WHERE (gtb_magazzini.mag_disponibilita=1) " + vbCrLF + _
		"	GROUP BY rel_id "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 73
'...........................................................................................
'ottimizza vista per inventario
'...........................................................................................
function Aggiornamento__B2B__73(conn)
	Aggiornamento__B2B__73 = _
		" DROP VIEW dbo.gv_inventario; " + _
		" CREATE VIEW dbo.gv_inventario AS " + vbCrLF + _
		" 	SELECT * FROM grel_giacenze " + vbCrLF + _
		" 	INNER JOIN grel_art_valori ON grel_giacenze.gia_art_var_id = grel_art_valori.rel_id " + vbCrLF + _
		" 	INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 74
'...........................................................................................
'rimuove vista calcolo giacenza
'...........................................................................................
function Aggiornamento__B2B__74(conn)
	Aggiornamento__B2B__74 = _
		" DROP VIEW dbo.gv_giacenza_pubblico; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 75
'...........................................................................................
'aggiornamento struttura ordini e shopping-cart con rimozione vincolo su relazione tra
'riga di dettaglio e articolo ed aggiorna relativo vincolo
'...........................................................................................
function Aggiornamento__B2B__75(conn)
	Aggiornamento__B2B__75 = _
		" ALTER TABLE gtb_dett_cart DROP CONSTRAINT FK_gtb_dett_cart_grel_art_valori1 ; " + _
		" ALTER TABLE gtb_dett_cart ADD CONSTRAINT FK_gtb_dett_cart_grel_art_valori " + _
		"	FOREIGN KEY (dett_art_var_id) REFERENCES grel_art_valori (rel_id) " + _
		"   ON DELETE NO ACTION ON UPDATE NO ACTION " + _
		"	NOT FOR REPLICATION ; " + _
		" ALTER TABLE gtb_dett_cart NOCHECK CONSTRAINT FK_gtb_dett_cart_grel_art_valori; " + _
		" ALTER TABLE gtb_dettagli_ord DROP CONSTRAINT FK_gtb_dettagli_ord_grel_art_valori ; " + _
		" ALTER TABLE gtb_dettagli_ord ADD CONSTRAINT FK_gtb_dettagli_ord_grel_art_valori " + _
		"	FOREIGN KEY (det_art_var_id) REFERENCES grel_art_valori (rel_id) " + _
		"   ON DELETE NO ACTION ON UPDATE NO ACTION " + _
		"	NOT FOR REPLICATION ; " + _
		" ALTER TABLE gtb_dettagli_ord NOCHECK CONSTRAINT FK_gtb_dettagli_ord_grel_art_valori; " + _
		DropObject(conn, "gv_dettagli_ord","VIEW") + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + vbCrLF + _
		"	SELECT * FROM gtb_dettagli_ord " + vbCrLF + _
		"		LEFT JOIN grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id " + vbCrLF + _
		"		LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id ; " + vbCrLF + _
		DropObject(conn, "gv_CartDetail","VIEW") + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrLF + _
		"	SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST, " + vbCrLF + _
		"			  (SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrLF + _
		"		FROM gtb_dett_cart LEFT JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"		LEFT JOIN grel_art_valori ON gtb_dett_cart.dett_art_var_id = grel_art_valori.rel_id " + vbCrLF + _
		"		LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLF + _
		"		LEFT JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
		"	WHERE (gtb_dett_cart.dett_art_var_id IS NULL) OR " + vbCrLf + _
		"		( ISNULL(gtb_articoli.art_disabilitato, 0)=0 AND " + vbCrLf + _
		"		  ISNULL(grel_art_valori.rel_disabilitato,0)=0 AND " + vbCrLf + _
		"		  ISNULL(gtb_tipologie.tip_albero_visibile, 0) = 1 AND " + vbCrLf + _
		"		  ISNULL(gtb_tipologie.tip_visibile, 0)= 1 ) ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 76
'...........................................................................................
'aggiunge campo descrizione e campo note a dettagli ordine e dettagli shopping cart
'...........................................................................................
function Aggiornamento__B2B__76(conn)
	Aggiornamento__B2B__76 = _
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_descr_IT nvarchar(500) NULL, " + _
		"	dett_descr_EN nvarchar(500) NULL, " + _
		"	dett_descr_FR nvarchar(500) NULL, " + _
		"	dett_descr_DE nvarchar(500) NULL, " + _
		"	dett_descr_ES nvarchar(500) NULL, " + _
		"	dett_note nvarchar(500) NULL ; " + _
		" ALTER TABLE gtb_dettagli_ord ADD " + _
		"	det_descr_IT nvarchar(500) NULL, " + _
		"	det_descr_EN nvarchar(500) NULL, " + _
		"	det_descr_FR nvarchar(500) NULL, " + _
		"	det_descr_DE nvarchar(500) NULL, " + _
		"	det_descr_ES nvarchar(500) NULL, " + _
		"	det_note nvarchar(500) NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 77
'...........................................................................................
'aggiunge campo per codifica promozioni
'...........................................................................................
function Aggiornamento__B2B__77(conn)
	Aggiornamento__B2B__77 = _
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_cod_promozione nvarchar(50) NULL ;" + _
		" ALTER TABLE gtb_dettagli_ord ADD " + _
		"	det_cod_promozione nvarchar(50) NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 78
'...........................................................................................
'rimuove tabelle di gestione punteggi e campi relativi
'...........................................................................................
function Aggiornamento__B2B__78(conn)
	Aggiornamento__B2B__78 = _
		" ALTER TABLE dbo.gtb_punteggi DROP CONSTRAINT FK_gtb_punteggi_grel_art_valori ; " + _
		" ALTER TABLE dbo.gtb_punteggi DROP CONSTRAINT FK_gtb_punteggi_gtb_lista_punteggi ; " + _
		" ALTER TABLE dbo.gtb_rivenditori DROP CONSTRAINT FK_gtb_rivenditori_gtb_lista_punteggi ; " + _
		DropObject(conn, "gtb_punteggi", "TABLE") + _
		DropObject(conn, "gtb_lista_punteggi", "TABLE") + _
		" ALTER TABLE gtb_rivenditori DROP COLUMN riv_punteggio_id ; " + _
		" ALTER TABLE gtb_rivenditori DROP COLUMN riv_punteggio ; " + _
		" ALTER TABLE grel_art_valori DROP COLUMN rel_punteggio ; " + _
		" ALTER TABLE gtb_articoli DROP COLUMN art_punteggio ; " + _
		" ALTER TABLE gtb_dettagli_ord DROP COLUMN det_punteggio ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 79
'...........................................................................................
'ClassCategorie: aggiunge il campo per la gestione della lista degli IDs dei padri
'...........................................................................................
function AggiornamentoSpeciale__B2B__79(DB, rs, version)
    CALL AggiornamentoSpeciale__FRAMEWORK_CORE__ListaPadriCategorie(DB, rs, version, "gtb_tipologie", "tip")
    AggiornamentoSpeciale__B2B__79 = "SELECT * FROM AA_versione"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 80
'...........................................................................................
'aggiunge tipi di riga d'ordine e relativi descrittori
'...........................................................................................
function Aggiornamento__B2B__80(conn)
	Aggiornamento__B2B__80 = _
		" CREATE TABLE dbo.gtb_dettagli_ord_tipo( " + _
           "   dot_id INT IDENTITY (1, 1) NOT NULL , " + _
           "   dot_nome_it nvarchar (255) NULL , " + _
           "   dot_nome_en nvarchar (255) NULL , " + _
           "   dot_nome_fr nvarchar (255) NULL , " + _
           "   dot_nome_de nvarchar (255) NULL , " + _
           "   dot_nome_es nvarchar (255) NULL , " + _
           "   CONSTRAINT PK_gtb_dettagli_ord_tipo PRIMARY KEY ( dot_id ) " + _
           " ) ; " + _
           " CREATE TABLE dbo.gtb_dettagli_ord_des ( " + _
           "    dod_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
           "    dod_nome_it nvarchar (255) NULL , " + vbCrLf + _
           "    dod_nome_en nvarchar (255) NULL , " + vbCrLf + _
           "    dod_nome_fr nvarchar (255) NULL , " + vbCrLf + _
           "    dod_nome_es nvarchar (255) NULL , " + vbCrLf + _
           "    dod_nome_de nvarchar (255) NULL , " + vbCrLf + _
           "    dod_unita_it nvarchar (50) NULL , " + vbCrLf + _
           "    dod_unita_en nvarchar (50) NULL , " + vbCrLf + _
           "    dod_unita_fr nvarchar (50) NULL , " + vbCrLf + _
           "    dod_unita_es nvarchar (50) NULL , " + vbCrLf + _
           "    dod_unita_de nvarchar (50) NULL , " + vbCrLf + _
           "    dod_tipo int NULL , " + _
           "   CONSTRAINT PK_gtb_dettagli_ord_des PRIMARY KEY ( dod_id ) " + _
           " ) ; " + _
           "CREATE TABLE dbo.grel_dettagli_ord_tipo_des ( " + vbCrLf + _
           "    rtd_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
           "    rtd_des_id int NULL , " + vbCrLf + _
           "    rtd_ordine int NULL , " + vbCrLf + _
           "    rtd_tipo_id int NULL , " + _
           "    CONSTRAINT PK_grel_dettagli_ord_tipo_des PRIMARY KEY ( rtd_id ), " + _
           "    CONSTRAINT FK_grel_dettagli_ord_tipo_des__gtb_dettagli_ord_des FOREIGN KEY ( rtd_des_id ) " + vbCrLf + _
           "       REFERENCES dbo.gtb_dettagli_ord_des ( dod_id ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
           "    CONSTRAINT FK_grel_dettagli_ord_tipo_des__gtb_dettagli_ord_tipo FOREIGN KEY ( rtd_tipo_id ) " + vbCrLf + _
           "       REFERENCES dbo.gtb_dettagli_ord_tipo ( dot_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
           " ) ; " + _
           "CREATE TABLE dbo.grel_dettagli_ord_des_value ( " + vbCrLf + _
           "    rel_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
           "    rel_dett_ord_id int NULL , " + vbCrLf + _
           "    rel_des_id int NULL , " + vbCrLf + _
           "    rel_des_value_de ntext NULL , " + vbCrLf + _
           "    rel_des_value_en ntext NULL , " + vbCrLf + _
           "    rel_des_value_es ntext NULL , " + vbCrLf + _
           "    rel_des_value_fr ntext NULL , " + vbCrLf + _
           "    rel_des_value_it ntext NULL , " + _
           "   CONSTRAINT PK_grel_dettagli_ord_des_value PRIMARY KEY ( rel_id ), " + _
           "    CONSTRAINT FK_grel_dettagli_ord_des_value__gtb_dettagli_ord FOREIGN KEY ( rel_dett_ord_id ) " + vbCrLf + _
           "       REFERENCES dbo.gtb_dettagli_ord ( det_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
           "    CONSTRAINT FK_grel_dettagli_ord_des_value__gtb_dettagli_ord_des FOREIGN KEY ( rel_des_id ) " + vbCrLf + _
           "       REFERENCES dbo.gtb_dettagli_ord_des ( dod_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
           " ) ; " + _
           "CREATE TABLE dbo.grel_dett_cart_des_value ( " + vbCrLf + _
           "    rel_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
           "    rel_det_cart_id int NULL , " + vbCrLf + _
           "    rel_des_id int NULL , " + vbCrLf + _
           "    rel_des_value_de ntext NULL , " + vbCrLf + _
           "    rel_des_value_en ntext NULL , " + vbCrLf + _
           "    rel_des_value_es ntext NULL , " + vbCrLf + _
           "    rel_des_value_fr ntext NULL , " + vbCrLf + _
           "    rel_des_value_it ntext NULL , " + _
           "    CONSTRAINT PK_grel_dett_cart_des_value PRIMARY KEY ( rel_id ), " + _
           "    CONSTRAINT FK_grel_dett_cart_des_value__gtb_dett_cart FOREIGN KEY ( rel_det_cart_id ) " + vbCrLf + _
           "       REFERENCES dbo.gtb_dett_cart ( dett_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
           "    CONSTRAINT FK_grel_dett_cart_des_value__gtb_dettagli_ord_des FOREIGN KEY ( rel_des_id ) " + vbCrLf + _
           "       REFERENCES dbo.gtb_dettagli_ord_des ( dod_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
           " ) ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 81
'...........................................................................................
'collega le righe d'ordine e della shopping cart alla tipologia
'...........................................................................................
function Aggiornamento__B2B__81(conn)
	Aggiornamento__B2B__81 = _
	    " ALTER TABLE gtb_dett_cart ADD " + _
	    "   dett_tipo_id INT NULL ; " + _
	    " ALTER TABLE gtb_dett_cart WITH NOCHECK ADD " + _
	    "   CONSTRAINT FK_gtb_dett_cart__gtb_dettagli_ord_tipo FOREIGN KEY ( dett_tipo_id ) " + _
	    "   REFERENCES gtb_dettagli_ord_tipo ( dot_ID ) ; " + _
	    " ALTER TABLE gtb_dett_cart NOCHECK CONSTRAINT FK_gtb_dett_cart__gtb_dettagli_ord_tipo; " + _
	    " ALTER TABLE gtb_dettagli_ord ADD " + _
	    "   det_tipo_id INT NULL ; " + _
	    " ALTER TABLE gtb_dettagli_ord WITH NOCHECK ADD " + _
	    "   CONSTRAINT FK_gtb_dettagli_ord__gtb_dettagli_ord_tipo FOREIGN KEY ( det_tipo_id ) " + _
	    "   REFERENCES gtb_dettagli_ord_tipo ( dot_ID ) ; " + _
	    " ALTER TABLE gtb_dettagli_ord NOCHECK CONSTRAINT FK_gtb_dettagli_ord__gtb_dettagli_ord_tipo; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 82
'...........................................................................................
'aggiornamento che rimuove l'integrita' referenziale tra Rivenditori e Shopping cart
'per permettere la gestione della shopping cart B2C solo con id di sessione
'...........................................................................................
function Aggiornamento__B2B__82(conn)
	Aggiornamento__B2B__82 = _
	    " ALTER TABLE gtb_shopping_cart DROP CONSTRAINT FK_gtb_shopping_cart__gtb_rivenditori; " + _
	    " ALTER TABLE gtb_shopping_cart WITH NOCHECK ADD CONSTRAINT FK_gtb_shopping_cart__gtb_rivenditori " + _
	    "   FOREIGN KEY (sc_riv_id) REFERENCES gtb_rivenditori(riv_id) " + _
	    "   ON UPDATE NO ACTION ON DELETE NO ACTION; " + _
	    " ALTER TABLE gtb_shopping_cart NOCHECK CONSTRAINT FK_gtb_shopping_cart__gtb_rivenditori; " + _
	    " ALTER TABLE gtb_shopping_cart ADD " + _
	    "   sc_session_id nvarchar(250) NULL, " + _
	    "   sc_ip_address nvarchar(15) NULL ; " + _
	    " ALTER TABLE gtb_shopping_cart ALTER COLUMN sc_ut_id INT NULL; " + _
	    " ALTER TABLE gtb_shopping_cart ALTER COLUMN sc_riv_id INT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 83
'...........................................................................................
'aggiornamento che corregge ed aggiunge i campi per la tracciatura delle modifiche ai dati
'degli articoli
'...........................................................................................
function Aggiornamento__B2B__83(conn)
	Aggiornamento__B2B__83 = _
	    " ALTER TABLE gtb_articoli ADD " + _
	    "   art_insData	datetime NULL, " + _
	    "   art_insAdmin_id	int	NULL, " + _
	    "   art_modData	datetime NULL, " + _
	    "   art_modAdmin_id	int NULL ; " + _
	    " ALTER TABLE grel_art_valori ADD " + _
	    "   rel_insData	datetime NULL, " + _
	    "   rel_insAdmin_id	int	NULL, " + _
	    "   rel_modData	datetime NULL, " + _
	    "   rel_modAdmin_id	int NULL ; " + _
	    " UPDATE gtb_articoli SET art_insData = art_data_insert, art_modData=art_data_update, " + _
	                            " art_insAdmin_id = (SELECT TOP 1 id_admin FROM tb_admin WHERE admin_login LIKE '%SISTEMA%'), " + _
	                            " art_modAdmin_id = (SELECT TOP 1 id_admin FROM tb_admin WHERE admin_login LIKE '%SISTEMA%') ;" + _
	    " UPDATE grel_art_valori SET rel_insData = (SELECT art_insData FROM gtb_articoli WHERE gtb_articoli.art_id = rel_art_id ), " + _
	                               " rel_modData = (SELECT art_modData FROM gtb_articoli WHERE gtb_articoli.art_id = rel_art_id ), " + _
	                               " rel_insAdmin_id = (SELECT art_insAdmin_id FROM gtb_articoli WHERE gtb_articoli.art_id = rel_art_id ), " + _
	                               " rel_modAdmin_id = (SELECT art_modAdmin_id FROM gtb_articoli WHERE gtb_articoli.art_id = rel_art_id ) ; " + _
	    " ALTER TABLE gtb_articoli DROP COLUMN art_data_insert; " + _
	    " ALTER TABLE gtb_articoli DROP COLUMN art_data_update; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 84
'...........................................................................................
'uniforma i descrittori dei dettagli shoppingcart e ordine a quelli standard
'...........................................................................................
function Aggiornamento__B2B__84(conn)
		Aggiornamento__B2B__84 = _
            " ALTER TABLE dbo.grel_dettagli_ord_des_value DROP CONSTRAINT FK_grel_dettagli_ord_des_value__gtb_dettagli_ord; " + vbCrLf + _
			" ALTER TABLE dbo.grel_dettagli_ord_des_value DROP CONSTRAINT FK_grel_dettagli_ord_des_value__gtb_dettagli_ord_des; " + vbCrLf + _
			" DROP TABLE dbo.grel_dettagli_ord_des_value;" + vbCrLf + _
			" CREATE TABLE dbo.grel_dettagli_ord_des_value ( " + vbCrLf + _
            "    rel_des_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
            "    rel_des_dett_ord_id int NULL , " + vbCrLf + _
            "    rel_des_descrittore_id int NULL , " + vbCrLf + _
			"    rel_des_valore_de NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_en NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_es NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_fr NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_it NVARCHAR(250) NULL , " + _
            "    rel_des_memo_de ntext NULL , " + vbCrLf + _
            "    rel_des_memo_en ntext NULL , " + vbCrLf + _
            "    rel_des_memo_es ntext NULL , " + vbCrLf + _
            "    rel_des_memo_fr ntext NULL , " + vbCrLf + _
            "    rel_des_memo_it ntext NULL , " + vbCrLf + _
            "    CONSTRAINT PK_grel_dettagli_ord_des_value PRIMARY KEY ( rel_des_id ), " + vbCrLf + _
            "    CONSTRAINT FK_grel_dettagli_ord_des_value__gtb_dettagli_ord FOREIGN KEY ( rel_des_dett_ord_id ) " + vbCrLf + _
            "       REFERENCES dbo.gtb_dettagli_ord ( det_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
            "    CONSTRAINT FK_grel_dettagli_ord_des_value__gtb_dettagli_ord_des FOREIGN KEY ( rel_des_descrittore_id ) " + vbCrLf + _
            "       REFERENCES dbo.gtb_dettagli_ord_des ( dod_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
            " ) ; " + vbCrLf + _
			" ALTER TABLE dbo.grel_dett_cart_des_value DROP CONSTRAINT FK_grel_dett_cart_des_value__gtb_dett_cart; " + vbCrLf + _
			" ALTER TABLE dbo.grel_dett_cart_des_value DROP CONSTRAINT FK_grel_dett_cart_des_value__gtb_dettagli_ord_des; " + vbCrLf + _
			" DROP TABLE dbo.grel_dett_cart_des_value; " + vbCrLf + _
			" CREATE TABLE dbo.grel_dett_cart_des_value ( " + vbCrLf + _
            "    rel_des_id int IDENTITY (1, 1) NOT NULL , " + vbCrLf + _
            "    rel_des_dett_cart_id int NULL , " + vbCrLf + _
            "    rel_des_descrittore_id int NULL , " + vbCrLf + _
			"    rel_des_valore_de NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_en NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_es NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_fr NVARCHAR(250) NULL , " + vbCrLf + _
            "    rel_des_valore_it NVARCHAR(250) NULL , " + _
            "    rel_des_memo_de ntext NULL , " + vbCrLf + _
            "    rel_des_memo_en ntext NULL , " + vbCrLf + _
            "    rel_des_memo_es ntext NULL , " + vbCrLf + _
            "    rel_des_memo_fr ntext NULL , " + vbCrLf + _
            "    rel_des_memo_it ntext NULL , " + vbCrLf + _
            "    CONSTRAINT PK_grel_dett_cart_des_value PRIMARY KEY ( rel_des_id ), " + vbCrLf + _
            "    CONSTRAINT FK_grel_dett_cart_des_value__gtb_dett_cart FOREIGN KEY ( rel_des_dett_cart_id ) " + vbCrLf + _
            "       REFERENCES dbo.gtb_dett_cart ( dett_ID ) ON DELETE CASCADE  ON UPDATE CASCADE , " + vbCrLf + _
            "    CONSTRAINT FK_grel_dett_cart_des_value__gtb_dettagli_ord_des FOREIGN KEY ( rel_des_descrittore_id ) " + vbCrLf + _
            "       REFERENCES dbo.gtb_dettagli_ord_des ( dod_id ) ON DELETE CASCADE  ON UPDATE CASCADE  " + vbCrLf + _
            " ) ; " + vbCrLf + _
			" ALTER TABLE dbo.grel_dettagli_ord_tipo_des DROP CONSTRAINT FK_grel_dettagli_ord_tipo_des__gtb_dettagli_ord_des;" + vbCrLf + _
			" ALTER TABLE dbo.grel_dettagli_ord_tipo_des DROP COLUMN rtd_des_id;" + vbCrLf + _
			" ALTER TABLE dbo.grel_dettagli_ord_tipo_des ADD rtd_descrittore_id INT NULL;" + vbCrLf + _
			" ALTER TABLE dbo.grel_dettagli_ord_tipo_des ADD" + vbCrLf + _
			" 	 CONSTRAINT FK_grel_dettagli_ord_tipo_des__gtb_dettagli_ord_des FOREIGN KEY ( rtd_descrittore_id ) " + vbCrLf + _
            "       REFERENCES dbo.gtb_dettagli_ord_des ( dod_id ) ON DELETE CASCADE  ON UPDATE CASCADE"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 85
'...........................................................................................
'aggiornamento che aggiunge colonna "giacenza iniziale" al magazzino
'...........................................................................................
function Aggiornamento__B2B__85(conn)
		Aggiornamento__B2B__85 = _
			"ALTER TABLE grel_giacenze ADD gia_iniziale INT NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 86
'...........................................................................................
'modifica il tipo della colonna che indica la quantita' nella riga d'ordine.
'...........................................................................................
function Aggiornamento__B2B__86(conn)
	Aggiornamento__B2B__86 = _
		"ALTER TABLE gtb_dettagli_ord ALTER COLUMN det_qta INT NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 87
'...........................................................................................
'aggiunge i dati della richiesta HTTP all'ordine
'...........................................................................................
function Aggiornamento__B2B__87(conn)
	Aggiornamento__B2B__87 = _
		" ALTER TABLE " + SQL_Dbo(conn) + "gtb_ordini ADD" + vbCrLf + _
		"	ord_ip_address " + SQL_CharField(conn, 39) + " NULL," + vbCrLf + _
		"	ord_request " + SQL_CharField(conn, 0) + " NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 88
'...........................................................................................
'aggiunge il campo visibile alle foto
'...........................................................................................
function Aggiornamento__B2B__88(conn)
	Aggiornamento__B2B__88 = _
		" ALTER TABLE " + SQL_Dbo(conn) + "gtb_art_foto ADD" + vbCrLf + _
		"	fo_visibile BIT NULL;" + vbCrLf + _
		" UPDATE gtb_art_foto SET fo_visibile = 1;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 89
'..........................................................................................................................
'aggiunge la tabella per la gestione delle spese di spedizione
'..........................................................................................................................
function Aggiornamento__B2B__89(conn)
	Aggiornamento__B2B__89 = _
		"CREATE TABLE dbo.gtb_spese_spedizione ( " + _
		"   sp_id " + SQL_PrimaryKey(conn, "gtb_spese_spedizione") + ", " + _
		"   sp_importo_euro real NOT NULL , " + _
		SQL_MultiLanguageField("	sp_area_nome_<lingua> " + SQL_CharField(Conn, 255)) + ", " + _
		SQL_MultiLanguageField("	sp_condizioni_<lingua> " + SQL_CharField(Conn, 0)) + _
		" ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 90
'...........................................................................................
'rimuove vecchie viste, stored procedure e funzioni non piu' utilizzate
'...........................................................................................
function Aggiornamento__B2B__90(conn)
	Aggiornamento__B2B__90 = _
		DropObject(conn, "fn_andycott_variante_listino", "FUNCTION") + _
		DropObject(conn, "gv_articoli_fornitori_F000078", "VIEW") + _
		DropObject(conn, "gv_Giacenze_Varianti", "VIEW") + _
		DropObject(conn, "gv_inventario", "VIEW") + _
		DropObject(conn, "sp_andycott_elenco_promozioni", "PROCEDURE")
		
	if lcase(GetDatabaseName(conn))<> "andycott" AND _
	   lcase(GetDatabaseName(conn))<> "b2b" then
		'se non sono sul database andycott rimuovo tabelle specifiche
		Aggiornamento__B2B__90 = Aggiornamento__B2B__90 + _
			DropObject(conn, "andycott_tb_attivita", "TABLE") + _
			DropObject(conn, "andycott_tb_dipendenti", "TABLE") + _
			DropObject(conn, "andycott_tb_piani", "TABLE") + _
			DropObject(conn, "andycott_tb_province", "TABLE") + _
			DropObject(conn, "andycott_tb_societa", "TABLE") + _
			DropObject(conn, "andycott_tb_tecnologie", "TABLE") + _
			DropObject(conn, "sp_andycott_articolo_gianceza", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_articoli_cliente_art_cod_int_ASC", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_articoli_cliente_art_cod_int_DESC", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_articoli_cliente_art_nome_ASC", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_articoli_cliente_art_nome_DESC", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_articoli_cliente_prezzo_ASC", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_articoli_cliente_prezzo_DESC", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_offerte", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_elenco_raggruppamenti_cliente", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_ordinati_articoli", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_ordinati_categorie", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_preferiti_articoli", "PROCEDURE") + _
			DropObject(conn, "sp_andycott_preferiti_categorie", "PROCEDURE")
		if lcase(GetDatabaseName(conn))<> "andycott" then
			Aggiornamento__B2B__90 = Aggiornamento__B2B__90 + _
				DropObject(conn, "andycott_gtb_commenti", "TABLE") + _
				DropObject(conn, "andycott_gtb_rivenditori_promo", "TABLE")
		end if
	end if
		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 91
'..........................................................................................................................
'aggiunge flag per gestione vendibilita' dell'articolo a livello di catalogo e listino
'..........................................................................................................................
function Aggiornamento__B2B__91(conn)
	Aggiornamento__B2B__91 = _
		" ALTER TABLE gtb_articoli ADD art_non_vendibile BIT NULL ; " + _
		" ALTER TABLE grel_art_valori ADD rel_non_vendibile BIT NULL ; " + _
		" ALTER TABLE gtb_prezzi ADD prz_non_vendibile BIT NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 92
'..........................................................................................................................
'corregge e completa la gestione delle valute
'..........................................................................................................................
function Aggiornamento__B2B__92(conn)
	Aggiornamento__B2B__92 = _
		" ALTER TABLE gtb_valute ADD " + _
		"	valu_num_decimali INT null, " + _
		"	valu_sep_decimali nvarchar(1) NULL, " + _
		"	valu_sep_migliaia nvarchar(1) NULL" + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 93
'..........................................................................................................................
'aggiunge gestione valuta di default per b2c
'..........................................................................................................................
function Aggiornamento__B2B__93(conn)
	Aggiornamento__B2B__93 = _
		" ALTER TABLE gtb_valute ADD " + _
		"	valu_B2C bit NULL " + _
		" ; " + _
		" UPDATE gtb_valute SET valu_B2C = CASE WHEN valu_codice LIKE 'EUR' THEN 1 ELSE 0 END "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 94
'..........................................................................................................................
'aggiunge campo di definizione dell'applicativo che ha creato l'articolo, usato per distinguere la tipologia di dato.
'..........................................................................................................................
function Aggiornamento__B2B__94(conn)
	Aggiornamento__B2B__94 = _
		" ALTER TABLE gtb_articoli ADD " + _
		"	art_applicativo_id INT NULL; " + _
		" UPDATE gtb_articoli SET art_applicativo_id = " & NEXTB2B & "; " + _
		" ALTER TABLE gtb_articoli ALTER COLUMN " + _
		"	art_applicativo_id INT NOT NULL ; " + _
		SQL_AddForeignKey(conn, "gtb_articoli", "art_applicativo_id", "tb_siti", "id_sito", true, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 95
'..........................................................................................................................
'aggiunge campo di definizione dello stato di pezzo unico
'..........................................................................................................................
function Aggiornamento__B2B__95(conn)
	Aggiornamento__B2B__95 = _
		" ALTER TABLE gtb_articoli ADD " + _
		"	art_unico BIT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 96
'..........................................................................................................................
'campo nome dello stato dell'ordine in multilingua
'..........................................................................................................................
function Aggiornamento__B2B__96(conn)
	Aggiornamento__B2B__96 = _
		" sp_rename 'gtb_stati_ordine.so_nome', 'so_nome_it', 'COLUMN'" + vbCrLf + _
		" ALTER TABLE gtb_stati_ordine ADD" + vbCrLf + _
		" so_nome_en NVARCHAR(200) NULL," + vbCrLf + _
		" so_nome_fr NVARCHAR(200) NULL," + vbCrLf + _
		" so_nome_es NVARCHAR(200) NULL," + vbCrLf + _
		" so_nome_de NVARCHAR(200) NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 97
'...........................................................................................
'converte i parametri del NextB2B
'...........................................................................................
Sub AggiornamentoSpeciale__B2B__97(DB, rs, version)
	CALL DB.Execute("SELECT * FROM aa_versione", version)
	if DB.last_update_executed then
		dim siti
		CALL ParametersImport(DB.objconn, 19)
	end if
End Sub
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 98
'..........................................................................................................................
'Sergio - 10/05/2009
'..........................................................................................................................
'tabelle per la gestione delle modalità di pagamento e delle eventuali spese di spedizione
'..........................................................................................................................
function Aggiornamento__B2B__98(conn)
	Aggiornamento__B2B__98 = _
		"CREATE TABLE dbo.gtb_modipagamento ( " + vbCrLf + _
        "    mosp_id int IDENTITY (1, 1) NOT NULL  , " + vbCrLf + _
		SQL_MultiLanguageField("	mosp_nome_<lingua> " + SQL_CharField(Conn, 255)) + ", " + _
		"    mosp_se_abilitato bit NOT NULL , " + vbCrLf + _
		"    mosp_se_spesespedizione bit NOT NULL , " + vbCrLf + _
		"    mosp_ammontare_spsp MONEY NULL ," + vbCrLf + _
		SQL_MultiLanguageField("	mosp_label_spsp_<lingua> " + SQL_CharField(Conn, 255)) + _
		");"  + vbCrLf + _
		"ALTER TABLE dbo.gtb_modipagamento WITH NOCHECK ADD  " + vbCrLf + _
        "    CONSTRAINT PK_gtb_modipagamento PRIMARY KEY ( mosp_id ) " + vbCrLf + _
        ";" + vbCrLf + _
        "INSERT INTO gtb_modipagamento ( mosp_nome_it,mosp_nome_en,mosp_se_abilitato,mosp_se_spesespedizione,mosp_ammontare_spsp,mosp_label_spsp_it,mosp_label_spsp_en ) " + vbCRLF +_
		"VALUES( 'Paypal','Paypal',1,0,0.0,'spedizione gratuita','free delivery');" + vbCRLF +_
		"INSERT INTO gtb_modipagamento ( mosp_nome_it,mosp_nome_en,mosp_se_abilitato,mosp_se_spesespedizione,mosp_ammontare_spsp,mosp_label_spsp_it,mosp_label_spsp_en ) " + vbCRLF +_
		"VALUES( 'Contrassegno','Shipping on delivery',1,1,10.0,'Sarà applicato un sovrapprezzo ...','...');" + vbCRLF +_
		"ALTER TABLE dbo.gtb_rivenditori ADD " + vbCrLf + _
        "    riv_modopagamento_id int NULL; "  + vbCrLf + _
		"ALTER TABLE dbo.gtb_ordini ADD " + vbCrLf + _
        "    ord_modopagamento_id int NOT NULL DEFAULT 1; "  + vbCrLf + _
		SQL_AddForeignKey(conn, "gtb_rivenditori", "riv_modopagamento_id", "gtb_modipagamento", "mosp_id", false, "") +_
		SQL_AddForeignKey(conn, "gtb_ordini", "ord_modopagamento_id", "gtb_modipagamento", "mosp_id", true, "") +_
		" UPDATE gtb_ordini SET ord_modopagamento_id = (SELECT mosp_id FROM gtb_modipagamento WHERE gtb_modipagamento.mosp_nome_it LIKE 'Paypal');" + vbCRLF
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 99
'..........................................................................................................................
'Sergio - 10/05/2009
'..........................................................................................................................
' campi per la gestione delle modalità di pagamento esterne come paypal
'..........................................................................................................................
function Aggiornamento__B2B__99(conn)
	Aggiornamento__B2B__99 = _
		" ALTER TABLE gtb_modipagamento ADD" + vbCrLf + _
		"    mosp_se_esterno bit NOT NULL DEFAULT 0 , " + vbCrLf + _
		"    mosp_url_servizio_esterno varchar(250) NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 100
'..........................................................................................................................
'Nicola - 11/05/2009
'..........................................................................................................................
'aggiunge campi per la gestione delle offerte speciali sull'articolo.
'modifica vista che filtra le offerte speciali
'..........................................................................................................................
function Aggiornamento__B2B__100(conn)
	Aggiornamento__B2B__100 = _
		" ALTER TABLE gtb_prezzi ADD " + _
		"	prz_offerta_dal SMALLDATETIME NULL, " + _
		"	prz_offerta_al SMALLDATETIME;  " + _
		 DropObject(conn, "gv_listino_offerte", "VIEW") + _
		"CREATE VIEW dbo.gv_listino_offerte AS " + vbCrLf + _
		"    SELECT * FROM gtb_articoli " + vbCrLF + _
		"        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLF + _
		"        INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLF + _
		"        INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLF + _
		"        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
		"    WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLF + _
		"          AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLF + _
		"          AND tip_visibile=1 " + vbCrLF + _
		"          AND tip_albero_visibile=1 " + vbCrLF + _
		"          AND ISNULL(listino_offerte, 0)=1 " + vbCrLF + _
		"          AND ISNULL(prz_visibile, 0)=1 " + vbCrLF + _
		"          AND ( " + vbCrLF + _
		"               ( GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1 ) " + vbCrLF + _
		"               OR " + vbCrLF + _
		"               ( listino_dataCreazione IS NULL AND " + vbCrLF + _
		"                 listino_dataScadenza IS NULL AND " + vbCrLF + _
		"                 prz_offerta_dal IS NOT NULL AND " + vbCrLF + _
		"                 GETDATE() BETWEEN prz_offerta_dal AND ISNULL(prz_offerta_al, GETDATE())+1 " + vbCrLF + _
		"               )  " + vbCrLF + _
		"              ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 101
'..........................................................................................................................
'Nicola - 12/05/2009
'..........................................................................................................................
'aggiunge campi per la descrizione "riassuntiva" negli articoli
'agginuge campi per la descrizione del prezzo dell'articolo
'aggiunge flag sui descrittori: descrittori per la ricerca, descrittori per il confronto
'..........................................................................................................................
function Aggiornamento__B2B__101(conn)
	Aggiornamento__B2B__101 = _
		" ALTER TABLE gtb_articoli ADD " + _
		SQL_MultiLanguageField("	art_descr_riassunto_<lingua> " + SQL_CharField(Conn, 0)) + ", " + _
		SQL_MultiLanguageField("	art_descr_prezzo_<lingua> " + SQL_CharField(Conn, 0)) + _
		";" + _
		" ALTER TABLE grel_art_valori ADD " + _
		SQL_MultiLanguageField("	rel_descr_prezzo_<lingua> " + SQL_CharField(Conn, 0)) + _
		";" + _
		" ALTER TABLE gtb_carattech ADD " + _
		"	ct_codice " + SQL_CharField(Conn, 255) + ", " + _
		"	ct_per_ricerca BIT NULL, " + _
		"	ct_per_confronto BIT NULL " + _
		";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 102
'..........................................................................................................................
'Nicola - 12/05/2009
'..........................................................................................................................
'aggiunge campo logo descrittori speciali
'..........................................................................................................................
function Aggiornamento__B2B__102(conn)
	Aggiornamento__B2B__102 = _
		" ALTER TABLE gtb_carattech ADD " + _
		"	ct_img " + SQL_CharField(Conn, 255) + _
		";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 103
'..........................................................................................................................
'Nicola - 13/05/2009
'..........................................................................................................................
'aggiunge raggruppamenti ai descrittori
'..........................................................................................................................
function Aggiornamento__B2B__103(conn)
	Aggiornamento__B2B__103 = _
		" ALTER TABLE gtb_carattech ADD " + _
		"	ct_raggruppamento_id INT NULL " + _
		" ; " + _
		" CREATE TABLE dbo.gtb_carattech_raggruppamenti ( " + vbCrLf + _
		"	ctr_id " + SQL_PrimaryKey(conn, "gtb_carattech_raggruppamenti") + ", " + _
		SQL_MultiLanguageField("	ctr_titolo_<lingua> " + SQL_CharField(Conn, 255)) + ", " + _
		"	ctr_ordine int NULL, " + vbCrLf + _
		"	ctr_codice " + SQL_CharField(Conn, 255) + ", " + _
		"	ctr_di_sistema int NULL" + vbCrLf + _
		" ) ; " + _
		SQL_AddForeignKey(conn, "gtb_carattech", "ct_raggruppamento_id", "gtb_carattech_raggruppamenti", "ctr_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 104
'..........................................................................................................................
'Sergio - 13/05/2009
'..........................................................................................................................
' campi per la gestione delle modalità di pagamento esterne come paypal
'..........................................................................................................................
function Aggiornamento__B2B__104(conn)
	Aggiornamento__B2B__104 = _
		" ALTER TABLE gtb_modipagamento ADD" + vbCrLf + _
		"    mosp_id_pagina_startup int NOT NULL DEFAULT 0 , " + vbCrLf + _
		"    mosp_url_logo_servizio_modo varchar(250) NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 105
'..........................................................................................................................
'Sergio - 17/06/2009
'..........................................................................................................................
' campo per mantenere le spese di spedizione dentro la shopping cart
'..........................................................................................................................
function Aggiornamento__B2B__105(conn)
	Aggiornamento__B2B__105 = _
		" ALTER TABLE dbo.gtb_shopping_cart ADD" + vbCrLf + _
		"    sc_spesespedizione money NOT NULL DEFAULT 0.0;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 106
'..........................................................................................................................
'Nicola - 06/07/2009
'..........................................................................................................................
' aggiunge funzioni sql per restituire il listino in vigore per gli articoli (prezzo minimo)
'..........................................................................................................................
function Aggiornamento__B2B__106(conn)
		if cIntero(DB_SQL_version(conn)) >= 9 then
			Aggiornamento__B2B__106 = _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iav, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid  " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) " + _
			" ; " + _
			" CREATE FUNCTION dbo.fn_listino_vendita_varianti( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT gv_articoli.*, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo) AS prezzo, " + vbCrLF + _
			" 		   COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore) AS iva, " + vbCrLF + _
			" 		   COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id) AS ivaId, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_scontoQ_id, cliente.prz_scontoQ_id, base.prz_scontoQ_id, rel_scontoQ_id) AS scontoQId, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0) AS listinoId, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) AS visibile, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_non_vendibile, cliente.prz_non_vendibile, base.prz_non_vendibile, 0) AS nonVendibile " + vbCrLF + _
			"	FROM gv_articoli " + vbCrLF + _
			" 		INNER JOIN gtb_prezzi base ON gv_articoli.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 		LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 		LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 						AND ISNULL(offerte.prz_visibile, 0) = 1 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 						AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 		LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id " + vbCrLF + _
			" 		LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 		LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"	WHERE ISNULL(art_disabilitato,0) = 0 " + vbCrLf + _
			"		AND ISNULL(rel_disabilitato,0) = 0 " + vbCrLF + _
			"		AND tip_visibile = 1 " + vbCrLf + _
			"		AND tip_albero_visibile = 1 " + vbCrLF + _
			"		AND COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) = 1 " + vbCrLF + _
			" ) "
		else
			Aggiornamento__B2B__106 = ""
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 107
'..........................................................................................................................
'Nicola - 17/06/2009
'..........................................................................................................................
' corregge campo flag su raggruppamenti
'..........................................................................................................................
function Aggiornamento__B2B__107(conn)
	Aggiornamento__B2B__107 = _
		" ALTER TABLE gtb_carattech_raggruppamenti " + _
		"	ALTER COLUMN ctr_di_sistema BIT NULL ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 108
'..........................................................................................................................
'Nicola - 18/06/2009
'..........................................................................................................................
'modifica listini: aggiunge campo per gestione variazione percentuale/euro di default direttamente su testata del listino
'..........................................................................................................................
function Aggiornamento__B2B__108(conn)
	Aggiornamento__B2B__108 = _
		" ALTER TABLE gtb_listini ADD " + _
		"	listino_default_var_euro REAL NULL, " + _
		"	listino_default_var_sconto REAL NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 109
'...........................................................................................
'Giacomo - 20/08/2009
'...........................................................................................
'aggiunge colonna codice alla tabella descrittore righe ordine
'...........................................................................................
function Aggiornamento__B2B__109(conn)
	Aggiornamento__B2B__109 = _
		" ALTER TABLE gtb_dettagli_ord_des ADD dod_codice "+SQL_CharField(Conn, 255)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 110
'...........................................................................................
'Giacomo - 25/08/2009
'...........................................................................................
'modifica il tipo dei campi mosp_label_spsp_it...della tabella gtb_modipagamento
'...........................................................................................
function Aggiornamento__B2B__110(conn)
	Aggiornamento__B2B__110 = _
		" ALTER TABLE gtb_modipagamento ALTER COLUMN mosp_label_spsp_it "+SQL_CharField(Conn, 0)  + vbCrLF + _
		" ALTER TABLE gtb_modipagamento ALTER COLUMN mosp_label_spsp_en "+SQL_CharField(Conn, 0)  + vbCrLF + _
		" ALTER TABLE gtb_modipagamento ALTER COLUMN mosp_label_spsp_de "+SQL_CharField(Conn, 0)  + vbCrLF + _
		" ALTER TABLE gtb_modipagamento ALTER COLUMN mosp_label_spsp_fr "+SQL_CharField(Conn, 0)  + vbCrLF + _
		" ALTER TABLE gtb_modipagamento ALTER COLUMN mosp_label_spsp_es "+SQL_CharField(Conn, 0)
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 111
'...........................................................................................
'Giacomo - 31/08/2009
'...........................................................................................
'aggiunge il campo mosp_default alla tabella gtb_modipagamento
'...........................................................................................
function Aggiornamento__B2B__111(conn)
	Aggiornamento__B2B__111 = _
		" ALTER TABLE gtb_modipagamento ADD mosp_default bit NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 112
'...........................................................................................
'Giacomo - 02/09/2009
'...........................................................................................
'crea nuova tabella grel_art_pag che collega gli articoli alle FAQ
'...........................................................................................
function Aggiornamento__B2B__112(conn)
	Aggiornamento__B2B__112 = _
		" CREATE TABLE dbo.grel_art_faq ( " + vbCrLf + _
		"	raf_id " + SQL_PrimaryKey(conn, "grel_art_faq") + ", " + _
		"	raf_art_id INT NULL, " + vbCrLf + _
		"	raf_faq_id INT NULL" + vbCrLf + _
		" ) ; " + _
		SQL_AddForeignKey(conn, "grel_art_faq", "raf_art_id", "gtb_articoli", "art_id", false, "") + _
		SQL_AddForeignKey(conn, "grel_art_faq", "raf_faq_id", "tb_FAQ", "faq_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 113
'..........................................................................................................................
'Nicola - 09/09/2009
'..........................................................................................................................
'modifica vista che filtra le offerte speciali
'..........................................................................................................................
function Aggiornamento__B2B__113(conn)
	Aggiornamento__B2B__113 = _
		 DropObject(conn, "gv_listino_offerte", "VIEW") + _
		"CREATE VIEW dbo.gv_listino_offerte AS " + vbCrLf + _
		"    SELECT * FROM gtb_articoli " + vbCrLF + _
		"        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLF + _
		"        INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLF + _
		"        INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLF + _
		"        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
		"    WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLF + _
		"          AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLF + _
		"          AND tip_visibile=1 " + vbCrLF + _
		"          AND tip_albero_visibile=1 " + vbCrLF + _
		"          AND ISNULL(listino_offerte, 0)=1 " + vbCrLF + _
		"          AND ISNULL(prz_visibile, 0)=1 " + vbCrLF + _
		"          AND ( " + vbCrLF + _
		"               ( GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1 ) " + vbCrLF + _
		"               OR " + vbCrLF + _
		"               ( listino_dataCreazione IS NULL AND " + vbCrLF + _
		"                 listino_dataScadenza IS NULL AND " + vbCrLF + _
		"                 GETDATE() BETWEEN ISNULL(prz_offerta_dal, GETDATE()-1) AND ISNULL(prz_offerta_al, GETDATE())+1 " + vbCrLF + _
		"               )  " + vbCrLF + _
		"              ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 114
'..........................................................................................................................
'Giacomo / Andrea - 25/09/2009
'..........................................................................................................................
'creazione tabella spese di spedizione articolo e collegamento con gtb_articoli
'..........................................................................................................................
function Aggiornamento__B2B__114(conn)
	Aggiornamento__B2B__114 = _
		"CREATE TABLE gtb_spese_spedizione_articolo ( " + vbCrLF + _ 
        "       spa_id int PRIMARY KEY IDENTITY (1, 1) NOT NULL , " + vbCrLF + _
		        SQL_MultiLanguageField("	spa_nome_<lingua> " + SQL_CharField(Conn, 255)) + ", " + _        
				SQL_MultiLanguageField("	spa_condizioni_<lingua> " + SQL_CharField(Conn, 0)) + ", " + _        
        "   	spa_annullamento_qta int NULL , " + vbCrLF + _
        "   	spa_annullamento_importo money NULL , " + vbCrLF + _
        "   	spa_importo_spese money NULL) " + vbCrLF + _
        ";" + vbCrLF + _
		" INSERT INTO gtb_spese_spedizione_articolo (spa_nome_it) VALUES ('Default') ; " + vbCrLF+ _	
		" ALTER TABLE gtb_articoli ADD art_spedizione_id int NULL ; " + vbCrLf + _
		" UPDATE gtb_articoli SET art_spedizione_id = 1; " + vbCrlf + _
		" ALTER TABLE gtb_articoli ALTER COLUMN art_spedizione_id int NOT NULL; " + _
		SQL_AddForeignKey(conn, "gtb_articoli", "art_spedizione_id", "gtb_spese_spedizione_articolo", "spa_id", true, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 115
'..........................................................................................................................
'Giacomo - 28/09/2009
'..........................................................................................................................
'ricreazione delle viste e delle funzioni dopo l'aggiunta della tabella gtb_spese_spedizione_articolo
'..........................................................................................................................
function Aggiornamento__B2B__115(conn)
	Aggiornamento__B2B__115 = _
		"DROP VIEW dbo.gv_articoli ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_articoli AS" + vbCrLf + _
		"	SELECT * FROM gtb_articoli" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLf + _
		"		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id " + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";" + vbCrLf + _
		"DROP VIEW dbo.gv_carichi ;" + vbCrLf + _
		"CREATE VIEW dbo.gv_carichi AS " + vbCrLf + _
		"	SELECT * FROM grel_carichi_var " + vbCrLf + _
		"		INNER JOIN grel_art_valori ON grel_carichi_var.rcv_art_var_id = grel_art_valori.rel_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";" + vbCrLf + _
		DropObject(conn, "gv_CartDetail","VIEW") + _
		" CREATE VIEW dbo.gv_CartDetail AS " + vbCrLF + _
		"	SELECT *, (SELECT COUNT(dd_ind_id) FROM gtb_dett_Cart_dest WHERE dd_dett_id = gtb_dett_cart.dett_id) AS N_DEST, " + vbCrLF + _
		"			  (SELECT COUNT(dp_ut_id) FROM gtb_dett_Cart_proposte WHERE dp_Dett_id=gtb_dett_Cart.dett_id) AS N_UT " + vbCrLF + _
		"		FROM gtb_dett_cart LEFT JOIN gtb_iva ON gtb_dett_cart.dett_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"		LEFT JOIN grel_art_valori ON gtb_dett_cart.dett_art_var_id = grel_art_valori.rel_id " + vbCrLF + _
		"		LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLF + _
		"		LEFT JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		"	WHERE (gtb_dett_cart.dett_art_var_id IS NULL) OR " + vbCrLf + _
		"		( ISNULL(gtb_articoli.art_disabilitato, 0)=0 AND " + vbCrLf + _
		"		  ISNULL(grel_art_valori.rel_disabilitato,0)=0 AND " + vbCrLf + _
		"		  ISNULL(gtb_tipologie.tip_albero_visibile, 0) = 1 AND " + vbCrLf + _
		"		  ISNULL(gtb_tipologie.tip_visibile, 0)= 1 ) " + vbCrLf + _
		";" + vbCrLf + _
		DropObject(conn, "gv_dettagli_ord","VIEW") + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + vbCrLF + _
		"	SELECT * FROM gtb_dettagli_ord " + vbCrLF + _
		"		LEFT JOIN grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id " + vbCrLF + _
		"		LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLF + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";" + vbCrLf + _
		" DROP VIEW dbo.gv_listini ;" + vbCrLf + _
		" CREATE VIEW dbo.gv_listini AS" + vbCrLf + _
		"	SELECT * FROM grel_art_valori" + vbCrLf + _
		"		INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id" + vbCrLf + _
		"		INNER JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id" + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";" + vbCrLf + _
		DropObject(conn, "gv_listino_offerte", "VIEW") + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + vbCrLf + _
		"    SELECT * FROM gtb_articoli " + vbCrLF + _
		"        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLF + _
		"        INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLF + _
		"        INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLF + _
		"        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
		"		 INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		"    WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLF + _
		"          AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLF + _
		"          AND tip_visibile=1 " + vbCrLF + _
		"          AND tip_albero_visibile=1 " + vbCrLF + _
		"          AND ISNULL(listino_offerte, 0)=1 " + vbCrLF + _
		"          AND ISNULL(prz_visibile, 0)=1 " + vbCrLF + _
		"          AND ( " + vbCrLF + _
		"               ( GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1 ) " + vbCrLF + _
		"               OR " + vbCrLF + _
		"               ( listino_dataCreazione IS NULL AND " + vbCrLF + _
		"                 listino_dataScadenza IS NULL AND " + vbCrLF + _
		"                 GETDATE() BETWEEN ISNULL(prz_offerta_dal, GETDATE()-1) AND ISNULL(prz_offerta_al, GETDATE())+1 " + vbCrLF + _
		"               )  " + vbCrLF + _
		"              ) " + vbCrLF + _
		";" + vbCrLf + _
		" DROP VIEW dbo.gv_listino_vendita; " + vbCrLf + _
		" CREATE VIEW dbo.gv_listino_vendita AS " + vbCrLf + _
		"		SELECT * FROM gtb_articoli " + vbCrLf + _
		" 		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLf + _
		"			INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLf + _
		"			INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLf + _
		" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLf + _
		" 		INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLf + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		"			WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLf + _
		"				  AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLf + _
		"				  AND tip_visibile=1 " + vbCrLf + _
		"				  AND tip_albero_visibile=1 " + vbCrLf + _
		"				  AND prz_visibile=1 " + vbCrLf + _
		"				  AND ( ( ISNULL(listino_offerte, 0)=1 " + vbCrLf + _
		"				  		  AND ISNULL(prz_visibile, 0)=1 " + vbCrLf + _
		"						  AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) " + vbCrLf + _
		"						) " + vbCrLf + _
		"						OR (ISNULL(listino_offerte, 0)=0 " + vbCrLf + _
		"				  AND prz_variante_id NOT IN ( " + vbCrLf + _
		"						SELECT prz_variante_id FROM gtb_listini INNER JOIN gtb_prezzi ON gtb_listini.listino_id=gtb_prezzi.prz_listino_id " + vbCrLf + _
		"						WHERE ISNULL(listino_offerte, 0)=1 AND ISNULL(prz_visibile, 0)=1 " + vbCrLf + _
		"						AND (GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GetDate())+1) " + vbCrLf + _
		"											 ) " + vbCrLf + _
		"						) " + vbCrLf + _
		"					  ) "
		 
		
		if cIntero(DB_SQL_version(conn)) >= 9 then
			Aggiornamento__B2B__115 = Aggiornamento__B2B__115 + ";" + vbCrLf + _
			" DROP FUNCTION dbo.fn_listino_vendita_articoli; " + vbCrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 AND " + vbCrLF + _
			" 				 								 GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 AND " + vbCrLF + _
			" 				 								 GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iav, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid  " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) " + _
			" ; "
		end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 116
'...........................................................................................
'Giacomo - 30/09/2009
'...........................................................................................
'aggiunge colonna sp_annullamento_importo alla tabella gtb_spese_spedizione
'...........................................................................................
function Aggiornamento__B2B__116(conn)
	Aggiornamento__B2B__116 = _
		" ALTER TABLE gtb_spese_spedizione ADD " + _
		"	sp_annullamento_importo money NULL "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 117
'...........................................................................................
'Giacomo - 06/10/2009
'...........................................................................................
'aggiunge colonne mosp_istruzioni_ alla tabella gtb_modipagamento
'...........................................................................................
function Aggiornamento__B2B__117(conn)
	Aggiornamento__B2B__117 = _
		" ALTER TABLE gtb_modipagamento ADD " + _
			SQL_MultiLanguageField("	mosp_istruzioni_<lingua> " + SQL_CharField(Conn, 0))
end function
'********************************************************************************************
		
		
'*******************************************************************************************
'AGGIORNAMENTO B2B 118
'..........................................................................................................................
'Matteo - 08/10/2009
'..........................................................................................................................
' campo per mantenere le spese di incasso, quelle fisse e le altre spese dentro la shopping cart
'..........................................................................................................................
function Aggiornamento__B2B__118(conn)
	Aggiornamento__B2B__118 = _
		" ALTER TABLE dbo.gtb_shopping_cart ADD" + vbCrLf + _
		"	sc_speseincasso money NULL, " + _
		"	sc_spesefisse money NULL, " + _
		"   sc_spesealtre money NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 119
'..........................................................................................................................
'Giacomo - 09/10/2009
'..........................................................................................................................
' campo per mantenere le spese di incasso, quelle fisse e le altre spese dentro gli ordini
'..........................................................................................................................
function Aggiornamento__B2B__119(conn)
	Aggiornamento__B2B__119 = _
		" ALTER TABLE dbo.gtb_ordini ADD" + vbCrLf + _
		"	ord_spesespedizione money NULL, " + _
		"	ord_speseincasso money NULL, " + _
		"	ord_spesefisse money NULL, " + _
		"   ord_spesealtre money NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 120
'..........................................................................................................................
'Giacomo - 30/11/2009
'..........................................................................................................................
' 
'..........................................................................................................................
function Aggiornamento__B2B__120(conn)
	Aggiornamento__B2B__120 = _
					"	ALTER TABLE gtb_articoli ADD " + _
					SQL_MultiLanguageField(" art_url_<lingua> " + SQL_CharField(Conn, 500)) + ";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 121
'..........................................................................................................................
'Matteo - 30/11/2009
'..........................................................................................................................
' campo per gestire il modo di pagamento personalizzato e l'inserimento con stato ordine a seconda del modo di pagamento
'..........................................................................................................................
function Aggiornamento__B2B__121(conn)
	Aggiornamento__B2B__121 = _
		" ALTER TABLE dbo.gtb_modipagamento ADD" + vbCrLf + _
		"	mosp_personalizzato bit NULL, " + _
		"   mosp_stato_ordine_id int NULL; " + _
		" UPDATE gtb_modipagamento " + _
		"    SET mosp_stato_ordine_id = (SELECT so_id FROM gtb_stati_ordine WHERE so_internet = 1), " + vbCrlf + _
		"    	 mosp_personalizzato = 0; " + vbCrlf + _
		" ALTER TABLE dbo.gtb_modipagamento " + _
		"  ALTER COLUMN mosp_stato_ordine_id int NOT NULL ;" + _
		SQL_AddForeignKeyExtended(conn, "gtb_modipagamento", "mosp_stato_ordine_id", "gtb_stati_ordine", "so_id", true, false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 122
'...........................................................................................
'Nicola - 03/12/2009
'...........................................................................................
'aggiunge colonne con codice di inserimento per shopping cart ed ordine
'...........................................................................................
function Aggiornamento__B2B__122(conn)
	Aggiornamento__B2B__122 = _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_codice_inserimento nvarchar(50) NULL; " + _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_codice_inserimento nvarchar(50) NULL; "
end function
'********************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 123
'...........................................................................................
'Nicola - 24/12/2009
'...........................................................................................
'
'...........................................................................................
function Aggiornamento__B2B__123(conn)
	Aggiornamento__B2B__123 = _
		" CREATE FUNCTION [dbo].[fn_articolo_lista_codici_varianti]  " + _
		"  (   " + _
		"		@ArtId int " + _
		"  )   " + _
		" RETURNS nvarchar(3000)  " + _
		" AS   " + _
		" BEGIN   " + _
		"		DECLARE @cod_int nvarchar(255), @cod_pro nvarchar(255), @cod_alt nvarchar(255) " + _
		"		DECLARE @codici nvarchar(3000) " + _
		"	  SELECT @codici = ''  " + _
		
		"		DECLARE RS CURSOR FOR " + _
		"			SELECT rel_cod_int, rel_cod_alt, rel_cod_pro FROM grel_art_valori WHERE rel_art_id= @ArtId " + _
		"		OPEN RS " + _
					
		"		FETCH NEXT FROM RS INTO @cod_int, @cod_pro, @cod_alt " + _
			  
		"	  WHILE (@@fetch_status <> -1) " + _
		"			BEGIN " + _
		"				  IF (@@fetch_status <> -2) " + _
		"				  BEGIN " + _
		"						IF(IsNull(@cod_int, '') <> '') " + _
		"							 SET @codici = @codici + ' ' + @cod_int " + _
		"						IF(IsNull(@cod_pro, '') <> '') " + _
		"							 SET @codici = @codici + ' ' + @cod_pro " + _
		"						IF(IsNull(@cod_alt, '') <> '') " + _
		"							 SET @codici = @codici + ' ' + @cod_alt " + _
		"				  END " + _
		"				  FETCH NEXT FROM RS INTO @cod_int, @cod_pro, @cod_alt " + _
		"			END " + _
			  
		"		CLOSE RS " + _
		"	  DEALLOCATE RS " + _
		"		RETURN  @codici   " + _
		" END "
end function
'********************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 124
'...........................................................................................
'aggiunge campo ordine a gtb_articoli
'Giacomo 11/05/2010
'...........................................................................................
function Aggiornamento__B2B__124(conn)
	Aggiornamento__B2B__124 = _
		" ALTER TABLE gtb_articoli ADD " + _
		"		art_ordine INT NULL; " + _
		" ALTER TABLE gItb_articoli ADD " + _
		"		Iart_ordine INT NULL; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 125
'...........................................................................................
'	Nicola, 11/05/2010
'...........................................................................................
function Aggiornamento__B2B__125(conn)
	Aggiornamento__B2B__125 = _
		" ALTER TABLE gtb_dettagli_ord_des ADD " + _
		"	dod_qta_in_detrazione BIT NULL, " + _
		" 	dod_percentuale_detrazione INT NULL ; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 126
'...........................................................................................
'Giacomo 11/06/2010
'aggiunge tabella per tipizzazione delle foto
'...........................................................................................
function Aggiornamento__B2B__126(conn)
	Aggiornamento__B2B__126 = _
		" ALTER TABLE gtb_art_foto ADD " + _
		"	fo_tipo_id INT NULL; " + _
		" CREATE TABLE " & SQL_Dbo(conn) & "gtb_foto_tipo ( " & _
		"	ft_id " + SQL_PrimaryKey(conn, "gtb_foto_tipo") + ", " + _
		"	ft_nome " + SQL_CharField(Conn, 255) + " NULL, "+ _
		"	ft_codice " + SQL_CharField(Conn, 255) + " NULL " + _
		" ) ; " + _
		SQL_AddForeignKey(conn, "gtb_art_foto", "fo_tipo_id", "gtb_foto_tipo", "ft_id", true, "") + _
		" INSERT INTO gtb_foto_tipo(ft_nome,ft_codice) VALUES ('immagini', 'img') ; " + _
		" UPDATE gtb_art_foto SET fo_tipo_id = 1 ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 127
'...........................................................................................
'	Matteo, 23/06/2010
'...........................................................................................
'   aggiunge campi data e numero fattura ordine
'...........................................................................................
function Aggiornamento__B2B__127(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento__B2B__127 = _
				" ALTER TABLE gtb_ordini ADD" + vbCrLf + _
				"	ord_fattura_numero INT NULL, " + _
				"	ord_fattura_data SMALLDATETIME NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 128
'...........................................................................................
'	Matteo, 25/06/2010
'...........................................................................................
'   aggiunge campi path file voucher e fattura ordine
'...........................................................................................
function Aggiornamento__B2B__128(conn)
	Select case DB_Type(conn)
		case DB_SQL
			Aggiornamento__B2B__128 = _
				" ALTER TABLE gtb_ordini ADD" + vbCrLf + _
				"	ord_fattura_file NVARCHAR(250) NULL, " + _
				"	ord_conferma_ordine_data SMALLDATETIME NULL, " + _
				"	ord_conferma_ordine_file NVARCHAR(250) NULL; "
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 129
'...........................................................................................
'	Andrea, 19/07/2010
'...........................................................................................
'   aggiunge campi totale ordine e totale iva alla tabella ordini
'...........................................................................................
function Aggiornamento__B2B__129(conn)
	Aggiornamento__B2B__129 = _
		" ALTER TABLE gtb_ordini ADD" + vbCrLf + _
		"	ord_totale money NULL, " + _
		"	ord_totale_iva money NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 130
'...........................................................................................
'	Andrea, 19/07/2010
'...........................................................................................
'   calcola i valori dei campi totale e totale iva per la tabella ordini
'...........................................................................................
function Aggiornamento__B2B__130(conn)
	Aggiornamento__B2B__130 = _
		" UPDATE gtb_ordini SET ord_totale=(SELECT SUM(det_prezzo_unitario*det_qta) " +_
		" FROM gtb_dettagli_ord WHERE det_ord_id=gtb_ordini.ord_id);" +_
		" UPDATE gtb_ordini SET ord_totale_iva=(SELECT SUM(det_prezzo_unitario*det_qta*det_iva/100) " +_
		" FROM gtb_dettagli_ord WHERE det_ord_id=gtb_ordini.ord_id);"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 131
'...........................................................................................
'	Andrea, 19/07/2010
'...........................................................................................
'   aggiunge campi totale ordine e totale iva alla tabella shopping cart
'...........................................................................................
function Aggiornamento__B2B__131(conn)
	Aggiornamento__B2B__131 = _
		" ALTER TABLE gtb_shopping_cart ADD" + vbCrLf + _
		"	sc_totale money NULL, " + _
		"	sc_totale_iva money NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 132
'...........................................................................................
'	Andrea, 19/07/2010
'...........................................................................................
'   calcola i valori dei campi totale e totale iva per la tabella shopping cart
'...........................................................................................
function Aggiornamento__B2B__132(conn)
	Aggiornamento__B2B__132 = _
		" UPDATE gtb_shopping_cart SET sc_totale=(SELECT SUM(dett_prezzo_unitario*dett_qta) " +_
		" FROM gtb_dett_cart WHERE dett_cart_id=gtb_shopping_cart.sc_id);" +_
		" UPDATE gtb_shopping_cart SET sc_totale_iva=(SELECT SUM(dett_prezzo_unitario*dett_qta*iva_valore/100) " +_
		" FROM gtb_dett_cart JOIN gtb_iva ON dett_iva_id=iva_id WHERE dett_cart_id=gtb_shopping_cart.sc_id);"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 133
'...........................................................................................
'	Matteo, 19/07/2010
'...........................................................................................
'   aggiunge agli articoli la chiave esterna verso i tipi di riga d'ordine
'...........................................................................................
function Aggiornamento__B2B__133(conn)
	Aggiornamento__B2B__133 = _
		" ALTER TABLE gtb_articoli ADD " + _
		"	art_dettagli_ord_tipo_id INT NULL; " + _
		SQL_AddForeignKey(conn, "gtb_articoli", "art_dettagli_ord_tipo_id", "gtb_dettagli_ord_tipo", "dot_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 134
'...........................................................................................
'	Andrea, 19/07/2010
'...........................................................................................
'   aggiungo il riferimento ai listini per il dettaglio ordine
'...........................................................................................
function Aggiornamento__B2B__134(conn)
	Aggiornamento__B2B__134 = _	
		" ALTER TABLE gtb_dettagli_ord_des ADD " + _
		"	dod_listino_id INT NULL; " + _	
		SQL_AddForeignKey(conn, "gtb_dettagli_ord_des", "dod_listino_id", "gtb_listini", "listino_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 135
'...........................................................................................
'	Matteo, 19/07/2010
'...........................................................................................
'   crea la tabella dei tipi di fatturazioni
'...........................................................................................
function Aggiornamento__B2B__135(conn)
	Aggiornamento__B2B__135 = _
		" CREATE TABLE " & SQL_Dbo(conn) & "gtb_fatturazioni ( " + _
		"   fatt_id " + SQL_PrimaryKey(conn, "gtb_fatturazioni") + ", " + _
		"   fatt_codice NVARCHAR (255) NOT NULL , " + _
		"   fatt_numero_corrente INT NOT NULL , " + _
		"   fatt_data_corrente SMALLDATETIME NOT NULL , " + _
		"   fatt_serie NVARCHAR (50) NOT NULL) ;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 136
'...........................................................................................
'	Matteo, 19/07/2010
'...........................................................................................
'   aggiunge agli ordini la chiave esterna verso i tipi di fatturazione
'...........................................................................................
function Aggiornamento__B2B__136(conn)
	Aggiornamento__B2B__136 = _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_fatturazione_id INT NULL; " + _
		SQL_AddForeignKey(conn, "gtb_ordini", "ord_fatturazione_id", "gtb_fatturazioni", "fatt_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 137
'...........................................................................................
'	Matteo, 20/07/2010
'...........................................................................................
'   aggiunge campo serie fattura ordini
'...........................................................................................
function Aggiornamento__B2B__137(conn)
	Aggiornamento__B2B__137 = _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_fattura_serie NVARCHAR (50) NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 138
'...........................................................................................
'	Matteo, 20/07/2010
'...........................................................................................
'   crea la SP che imposta numero e data fattura nell'ordine passato in input
'	ed aggiorna la tabella delle fatturazioni con numero e data correnti
'...........................................................................................
function Aggiornamento__B2B__138(conn)
	Aggiornamento__B2B__138 = _
		DropObject(conn, "sp_UpdateFattura", "PROCEDURE") + _
		" CREATE PROCEDURE " & SQL_Dbo(conn) & "sp_UpdateFattura ( " + vbCrLf + _
		"	  @ord_id INT, " + vbCrLf + _
		"	  @fatt_id INT, " + vbCrLf + _
		"	  @ord_stato_confermato_id INT, " + vbCrLf + _
		"	  @ord_fattura_data SMALLDATETIME " + vbCrLf + _
		" ) " + vbCrLf + _
		vbCrLf + _
		" AS " + vbCrLf + _
		"	  DECLARE @codice NVARCHAR(255) " + vbCrLf + _
		"	  DECLARE @numero_corrente INT " + vbCrLf + _
		"	  DECLARE @serie NVARCHAR(50) " + vbCrLf + _
		"	  DECLARE @data_corrente SMALLDATETIME " + vbCrLf + _
		"	  DECLARE @data_fattura SMALLDATETIME " + vbCrLf + _
		vbCrLf + _
		"	  -- se la data in input non è valorizzata si prende la data odierna " + vbCrLf + _
		"	  SET @data_fattura = ISNULL(@ord_fattura_data, GETDATE()) " + vbCrLf + _
		vbCrLf + _
		"	  -- recupera i dati per la fatturazione appropriata " + vbCrLf + _
		"	  SELECT @codice = fatt_codice, " + vbCrLf + _
		"		     @numero_corrente = fatt_numero_corrente, " + vbCrLf + _
		"		     @data_corrente = fatt_data_corrente, " + vbCrLf + _
		"		     @serie = fatt_serie " + vbCrLf + _
		"	  FROM gtb_fatturazioni " + vbCrLf + _
		"	  WHERE fatt_id = @fatt_id " + vbCrLf + _
		vbCrLf + _
		"	  -- se non trova il tipo di fatturazione per l'ordine in input ritorna errore 0 " + vbCrLf + _
		"	  IF @@ROWCOUNT = 0 " + vbCrLf + _
		"		  RETURN 0 " + vbCrLf + _
		vbCrLf + _
		"	  ELSE " + vbCrLf + _
		"	  BEGIN " + vbCrLf + _
		"		  -- se l'anno corrente è diverso da quello attuale azzera il numero fattura corrente " + vbCrLf + _
		"		  IF YEAR(@data_corrente) <> YEAR(@data_fattura) " + vbCrLf + _
		"			  SET @numero_corrente = 0 " + vbCrLf + _
		"	  END " + vbCrLf + _
		vbCrLf + _
		"	  -- incrementa il numero fattura " + vbCrLf + _
		"	  SET @numero_corrente = @numero_corrente + 1 " + vbCrLf + _
		vbCrLf + _
		"	  -- setta la data fattura " + vbCrLf + _
		"	  SET @data_corrente = @data_fattura " + vbCrLf + _
		vbCrLf + _
		"	  BEGIN TRAN " + vbCrLf + _
		"		  -- aggiorna numero e data fattura per l'ordine in input " + vbCrLf + _
		"		  -- con stato confermato in input e numero fattura non valorizzato " + vbCrLf + _
		"		  UPDATE gtb_ordini " + vbCrLf + _
		"		  SET ord_fattura_numero = @numero_corrente " + vbCrLf + _
		"		    , ord_fattura_data = @data_corrente " + vbCrLf + _
		"		    , ord_fattura_serie = @serie " + vbCrLf + _
		"		    , ord_fatturazione_id = @fatt_id " + vbCrLf + _
		"		  WHERE ord_id = @ord_id " + vbCrLf + _
		"		  AND ord_stato_id = @ord_stato_confermato_id " + vbCrLf + _
		"		  AND ord_fattura_numero IS NULL " + vbCrLf + _
		vbCrLf + _
		"		  -- se non ha aggiornato l'ordine ritorna errore -1 " + vbCrLf + _
		"		  IF @@ROWCOUNT <> 1 " + vbCrLf + _
		"		  BEGIN" + vbCrLf + _
		"			  ROLLBACK TRAN " + vbCrLf + _
		"			  RETURN -1 " + vbCrLf + _
		"		  END " + vbCrLf + _
		vbCrLf + _
		"		  ELSE " + vbCrLf + _
		"		  BEGIN " + vbCrLf + _
		"			  -- aggiorna la tabella delle fatturazioni con numero e data corrente " + vbCrLf + _
		"			  UPDATE gtb_fatturazioni " + vbCrLf + _
		"			  SET fatt_numero_corrente = @numero_corrente " + vbCrLf + _
		"			    , fatt_data_corrente = @data_corrente " + vbCrLf + _
		"			  WHERE fatt_codice = @codice " + vbCrLf + _
		vbCrLf + _
		"			  -- se non ha aggiornato la tabella delle fatturazioni ritorna errore -2 " + vbCrLf + _
		"			  IF @@ROWCOUNT <> 1 " + vbCrLf + _
		"			  BEGIN " + vbCrLf + _
		"				  ROLLBACK TRAN " + vbCrLf + _
		"				  RETURN -2 " + vbCrLf + _
		"			  END " + vbCrLf + _
		vbCrLf + _
		"			  ELSE " + vbCrLf + _
		"			  BEGIN " + vbCrLf + _
		"				  -- transazione OK (ritorna 1) " + vbCrLf + _
		"				  COMMIT TRAN " + vbCrLf + _
		"				  RETURN 1 " + vbCrLf + _
		"			  END " + vbCrLf + _
		"		  END "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 139
'...........................................................................................
'	Andrea, 21/07/2010
'...........................................................................................
'   aggiunge campi per la gestione delle spese e dell'iva
'...........................................................................................
function Aggiornamento__B2B__139(conn)
	
	Aggiornamento__B2B__139 = _
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_spesespedizione money NULL, " + _
		"	dett_speseincasso money null, " + _
		"	dett_spesefisse money null, " + _
		"	dett_spesealtre money null, " + _
		"	dett_spesespedizione_iva_id int null, " + _
		"	dett_speseincasso_iva_id int null, " + _
		"	dett_spesefisse_iva_id int null, " + _
		"	dett_spesealtre_iva_id int null, " + _
		"	dett_totale money null, " + _
		"	dett_totale_iva money null, " + _
		"	dett_totale_spese money null, " + _
		"	dett_totale_spese_iva money null; " + _
		SQL_AddForeignKey(conn, "gtb_dett_cart", "dett_spesespedizione_iva_id", "gtb_iva", "iva_id", false, "spesespedizione") + _
		SQL_AddForeignKey(conn, "gtb_dett_cart", "dett_speseincasso_iva_id", "gtb_iva", "iva_id", false, "speseincasso") + _
		SQL_AddForeignKey(conn, "gtb_dett_cart", "dett_spesefisse_iva_id", "gtb_iva", "iva_id", false, "spesefisse") + _
		SQL_AddForeignKey(conn, "gtb_dett_cart", "dett_spesealtre_iva_id", "gtb_iva", "iva_id", false, "spesealtre") + _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_spesespedizione_iva_id int null, " + _
		"	sc_speseincasso_iva_id int null, " + _
		"	sc_spesefisse_iva_id int null, " + _
		"	sc_spesealtre_iva_id int null, " + _		
		"	sc_dett_totale_spese money null, " + _
		"	sc_dett_totale_spese_iva money null; " + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_spesespedizione_iva_id", "gtb_iva", "iva_id", false, "spesespedizione") + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_speseincasso_iva_id", "gtb_iva", "iva_id", false, "speseincasso") + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_spesefisse_iva_id", "gtb_iva", "iva_id", false, "spesefisse") + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_spesealtre_iva_id", "gtb_iva", "iva_id", false, "spesealtre") + _
		" ALTER TABLE gtb_dettagli_ord ADD " + _
		"	det_spesespedizione money NULL, " + _
		"	det_speseincasso money null, " + _
		"	det_spesefisse money null, " + _
		"	det_spesealtre money null, " + _
		"	det_spesespedizione_iva real null, " + _
		"	det_speseincasso_iva real null, " + _
		"	det_spesefisse_iva real null, " + _
		"	det_spesealtre_iva real null, " + _
		"	det_totale money null, " + _
		"	det_totale_iva money null, " + _
		"	det_totale_spese money null, " + _
		"	det_totale_spese_iva money null; " + _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_spesespedizione_iva real null, " + _
		"	ord_speseincasso_iva real null, " + _
		"	ord_spesefisse_iva real null, " + _
		"	ord_spesealtre_iva real null, " + _
		"	ord_det_totale_spese money null, " + _
		"	ord_det_totale_spese_iva money null, " + _
		"	ord_totale_spese money null, " + _
		"	ord_totale_spese_iva money null; " 
		
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 140
'...........................................................................................
'	Matteo, 23/07/2010
'...........................................................................................
'   aggiunge campi per la gestione delle spese e dell'iva (vedi aggiornamento 139)
'...........................................................................................
function Aggiornamento__B2B__140(conn)
	Aggiornamento__B2B__140 = _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_totale_spese money null, " + _
		"	sc_totale_spese_iva money null; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 141
'...........................................................................................
'	Giacomo, 24/08/2010
'...........................................................................................
'   crea e richiama per ogni ordine la stored procedure che aggiorna i campi aggiunti con l'agg. 139 e 140
'...........................................................................................
function Aggiornamento__B2B__141(conn)
	Aggiornamento__B2B__141 = _
		"	CREATE PROCEDURE [dbo].[gsp_totale_ordini] " + vbCrLf + _
		"		@ord_id INT " + vbCrLf + _
		"	AS " + vbCrLf + _
		"	BEGIN " + vbCrLf + _
		"		UPDATE gtb_dettagli_ord  " + vbCrLf + _
		"		SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0),2) ,  " + vbCrLf + _
		"			det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0)*ISNULL(det_iva,0)/100,2) ,  " + vbCrLf + _
		"			det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +  " + vbCrLf + _
		"									ISNULL(det_speseincasso,0) + " + vbCrLf + _
		"									ISNULL(det_spesefisse,0)+ " + vbCrLf + _
		"									ISNULL(det_spesealtre,0),2) ,  " + vbCrLf + _
		"			det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 + " + vbCrLf + _
		"										ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 + " + vbCrLf + _
		"										ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 + " + vbCrLf + _
		"										ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)  " + vbCrLf + _
		"		WHERE det_ord_id=@ord_id " + vbCrLf + _
		"		UPDATE gtb_ordini " + vbCrLf + _
		"		SET ord_totale=(SELECT SUM(det_totale) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale IS NOT NULL) , " + vbCrLf + _
		"			ord_totale_iva=(SELECT SUM(det_totale_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_iva IS NOT NULL) ,  " + vbCrLf + _
		"			ord_det_totale_spese=(SELECT SUM(det_totale_spese) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese IS NOT NULL) ,  " + vbCrLf + _
		"			ord_det_totale_spese_iva=(SELECT SUM(det_totale_spese_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese_iva IS NOT NULL)  " + vbCrLf + _	
		"		WHERE ord_id=@ord_id " + vbCrLf + _
		"		UPDATE gtb_ordini " + vbCrLf + _
		"		SET ord_totale_spese=ROUND(ISNULL(ord_spesespedizione,0) + " + vbCrLf + _
		"								   ISNULL(ord_speseincasso,0) + " + vbCrLf + _
		"								   ISNULL(ord_spesefisse,0) + " + vbCrLf + _
		"								   ISNULL(ord_spesealtre,0) + " + vbCrLf + _
		"								   ISNULL(ord_det_totale_spese,0),2) ,  " + vbCrLf + _
		"			ord_totale_spese_iva=ROUND(ISNULL(ord_spesespedizione,0)*ISNULL(ord_spesespedizione_iva,0)/100 + " + vbCrLf + _
		"									   ISNULL(ord_speseincasso,0)*ISNULL(ord_speseincasso_iva,0)/100 + " + vbCrLf + _
		"									   ISNULL(ord_spesefisse,0)*ISNULL(ord_spesefisse_iva,0)/100 + " + vbCrLf + _
		"									   ISNULL(ord_spesealtre,0)*ISNULL(ord_spesealtre_iva,0)/100 + " + vbCrLf + _
		"									   ISNULL(ord_det_totale_spese_iva,0),2)  " + vbCrLf + _
		"		WHERE ord_id=@ord_id " + vbCrLf + _
		"	END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		"	DECLARE rs CURSOR  " + vbCrLf + _
		"	READ_ONLY " + vbCrLf + _
		"	FOR SELECT ord_id FROM gtb_ordini " + vbCrLf + _
		"	DECLARE @name int " + vbCrLf + _
		" " + vbCrLf + _
		"	OPEN rs " + vbCrLf + _
		" " + vbCrLf + _
		"	FETCH NEXT FROM rs INTO @name " + vbCrLf + _
		" " + vbCrLf + _
		"	WHILE (@@fetch_status <> -1) " + vbCrLf + _
		"	BEGIN " + vbCrLf + _
		"		IF (@@fetch_status <> -2) " + vbCrLf + _
		"		BEGIN " + vbCrLf + _
		"			EXECUTE gsp_totale_ordini @name " + vbCrLf + _
		"		END " + vbCrLf + _
		"		FETCH NEXT FROM rs INTO @name " + vbCrLf + _
		"	END " + vbCrLf + _
		"	  " + vbCrLf + _
		"	CLOSE rs " + vbCrLf + _
		"	DEALLOCATE rs; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 142
'...........................................................................................
'	Sergio, 27/08/2010
'...........................................................................................
'   aggiunge parametro
'...........................................................................................
function Aggiornamento_B2B__142(conn)
	Aggiornamento_B2B__142 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale_B2B__142(conn)
	' sub AddParametroSito(conn, codice, raggruppamento_id, nome, unita, tipo, principale, immagine, admin, personalizzato, sito_id, valore_it, valore_en, valore_fr, valore_de, valore_es)
	dim id_home,rs
	id_home = GetValueList(conn,rs,"SELECT id_home_page FROM tb_webs WHERE id_webs = " & Application("AZ_ID"))
	
	CALL AddParametroSito(conn, "CATALOGO_HOME_PAGE", _
								null, _
								"pagina di inizio consultazione catalogo ", _
								"", _
								adGUID, _
								0, _
								"", _
								1, _
								1, _
								NEXTB2B, _
								id_home, null, null, null, null)
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 143
'...........................................................................................
'	Giacomo, 16/09/2010
'...........................................................................................
'   crea le stored procedure gsp_totale_ordini, gsp_totale_shopping_cart 
'	e i trigger grel_dett_cart_des_value_delete, grel_dett_cart_des_value_insert, grel_dett_cart_des_value_update,
'	grel_dettagli_ord_des_value_delete, grel_dettagli_ord_des_value_insert, grel_dettagli_ord_des_value_update,
'	gtb_dett_cart_delete, gtb_dettagli_ord_delete, gtb_dettagli_ord_insert, gtb_dett_cart_insert, gtb_dett_cart_update,
'	gtb_dettagli_ord_update, gtb_ordini_insert, gtb_shopping_cart_insert, gtb_ordini_update, gtb_shopping_cart_update
'...........................................................................................
function Aggiornamento__B2B__143(conn)
	Aggiornamento__B2B__143 = _
		DropObject(conn, "gsp_totale_ordini", "PROCEDURE") + vbCrLf + _
		" CREATE PROCEDURE [dbo].[gsp_totale_ordini] " + vbCrLf + _
		" 	@ord_id INT  " + vbCrLf + _
		" AS  " + vbCrLf + _
		" BEGIN  " + vbCrLf + _
		" 	IF (EXISTS(SELECT det_id  " + vbCrLf + _
		" 			  FROM gtb_dettagli_ord INNER JOIN  " + vbCrLf + _
		" 					gtb_dettagli_ord_tipo ON gtb_dettagli_ord.det_tipo_id = gtb_dettagli_ord_tipo.dot_id INNER JOIN " + vbCrLf + _
		" 					grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_tipo.dot_id = grel_dettagli_ord_tipo_des.rtd_tipo_id INNER JOIN " + vbCrLf + _
		" 					gtb_dettagli_ord_des ON grel_dettagli_ord_tipo_des.rtd_descrittore_id = gtb_dettagli_ord_des.dod_id INNER JOIN " + vbCrLf + _
		" 					grel_dettagli_ord_des_value ON ( gtb_dettagli_ord_des.dod_id = grel_dettagli_ord_des_value.rel_des_descrittore_id " + vbCrLF + _
			" 				 								 AND grel_dettagli_ord_des_value.rel_des_dett_ord_id = gtb_dettagli_ord.det_id ) " + vbCrLf + _
		" 			  WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 					IsNull(dod_percentuale_detrazione,0) <> 0 AND " + vbCrLf + _
		" 					det_ord_id = @ord_id " + vbCrLf + _
		" 			 )) BEGIN " + vbCrLf + _
		" 		--ci sono dei descrittori su riga che variano il conteggio della quantità su almeno un dettaglio " + vbCrLf + _
		" 		--uso un cursore per ogni dettaglio per fare i conti. " + vbCrLf + _
		" 		DECLARE @det_id INT, @det_tipo_id INT " + vbCrLf + _
		" 		DECLARE @det_qta REAL, @detrazione_qta REAL " + vbCrLf + _
		" 	" + vbCrLf + _
		" 		DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 		SELECT det_id, det_qta, det_tipo_id FROM gtb_dettagli_ord WHERE det_ord_id = @ord_id " + vbCrLf + _
		" 	" + vbCrLf + _
		" 		OPEN rs " + vbCrLf + _
		" 		FETCH NEXT FROM rs INTO @det_id, @det_qta, @det_tipo_id " + vbCrLf + _
		" 		WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" 		BEGIN " + vbCrLf + _
		" 			--calcolo quantità in detrazione per ogni singolo dettaglio " + vbCrLf + _
		" 			SELECT @detrazione_qta = SUM(CAST(IsNull(rel_des_valore_it,'0') AS real) * (CAST(dod_percentuale_detrazione AS real)/100)) " + vbCrLf + _
		" 				FROM grel_dettagli_ord_des_value INNER JOIN " + vbCrLf + _
		" 					 gtb_dettagli_ord_des ON grel_dettagli_ord_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id INNER JOIN " + vbCrLf + _
		" 					 grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_des.dod_id = grel_dettagli_ord_tipo_des.rtd_descrittore_id " + vbCrLf + _
		" 				WHERE rel_des_dett_ord_id = @det_id AND rtd_tipo_id = @det_tipo_id " + vbCrLf + _
		" 					  AND IsNull(dod_qta_in_detrazione,0)=1 " + vbCrLf + _
		" 					  AND IsNull(dod_percentuale_detrazione,0)<>0 " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> ''  " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> '0' " + vbCrLf + _
		"   " + vbCrLf + _
		" 			SET @det_qta = @det_qta - @detrazione_qta " + vbCrLf + _
		" 	" + vbCrLf + _
		" 			--conteggio quantità in base a valori derivati per il singolo dettaglio " + vbCrLf + _
		" 			UPDATE gtb_dettagli_ord   " + vbCrLf + _
		" 				SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(@det_qta,0),2) ,   " + vbCrLf + _
		" 					det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(@det_qta,0)*ISNULL(det_iva,0)/100,2) ,   " + vbCrLf + _
		" 					det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +   " + vbCrLf + _
		" 											ISNULL(det_speseincasso,0) +  " + vbCrLf + _
		" 											ISNULL(det_spesefisse,0)+  " + vbCrLf + _
		" 											ISNULL(det_spesealtre,0),2) ,   " + vbCrLf + _
		" 					det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 +  " + vbCrLf + _
		" 												ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 +  " + vbCrLf + _
		" 												ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 +  " + vbCrLf + _
		" 												ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)   " + vbCrLf + _
		" 				WHERE det_id=@det_id " + vbCrLf + _
		"    " + vbCrLf + _
		" 			FETCH NEXT FROM rs INTO @det_id, @det_qta, @det_tipo_id " + vbCrLf + _
		" 		END " + vbCrLf + _
		"    " + vbCrLf + _
		" 	END " + vbCrLf + _
		" 	ELSE  " + vbCrLf + _
		" 	BEGIN  " + vbCrLf + _
		" 		--calcolo normale dei totali per i dettagli dell'ordine " + vbCrLf + _
		" 		UPDATE gtb_dettagli_ord   " + vbCrLf + _
		" 		SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0),2) ,   " + vbCrLf + _
		" 			det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0)*ISNULL(det_iva,0)/100,2) ,   " + vbCrLf + _
		" 			det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +   " + vbCrLf + _
		" 									ISNULL(det_speseincasso,0) +  " + vbCrLf + _
		" 									ISNULL(det_spesefisse,0)+  " + vbCrLf + _
		" 									ISNULL(det_spesealtre,0),2) ,   " + vbCrLf + _
		" 			det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 +  " + vbCrLf + _
		" 										ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 +  " + vbCrLf + _
		" 										ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 +  " + vbCrLf + _
		" 										ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)   " + vbCrLf + _
		" 		WHERE det_ord_id=@ord_id " + vbCrLf + _
		" 	END  " + vbCrLf + _
		" 	 " + vbCrLf + _
		"    " + vbCrLf + _
		" 	--calcolo dei totali dei dettagli sulla testata dell'ordine " + vbCrLf + _
		" 	UPDATE gtb_ordini  " + vbCrLf + _
		" 	SET ord_totale=(SELECT SUM(det_totale) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale IS NOT NULL) ,  " + vbCrLf + _
		" 		ord_totale_iva=(SELECT SUM(det_totale_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_iva IS NOT NULL) ,   " + vbCrLf + _
		" 		ord_det_totale_spese=(SELECT SUM(det_totale_spese) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese IS NOT NULL) ,   " + vbCrLf + _
		" 		ord_det_totale_spese_iva=(SELECT SUM(det_totale_spese_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese_iva IS NOT NULL)   " + vbCrLf + _
		" 	WHERE ord_id=@ord_id  " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 	--calcolo dei totali generali dell'ordine " + vbCrLf + _
		" 	UPDATE gtb_ordini  " + vbCrLf + _
		" 	SET ord_totale_spese=ROUND(ISNULL(ord_spesespedizione,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_speseincasso,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_spesefisse,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_spesealtre,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_det_totale_spese,0),2) ,   " + vbCrLf + _
		" 		ord_totale_spese_iva=ROUND(ISNULL(ord_spesespedizione,0)*ISNULL(ord_spesespedizione_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_speseincasso,0)*ISNULL(ord_speseincasso_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_spesefisse,0)*ISNULL(ord_spesefisse_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_spesealtre,0)*ISNULL(ord_spesealtre_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_det_totale_spese_iva,0),2)   " + vbCrLf + _
		" 	WHERE ord_id=@ord_id  " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gsp_totale_shopping_cart", "PROCEDURE") + vbCrLf + _
		" CREATE PROCEDURE [dbo].[gsp_totale_shopping_cart]  " + vbCrLf + _
		" 	@sc_id INT  " + vbCrLf + _
		" AS  " + vbCrLf + _
		" BEGIN  " + vbCrLf + _
		" 	IF (EXISTS(SELECT dett_id  " + vbCrLf + _
		" 			  FROM gtb_dett_cart INNER JOIN " + vbCrLf + _
		" 					gtb_dettagli_ord_tipo ON gtb_dett_cart.dett_tipo_id = gtb_dettagli_ord_tipo.dot_id INNER JOIN " + vbCrLf + _
		" 					grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_tipo.dot_id = grel_dettagli_ord_tipo_des.rtd_tipo_id INNER JOIN " + vbCrLf + _
		" 					gtb_dettagli_ord_des ON grel_dettagli_ord_tipo_des.rtd_descrittore_id = gtb_dettagli_ord_des.dod_id INNER JOIN " + vbCrLf + _
		" 					grel_dett_cart_des_value ON ( gtb_dettagli_ord_des.dod_id = grel_dett_cart_des_value.rel_des_descrittore_id " + vbCrLF + _
		" 				 							      AND grel_dett_cart_des_value.rel_des_dett_cart_id = gtb_dett_cart.dett_id ) " + vbCrLf + _
		" 			  WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 					IsNull(dod_percentuale_detrazione,0) <> 0 AND " + vbCrLf + _
		" 					dett_cart_id = @sc_id " + vbCrLf + _
		" 			 )) BEGIN " + vbCrLf + _
		" 		--ci sono dei descrittori su riga che variano il conteggio della quantità su almeno un dettaglio " + vbCrLf + _
		" 		--uso un cursore per ogni dettaglio per fare i conti. " + vbCrLf + _
		" 		DECLARE @dett_id INT, @dett_tipo_id INT " + vbCrLf + _
		" 		DECLARE @dett_qta REAL, @detrazione_qta REAL " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 		DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 		SELECT dett_id, dett_qta, dett_tipo_id FROM gtb_dett_cart WHERE dett_cart_id = @sc_id " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 		OPEN rs " + vbCrLf + _
		" 		FETCH NEXT FROM rs INTO @dett_id, @dett_qta, @dett_tipo_id " + vbCrLf + _
		" 		WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" 		BEGIN " + vbCrLf + _
		" 			--calcolo quantità in detrazione per ogni singolo dettaglio " + vbCrLf + _
		" 			SELECT @detrazione_qta = SUM(CAST(IsNull(rel_des_valore_it,'0') AS real) * (CAST(dod_percentuale_detrazione AS real)/100)) " + vbCrLf + _
		" 				FROM grel_dett_cart_des_value INNER JOIN " + vbCrLf + _
		" 					 gtb_dettagli_ord_des ON grel_dett_cart_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id INNER JOIN " + vbCrLf + _
		" 					 grel_dettagli_ord_tipo_des ON gtb_dettagli_ord_des.dod_id = grel_dettagli_ord_tipo_des.rtd_descrittore_id " + vbCrLf + _
		" 				WHERE rel_des_dett_cart_id = @dett_id AND rtd_tipo_id = @dett_tipo_id " + vbCrLf + _
		" 					  AND IsNull(dod_qta_in_detrazione,0)=1 " + vbCrLf + _
		" 					  AND IsNull(dod_percentuale_detrazione,0)<>0 " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> ''  " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> '0' " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 			SET @dett_qta = @dett_qta - @detrazione_qta " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 			--calcolo normale dei totali per i dettagli della shopping cart " + vbCrLf + _
		" 			UPDATE gtb_dett_cart " + vbCrLf + _
		" 				SET dett_totale= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0),2) ,   " + vbCrLf + _
		" 					dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_iva_id),0)/100,2) ,   " + vbCrLf + _
		" 					dett_totale_spese= ROUND(ISNULL(dett_spesespedizione,0) +   " + vbCrLf + _
		" 											 ISNULL(dett_speseincasso,0) +  " + vbCrLf + _
		" 											 ISNULL(dett_spesefisse,0)+  " + vbCrLf + _
		" 											 ISNULL(dett_spesealtre,0),2) ,   " + vbCrLf + _
		" 					dett_totale_spese_iva = ROUND(ISNULL(dett_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesespedizione_iva_id),0)/100 +  " + vbCrLf + _
		" 												  ISNULL(dett_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_speseincasso_iva_id),0)/100 +  " + vbCrLf + _
		" 												  ISNULL(dett_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesefisse_iva_id),0)/100 +  " + vbCrLf + _
		" 												  ISNULL(dett_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesealtre_iva_id),0)/100,2)   " + vbCrLf + _
		" 				WHERE dett_id=@dett_id " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 			FETCH NEXT FROM rs INTO @dett_id, @dett_qta, @dett_tipo_id " + vbCrLf + _
		" 		END " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 	END " + vbCrLf + _
		" 	ELSE  " + vbCrLf + _
		" 	BEGIN " + vbCrLf + _
		" 		--calcolo dei totali dei dettagli sulla testata della shopping cart " + vbCrLf + _
		" 		UPDATE gtb_dett_cart   " + vbCrLf + _
		" 		SET dett_totale= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0),2) ,   " + vbCrLf + _
		" 			dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_iva_id),0)/100,2) ,   " + vbCrLf + _
		" 			dett_totale_spese= ROUND(ISNULL(dett_spesespedizione,0) +   " + vbCrLf + _
		" 									 ISNULL(dett_speseincasso,0) +  " + vbCrLf + _
		" 									 ISNULL(dett_spesefisse,0)+  " + vbCrLf + _
		" 									 ISNULL(dett_spesealtre,0),2) ,   " + vbCrLf + _
		" 			dett_totale_spese_iva= ROUND(ISNULL(dett_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesespedizione_iva_id),0)/100 +  " + vbCrLf + _
		" 										ISNULL(dett_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_speseincasso_iva_id),0)/100 +  " + vbCrLf + _
		" 										ISNULL(dett_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesefisse_iva_id),0)/100 +  " + vbCrLf + _
		" 										ISNULL(dett_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesealtre_iva_id),0)/100,2)   " + vbCrLf + _
		" 		WHERE dett_cart_id=@sc_id " + vbCrLf + _
		" 	END  " + vbCrLf + _
		" 	  " + vbCrLf + _
		" 	--calcolo dei totali dei dettagli sulla testata della shopping cart " + vbCrLf + _
		" 	UPDATE gtb_shopping_cart  " + vbCrLf + _
		" 	SET sc_totale=(SELECT SUM(dett_totale) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale IS NOT NULL) ,  " + vbCrLf + _
		" 		sc_totale_iva=(SELECT SUM(dett_totale_iva) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_iva IS NOT NULL) ,   " + vbCrLf + _
		" 		sc_dett_totale_spese=(SELECT SUM(dett_totale_spese) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_spese IS NOT NULL) ,   " + vbCrLf + _
		" 		sc_dett_totale_spese_iva=(SELECT SUM(dett_totale_spese_iva) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_spese_iva IS NOT NULL)   " + vbCrLf + _
		" 	WHERE sc_id=@sc_id  " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 	--calcolo dei totali generali della shopping cart " + vbCrLf + _
		" 	UPDATE gtb_shopping_cart  " + vbCrLf + _
		" 	SET sc_totale_spese=ROUND(ISNULL(sc_spesespedizione,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_speseincasso,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_spesefisse,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_spesealtre,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_dett_totale_spese,0),2) ,   " + vbCrLf + _
		" 		sc_totale_spese_iva=ROUND(ISNULL(sc_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesespedizione_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_speseincasso_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesefisse_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesealtre_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_dett_totale_spese_iva,0),2)   " + vbCrLf + _
		" 	WHERE sc_id=@sc_id  " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "grel_dettagli_ord_des_value_delete", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [grel_dettagli_ord_des_value_delete] " + vbCrLf + _
		" ON [dbo].[grel_dettagli_ord_des_value] " + vbCrLf + _
		" AFTER DELETE " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini ai quali è stata rimossa una  " + vbCrLf + _
		" informazione di riga con effetto sulla quantità del dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT det_ord_id  " + vbCrLf + _
		" FROM deleted  " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord ON deleted.rel_des_dett_ord_id = gtb_dettagli_ord.det_id " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord_des ON deleted.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 	  IsNull(dod_percentuale_detrazione,0) <> 0 " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "grel_dettagli_ord_des_value_insert", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [grel_dettagli_ord_des_value_insert] " + vbCrLf + _
		" ON [dbo].[grel_dettagli_ord_des_value] " + vbCrLf + _
		" AFTER INSERT " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini ai quali è stata aggiunta una  " + vbCrLf + _
		" informazione di riga con effetto sulla quantità del dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT det_ord_id  " + vbCrLf + _
		" FROM inserted  " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord ON inserted.rel_des_dett_ord_id = gtb_dettagli_ord.det_id " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord_des ON inserted.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 	  IsNull(dod_percentuale_detrazione,0) <> 0 " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "grel_dettagli_ord_des_value_update", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [grel_dettagli_ord_des_value_update] " + vbCrLf + _
		" ON [dbo].[grel_dettagli_ord_des_value] " + vbCrLf + _
		" AFTER UPDATE " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini ai quali è stata modificata una " + vbCrLf + _
		" informazione di riga con effetto sulla quantità del dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT det_ord_id  " + vbCrLf + _
		" FROM deleted  " + vbCrLf + _
		" INNER JOIN inserted ON (deleted.rel_des_id = inserted.rel_des_id AND deleted.rel_des_valore_it <> inserted.rel_des_valore_it) " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord ON deleted.rel_des_dett_ord_id = gtb_dettagli_ord.det_id " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord_des ON deleted.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 	  (IsNull(inserted.rel_des_valore_it,'') <> '' OR IsNull(deleted.rel_des_valore_it,'') <> '')AND " + vbCrLf + _
		" 	  (IsNull(inserted.rel_des_valore_it,'') <> '0' OR IsNull(deleted.rel_des_valore_it,'') <> '0')AND " + vbCrLf + _
		" 	  IsNull(dod_percentuale_detrazione,0) <> 0 " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "grel_dett_cart_des_value_update", "TRIGGER") + vbCrLf + _
		" create TRIGGER [grel_dett_cart_des_value_update] " + vbCrLf + _
		" ON [dbo].[grel_dett_cart_des_value] " + vbCrLf + _
		" AFTER UPDATE " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart alle quali è stata modificata una  " + vbCrLf + _
		" informazione di riga con effetto sulla quantità del dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT dett_cart_id  " + vbCrLf + _
		" FROM deleted  " + vbCrLf + _
		" INNER JOIN inserted ON (deleted.rel_des_id = inserted.rel_des_id AND deleted.rel_des_valore_it <> inserted.rel_des_valore_it) " + vbCrLf + _
		" INNER JOIN dbo.gtb_dett_cart ON deleted.rel_des_dett_cart_id = gtb_dett_cart.dett_id " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord_des ON deleted.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 	  (IsNull(inserted.rel_des_valore_it,'') <> '' OR IsNull(deleted.rel_des_valore_it,'') <> '')AND " + vbCrLf + _
		" 	  (IsNull(inserted.rel_des_valore_it,'') <> '0' OR IsNull(deleted.rel_des_valore_it,'') <> '0')AND " + vbCrLf + _
		" 	  IsNull(dod_percentuale_detrazione,0) <> 0 " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "grel_dett_cart_des_value_insert", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [grel_dett_cart_des_value_insert] " + vbCrLf + _
		" ON [dbo].[grel_dett_cart_des_value] " + vbCrLf + _
		" AFTER INSERT " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart alle quali è stata rimossa una  " + vbCrLf + _
		" informazione di riga con effetto sulla quantità del dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT dett_cart_id  " + vbCrLf + _
		" FROM inserted  " + vbCrLf + _
		" INNER JOIN dbo.gtb_dett_cart ON inserted.rel_des_dett_cart_id = gtb_dett_cart.dett_id " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord_des ON inserted.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 	  IsNull(dod_percentuale_detrazione,0) <> 0 " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "grel_dett_cart_des_value_delete", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [grel_dett_cart_des_value_delete] " + vbCrLf + _
		" ON [dbo].[grel_dett_cart_des_value] " + vbCrLf + _
		" AFTER DELETE " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart alle quali è stata rimossa una  " + vbCrLf + _
		" informazione di riga con effetto sulla quantità del dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT dett_cart_id  " + vbCrLf + _
		" FROM deleted  " + vbCrLf + _
		" INNER JOIN dbo.gtb_dett_cart ON deleted.rel_des_dett_cart_id = gtb_dett_cart.dett_id " + vbCrLf + _
		" INNER JOIN dbo.gtb_dettagli_ord_des ON deleted.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 	  IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 	  IsNull(dod_percentuale_detrazione,0) <> 0 " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_dett_cart_update", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_dett_cart_update] " + vbCrLf + _
		" ON [dbo].[gtb_dett_cart] " + vbCrLf + _
		" AFTER UPDATE " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart alle quali è stato modificato un dettaglio " + vbCrLf + _
		" in almeno uno dei campi che concorrono al calcolo dei totali " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 	SELECT DISTINCT inserted.dett_cart_id FROM " + vbCrLf + _
		" 	inserted INNER JOIN deleted ON  " + vbCrLf + _
		" 		inserted.dett_id = deleted.dett_id " + vbCrLf + _
		" 		AND (  " + vbCrLf + _
		" 			inserted.dett_art_var_id <> deleted.dett_art_var_id OR " + vbCrLf + _
		" 			inserted.dett_qta <> deleted.dett_qta OR " + vbCrLf + _
		" 			inserted.dett_prezzo_unitario <> deleted.dett_prezzo_unitario OR " + vbCrLf + _
		" 			inserted.dett_iva_id <> deleted.dett_iva_id OR " + vbCrLf + _
		" 			inserted.dett_prezzo_listino <> deleted.dett_prezzo_listino OR " + vbCrLf + _
		" 			inserted.dett_sconto <> deleted.dett_sconto OR " + vbCrLf + _
		" 			inserted.dett_spesespedizione <> deleted.dett_spesespedizione OR " + vbCrLf + _
		" 			inserted.dett_speseincasso <> deleted.dett_speseincasso OR " + vbCrLf + _
		" 			inserted.dett_spesefisse <> deleted.dett_spesefisse OR " + vbCrLf + _
		" 			inserted.dett_spesealtre <> deleted.dett_spesealtre OR " + vbCrLf + _
		" 			inserted.dett_spesespedizione_iva_id <> deleted.dett_spesespedizione_iva_id OR " + vbCrLf + _
		" 			inserted.dett_speseincasso_iva_id <> deleted.dett_speseincasso_iva_id OR " + vbCrLf + _
		" 			inserted.dett_spesefisse_iva_id <> deleted.dett_spesefisse_iva_id OR " + vbCrLf + _
		" 			inserted.dett_spesealtre_iva_id <> deleted.dett_spesealtre_iva_id " + vbCrLf + _
		" 			) " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_dett_cart_insert", "TRIGGER") + vbCrLf + _
		" create TRIGGER [gtb_dett_cart_insert] " + vbCrLf + _
		" ON [dbo].[gtb_dett_cart] " + vbCrLf + _
		" AFTER INSERT " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart alle quali è stato aggiunto un dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT dett_cart_id FROM inserted " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_dett_cart_delete", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_dett_cart_delete] " + vbCrLf + _
		" ON [dbo].[gtb_dett_cart] " + vbCrLf + _
		" AFTER DELETE " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart alle quali è stato rimosso un dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT dett_cart_id FROM deleted " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_dettagli_ord_update", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_dettagli_ord_update] " + vbCrLf + _
		" ON [dbo].[gtb_dettagli_ord] " + vbCrLf + _
		" AFTER UPDATE " + vbCrLf + _
		" AS " + vbCrLf + _
		"  " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini ai quali è stato modificato un dettaglio " + vbCrLf + _
		" in almeno uno dei campi che concorrono al calcolo dei totali " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 	SELECT DISTINCT inserted.det_ord_id FROM " + vbCrLf + _
		" 	inserted INNER JOIN deleted ON  " + vbCrLf + _
		" 		inserted.det_id = deleted.det_id " + vbCrLf + _
		" 		AND (  " + vbCrLf + _
		" 			inserted.det_art_var_id <> deleted.det_art_var_id OR " + vbCrLf + _
		" 			inserted.det_qta <> deleted.det_qta OR " + vbCrLf + _
		" 			inserted.det_prezzo_unitario <> deleted.det_prezzo_unitario OR " + vbCrLf + _
		" 			inserted.det_iva <> deleted.det_iva OR " + vbCrLf + _
		" 			inserted.det_prezzo_listino <> deleted.det_prezzo_listino OR " + vbCrLf + _
		" 			inserted.det_sconto <> deleted.det_sconto OR " + vbCrLf + _
		" 			inserted.det_spesespedizione <> deleted.det_spesespedizione OR " + vbCrLf + _
		" 			inserted.det_speseincasso <> deleted.det_speseincasso OR " + vbCrLf + _
		" 			inserted.det_spesefisse <> deleted.det_spesefisse OR " + vbCrLf + _
		" 			inserted.det_spesealtre <> deleted.det_spesealtre OR " + vbCrLf + _
		" 			inserted.det_spesespedizione_iva <> deleted.det_spesespedizione_iva OR " + vbCrLf + _
		" 			inserted.det_speseincasso_iva <> deleted.det_speseincasso_iva OR " + vbCrLf + _
		" 			inserted.det_spesefisse_iva <> deleted.det_spesefisse_iva OR " + vbCrLf + _
		" 			inserted.det_spesealtre_iva <> deleted.det_spesealtre_iva " + vbCrLf + _
		" 			) " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_dettagli_ord_insert", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_dettagli_ord_insert] " + vbCrLf + _
		" ON [dbo].[gtb_dettagli_ord] " + vbCrLf + _
		" AFTER INSERT " + vbCrLf + _
		" AS " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini ai quali è stato aggiunto un dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT det_ord_id FROM inserted " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_dettagli_ord_delete", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_dettagli_ord_delete] " + vbCrLf + _
		" ON [dbo].[gtb_dettagli_ord] " + vbCrLf + _
		" AFTER DELETE " + vbCrLf + _
		" AS " + vbCrLf + _
		"  " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini ai quali è stato rimosso un dettaglio " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT det_ord_id FROM deleted " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_ordini_update", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_ordini_update] " + vbCrLf + _
		" ON [dbo].[gtb_ordini] " + vbCrLf + _
		" AFTER UPDATE " + vbCrLf + _
		" AS " + vbCrLf + _
		"  " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset degli ordini modificati " + vbCrLf + _
		" in almeno uno dei campi che concorrono al calcolo dei totali " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 	SELECT DISTINCT inserted.ord_id FROM " + vbCrLf + _
		" 	inserted INNER JOIN deleted ON  " + vbCrLf + _
		" 		inserted.ord_id = deleted.ord_id " + vbCrLf + _
		" 		AND (  " + vbCrLf + _
		" 			inserted.ord_spesespedizione <> deleted.ord_spesespedizione OR " + vbCrLf + _
		" 			inserted.ord_speseincasso <> deleted.ord_speseincasso OR " + vbCrLf + _
		" 			inserted.ord_spesefisse <> deleted.ord_spesefisse OR " + vbCrLf + _
		" 			inserted.ord_spesealtre <> deleted.ord_spesealtre OR " + vbCrLf + _
		" 			inserted.ord_spesespedizione_iva <> deleted.ord_spesespedizione_iva OR " + vbCrLf + _
		" 			inserted.ord_speseincasso_iva <> deleted.ord_speseincasso_iva OR " + vbCrLf + _
		" 			inserted.ord_spesefisse_iva <> deleted.ord_spesefisse_iva OR " + vbCrLf + _
		" 			inserted.ord_spesealtre_iva <> deleted.ord_spesealtre_iva " + vbCrLf + _
		" 			) " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_ordini_insert", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_ordini_insert] " + vbCrLf + _
		" ON [dbo].[gtb_ordini] " + vbCrLf + _
		" AFTER INSERT " + vbCrLf + _
		" AS " + vbCrLf + _
		"  " + vbCrLf + _
		" DECLARE @O_id INT " + vbCrLf + _
		" /*apre recordset con ordini inseriti*/ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT ord_id FROM inserted " + vbCrLf + _
		"  " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo dell'ordine */ " + vbCrLf + _
		" 	EXEC gsp_totale_ordini @ord_id=@O_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @O_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_shopping_cart_update", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_shopping_cart_update] " + vbCrLf + _
		" ON [dbo].[gtb_shopping_cart] " + vbCrLf + _
		" AFTER UPDATE " + vbCrLf + _
		" AS " + vbCrLf + _
		"  " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /* " + vbCrLf + _
		" apre recordset delle shopping cart modificate " + vbCrLf + _
		" in almeno uno dei campi che concorrono al calcolo dei totali " + vbCrLf + _
		" */ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 	SELECT DISTINCT inserted.sc_id FROM " + vbCrLf + _
		" 	inserted INNER JOIN deleted ON  " + vbCrLf + _
		" 		inserted.sc_id = deleted.sc_id " + vbCrLf + _
		" 		AND (  " + vbCrLf + _
		" 			inserted.sc_spesespedizione <> deleted.sc_spesespedizione OR " + vbCrLf + _
		" 			inserted.sc_speseincasso <> deleted.sc_speseincasso OR " + vbCrLf + _
		" 			inserted.sc_spesefisse <> deleted.sc_spesefisse OR " + vbCrLf + _
		" 			inserted.sc_spesealtre <> deleted.sc_spesealtre OR " + vbCrLf + _
		" 			inserted.sc_spesespedizione_iva_id <> deleted.sc_spesespedizione_iva_id OR " + vbCrLf + _
		" 			inserted.sc_speseincasso_iva_id <> deleted.sc_speseincasso_iva_id OR " + vbCrLf + _
		" 			inserted.sc_spesefisse_iva_id <> deleted.sc_spesefisse_iva_id OR " + vbCrLf + _
		" 			inserted.sc_spesealtre_iva_id <> deleted.sc_spesealtre_iva_id " + vbCrLf + _
		" 			) " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gtb_shopping_cart_insert", "TRIGGER") + vbCrLf + _
		" CREATE TRIGGER [gtb_shopping_cart_insert] " + vbCrLf + _
		" ON [dbo].[gtb_shopping_cart] " + vbCrLf + _
		" AFTER INSERT " + vbCrLf + _
		" AS " + vbCrLf + _
		"  " + vbCrLf + _
		" DECLARE @s_id INT " + vbCrLf + _
		" /*apre recordset con ordini inseriti*/ " + vbCrLf + _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" SELECT DISTINCT sc_id FROM inserted " + vbCrLf + _
		" OPEN rs " + vbCrLf + _
		" FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" BEGIN " + vbCrLf + _
		" 	/* esegue ricalcolo della shopping cart */ " + vbCrLf + _
		" 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id " + vbCrLf + _
		" 	FETCH NEXT FROM rs INTO @s_id " + vbCrLf + _
		" END "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 144
'...........................................................................................
'	Giacomo, 16/09/2010
'...........................................................................................
'   richiama per ogni ordine la stored procedure gsp_totale_ordini,
'   richiama per ogni ordine la stored procedure gsp_totale_shopping_cart	
'...........................................................................................
function Aggiornamento__B2B__144(conn)
	Aggiornamento__B2B__144 = _
		"	DECLARE rs CURSOR  " + vbCrLf + _
		"	READ_ONLY " + vbCrLf + _
		"	FOR SELECT ord_id FROM gtb_ordini " + vbCrLf + _
		"	DECLARE @name int " + vbCrLf + _
		" " + vbCrLf + _
		"	OPEN rs " + vbCrLf + _
		" " + vbCrLf + _
		"	FETCH NEXT FROM rs INTO @name " + vbCrLf + _
		" " + vbCrLf + _
		"	WHILE (@@fetch_status <> -1) " + vbCrLf + _
		"	BEGIN " + vbCrLf + _
		"		IF (@@fetch_status <> -2) " + vbCrLf + _
		"		BEGIN " + vbCrLf + _
		"			EXECUTE gsp_totale_ordini @name " + vbCrLf + _
		"		END " + vbCrLf + _
		"		FETCH NEXT FROM rs INTO @name " + vbCrLf + _
		"	END " + vbCrLf + _
		"	  " + vbCrLf + _
		"	CLOSE rs " + vbCrLf + _
		"	DEALLOCATE rs; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		"	DECLARE rs CURSOR  " + vbCrLf + _
		"	READ_ONLY " + vbCrLf + _
		"	FOR SELECT sc_id FROM gtb_shopping_cart " + vbCrLf + _
		"	DECLARE @name int " + vbCrLf + _
		" " + vbCrLf + _
		"	OPEN rs " + vbCrLf + _
		" " + vbCrLf + _
		"	FETCH NEXT FROM rs INTO @name " + vbCrLf + _
		" " + vbCrLf + _
		"	WHILE (@@fetch_status <> -1) " + vbCrLf + _
		"	BEGIN " + vbCrLf + _
		"		IF (@@fetch_status <> -2) " + vbCrLf + _
		"		BEGIN " + vbCrLf + _
		"			EXECUTE gsp_totale_shopping_cart @name " + vbCrLf + _
		"		END " + vbCrLf + _
		"		FETCH NEXT FROM rs INTO @name " + vbCrLf + _
		"	END " + vbCrLf + _
		"	  " + vbCrLf + _
		"	CLOSE rs " + vbCrLf + _
		"	DEALLOCATE rs; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 145
'...........................................................................................
'	Matteo, 02/11/2010
'...........................................................................................
'   aggiunge il campo per la quantità massima ordinabile in gtb_articoli
'...........................................................................................
function Aggiornamento__B2B__145(conn)
	Aggiornamento__B2B__145 = _
		"ALTER TABLE gtb_articoli ADD " + _
		" 	 art_qta_max_ord int NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 146
'...........................................................................................
'	Matteo, 25/11/2010
'...........................................................................................
'   aggiunge il campo per il valore del prezzo scontato in gtb_scontiQ
'...........................................................................................
function Aggiornamento__B2B__146(conn)
	Aggiornamento__B2B__146 = _
		"ALTER TABLE gtb_scontiQ ADD " + _
		" 	 sco_prezzo money NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 147
'...........................................................................................
'	Matteo, 10/12/2010
'...........................................................................................
'   modifica il tipo del campo precentuale detrazione in gtb_dettagli_ord_des da int a real
'...........................................................................................
function Aggiornamento__B2B__147(conn)
	Aggiornamento__B2B__147 = _
		"ALTER TABLE gtb_dettagli_ord_des ADD " + _
		" 	 dod_percentuale_detrazione_temp INT NULL; " + _
		"UPDATE gtb_dettagli_ord_des " + _
		"  SET dod_percentuale_detrazione_temp = dod_percentuale_detrazione; " + _
		"ALTER TABLE gtb_dettagli_ord_des DROP COLUMN " + _
		" 	 dod_percentuale_detrazione; " + _
		"ALTER TABLE gtb_dettagli_ord_des ADD " + _
		" 	 dod_percentuale_detrazione REAL NULL; " + _
		"UPDATE gtb_dettagli_ord_des " + _
		"  SET dod_percentuale_detrazione = dod_percentuale_detrazione_temp; " + _
		"ALTER TABLE gtb_dettagli_ord_des DROP COLUMN " + _
		" 	 dod_percentuale_detrazione_temp; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 148
'...........................................................................................
'	Matteo, 10/01/2011
'...........................................................................................
'   aggiunge il campo descrizione alla tabella delle tipologie di riga d'ordine
'...........................................................................................
function Aggiornamento__B2B__148(conn)
	Aggiornamento__B2B__148 = _
		" ALTER TABLE gtb_dettagli_ord_tipo ADD " + _
		" 	  dot_codice NVARCHAR (50) NULL , " + _
		SQL_MultiLanguageField("	dot_descrizione_<lingua> " + SQL_CharField(Conn, 255)) + " ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 149
'...........................................................................................
'	Matteo, 31/01/2011
'...........................................................................................
'   corregge funzioni di calcolo totali dell'ordine e delle shopping cart
'...........................................................................................
function Aggiornamento__B2B__149(conn)
	Aggiornamento__B2B__149 = _
	DropObject(conn, "gsp_totale_ordini", "PROCEDURE") + vbCrLf + _
		" CREATE PROCEDURE [dbo].[gsp_totale_ordini] " + vbCrLf + _
		" 	@ord_id INT  " + vbCrLf + _
		" AS  " + vbCrLf + _
		" BEGIN  " + vbCrLf + _
		" 	IF (EXISTS(SELECT det_id  " + vbCrLf + _
		" 			  FROM gtb_dettagli_ord " + vbCrLf + _
		"				INNER JOIN grel_dettagli_ord_des_value ON gtb_dettagli_ord.det_id = grel_dettagli_ord_des_value.rel_des_dett_ord_id " + vbCrLF + _
		"				INNER JOIN gtb_dettagli_ord_des ON grel_dettagli_ord_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" 			  WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 					IsNull(dod_percentuale_detrazione,0) <> 0 AND " + vbCrLf + _
		" 					det_ord_id = @ord_id " + vbCrLf + _
		" 			 )) BEGIN " + vbCrLf + _
		" 		--ci sono dei descrittori su riga che variano il conteggio della quantità su almeno un dettaglio " + vbCrLf + _
		" 		--uso un cursore per ogni dettaglio per fare i conti. " + vbCrLf + _
		" 		DECLARE @det_id INT " + vbCrLf + _
		" 		DECLARE @det_qta REAL, @detrazione_qta REAL " + vbCrLf + _
		" 	" + vbCrLf + _
		" 		DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 		SELECT det_id, det_qta FROM gtb_dettagli_ord WHERE det_ord_id = @ord_id " + vbCrLf + _
		" 	" + vbCrLf + _
		" 		OPEN rs " + vbCrLf + _
		" 		FETCH NEXT FROM rs INTO @det_id, @det_qta " + vbCrLf + _
		" 		WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" 		BEGIN " + vbCrLf + _
		" 			--calcolo quantità in detrazione per ogni singolo dettaglio " + vbCrLf + _
		" 			SELECT @detrazione_qta = SUM(CAST(IsNull(rel_des_valore_it,'0') AS real) * (CAST(dod_percentuale_detrazione AS real)/100)) " + vbCrLf + _
		" 				FROM grel_dettagli_ord_des_value INNER JOIN " + vbCrLf + _
		" 					 gtb_dettagli_ord_des ON grel_dettagli_ord_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" 				WHERE rel_des_dett_ord_id = @det_id " + vbCrLf + _
		" 					  AND IsNull(dod_qta_in_detrazione,0)=1 " + vbCrLf + _
		" 					  AND IsNull(dod_percentuale_detrazione,0)<>0 " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> ''  " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> '0' " + vbCrLf + _
		"   " + vbCrLf + _
		" 			SET @det_qta = @det_qta - @detrazione_qta " + vbCrLf + _
		" 	" + vbCrLf + _
		" 			--conteggio quantità in base a valori derivati per il singolo dettaglio " + vbCrLf + _
		" 			UPDATE gtb_dettagli_ord   " + vbCrLf + _
		" 				SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(@det_qta,0),2) ,   " + vbCrLf + _
		" 					det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(@det_qta,0)*ISNULL(det_iva,0)/100,2) ,   " + vbCrLf + _
		" 					det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +   " + vbCrLf + _
		" 											ISNULL(det_speseincasso,0) +  " + vbCrLf + _
		" 											ISNULL(det_spesefisse,0)+  " + vbCrLf + _
		" 											ISNULL(det_spesealtre,0),2) ,   " + vbCrLf + _
		" 					det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 +  " + vbCrLf + _
		" 												ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 +  " + vbCrLf + _
		" 												ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 +  " + vbCrLf + _
		" 												ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)   " + vbCrLf + _
		" 				WHERE det_id=@det_id " + vbCrLf + _
		"    " + vbCrLf + _
		" 			FETCH NEXT FROM rs INTO @det_id, @det_qta " + vbCrLf + _
		" 		END " + vbCrLf + _
		"    " + vbCrLf + _
		" 	END " + vbCrLf + _
		" 	ELSE  " + vbCrLf + _
		" 	BEGIN  " + vbCrLf + _
		" 		--calcolo normale dei totali per i dettagli dell'ordine " + vbCrLf + _
		" 		UPDATE gtb_dettagli_ord   " + vbCrLf + _
		" 		SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0),2) ,   " + vbCrLf + _
		" 			det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0)*ISNULL(det_iva,0)/100,2) ,   " + vbCrLf + _
		" 			det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +   " + vbCrLf + _
		" 									ISNULL(det_speseincasso,0) +  " + vbCrLf + _
		" 									ISNULL(det_spesefisse,0)+  " + vbCrLf + _
		" 									ISNULL(det_spesealtre,0),2) ,   " + vbCrLf + _
		" 			det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 +  " + vbCrLf + _
		" 										ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 +  " + vbCrLf + _
		" 										ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 +  " + vbCrLf + _
		" 										ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)   " + vbCrLf + _
		" 		WHERE det_ord_id=@ord_id " + vbCrLf + _
		" 	END  " + vbCrLf + _
		" 	 " + vbCrLf + _
		"    " + vbCrLf + _
		" 	--calcolo dei totali dei dettagli sulla testata dell'ordine " + vbCrLf + _
		" 	UPDATE gtb_ordini  " + vbCrLf + _
		" 	SET ord_totale=(SELECT SUM(det_totale) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale IS NOT NULL) ,  " + vbCrLf + _
		" 		ord_totale_iva=(SELECT SUM(det_totale_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_iva IS NOT NULL) ,   " + vbCrLf + _
		" 		ord_det_totale_spese=(SELECT SUM(det_totale_spese) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese IS NOT NULL) ,   " + vbCrLf + _
		" 		ord_det_totale_spese_iva=(SELECT SUM(det_totale_spese_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese_iva IS NOT NULL)   " + vbCrLf + _
		" 	WHERE ord_id=@ord_id  " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 	--calcolo dei totali generali dell'ordine " + vbCrLf + _
		" 	UPDATE gtb_ordini  " + vbCrLf + _
		" 	SET ord_totale_spese=ROUND(ISNULL(ord_spesespedizione,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_speseincasso,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_spesefisse,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_spesealtre,0) +  " + vbCrLf + _
		" 							   ISNULL(ord_det_totale_spese,0),2) ,   " + vbCrLf + _
		" 		ord_totale_spese_iva=ROUND(ISNULL(ord_spesespedizione,0)*ISNULL(ord_spesespedizione_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_speseincasso,0)*ISNULL(ord_speseincasso_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_spesefisse,0)*ISNULL(ord_spesefisse_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_spesealtre,0)*ISNULL(ord_spesealtre_iva,0)/100 +  " + vbCrLf + _
		" 								   ISNULL(ord_det_totale_spese_iva,0),2)   " + vbCrLf + _
		" 	WHERE ord_id=@ord_id  " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		DropObject(conn, "gsp_totale_shopping_cart", "PROCEDURE") + vbCrLf + _
		" CREATE PROCEDURE [dbo].[gsp_totale_shopping_cart]  " + vbCrLf + _
		" 	@sc_id INT  " + vbCrLf + _
		" AS  " + vbCrLf + _
		" BEGIN  " + vbCrLf + _
		" 	IF (EXISTS(SELECT dett_id  " + vbCrLf + _
		
		" 			  FROM gtb_dett_cart " + vbCrLf + _
		"				INNER JOIN grel_dett_cart_des_value ON gtb_dett_cart.dett_id = grel_dett_cart_des_value.rel_des_dett_cart_id " + vbCrLF + _
		"				INNER JOIN gtb_dettagli_ord_des ON grel_dett_cart_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" 			  WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '' AND " + vbCrLf + _
		" 					IsNull(rel_des_valore_it,'') <> '0' AND " + vbCrLf + _
		" 					IsNull(dod_percentuale_detrazione,0) <> 0 AND " + vbCrLf + _
		" 					dett_cart_id = @sc_id " + vbCrLf + _
		" 			 )) BEGIN " + vbCrLf + _
		" 		--ci sono dei descrittori su riga che variano il conteggio della quantità su almeno un dettaglio " + vbCrLf + _
		" 		--uso un cursore per ogni dettaglio per fare i conti. " + vbCrLf + _
		" 		DECLARE @dett_id INT " + vbCrLf + _
		" 		DECLARE @dett_qta REAL, @detrazione_qta REAL " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 		DECLARE rs CURSOR local FAST_FORWARD FOR  " + vbCrLf + _
		" 		SELECT dett_id, dett_qta FROM gtb_dett_cart WHERE dett_cart_id = @sc_id " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 		OPEN rs " + vbCrLf + _
		" 		FETCH NEXT FROM rs INTO @dett_id, @dett_qta " + vbCrLf + _
		" 		WHILE @@FETCH_STATUS = 0 " + vbCrLf + _
		" 		BEGIN " + vbCrLf + _
		" 			--calcolo quantità in detrazione per ogni singolo dettaglio " + vbCrLf + _
		" 			SELECT @detrazione_qta = SUM(CAST(IsNull(rel_des_valore_it,'0') AS real) * (CAST(dod_percentuale_detrazione AS real)/100)) " + vbCrLf + _
		" 				FROM grel_dett_cart_des_value INNER JOIN " + vbCrLf + _
		" 					 gtb_dettagli_ord_des ON grel_dett_cart_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id " + vbCrLf + _
		" 				WHERE rel_des_dett_cart_id = @dett_id  " + vbCrLf + _
		" 					  AND IsNull(dod_qta_in_detrazione,0)=1 " + vbCrLf + _
		" 					  AND IsNull(dod_percentuale_detrazione,0)<>0 " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> ''  " + vbCrLf + _
		" 					  AND IsNull(rel_des_valore_it,'') <> '0' " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 			SET @dett_qta = @dett_qta - @detrazione_qta " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 			--calcolo normale dei totali per i dettagli della shopping cart " + vbCrLf + _
		" 			UPDATE gtb_dett_cart " + vbCrLf + _
		" 				SET dett_totale= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0),2) ,   " + vbCrLf + _
		" 					dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_iva_id),0)/100,2) ,   " + vbCrLf + _
		" 					dett_totale_spese= ROUND(ISNULL(dett_spesespedizione,0) +   " + vbCrLf + _
		" 											 ISNULL(dett_speseincasso,0) +  " + vbCrLf + _
		" 											 ISNULL(dett_spesefisse,0)+  " + vbCrLf + _
		" 											 ISNULL(dett_spesealtre,0),2) ,   " + vbCrLf + _
		" 					dett_totale_spese_iva = ROUND(ISNULL(dett_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesespedizione_iva_id),0)/100 +  " + vbCrLf + _
		" 												  ISNULL(dett_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_speseincasso_iva_id),0)/100 +  " + vbCrLf + _
		" 												  ISNULL(dett_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesefisse_iva_id),0)/100 +  " + vbCrLf + _
		" 												  ISNULL(dett_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesealtre_iva_id),0)/100,2)   " + vbCrLf + _
		" 				WHERE dett_id=@dett_id " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 			FETCH NEXT FROM rs INTO @dett_id, @dett_qta " + vbCrLf + _
		" 		END " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 	END " + vbCrLf + _
		" 	ELSE  " + vbCrLf + _
		" 	BEGIN " + vbCrLf + _
		" 		--calcolo dei totali dei dettagli sulla testata della shopping cart " + vbCrLf + _
		" 		UPDATE gtb_dett_cart   " + vbCrLf + _
		" 		SET dett_totale= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0),2) ,   " + vbCrLf + _
		" 			dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_iva_id),0)/100,2) ,   " + vbCrLf + _
		" 			dett_totale_spese= ROUND(ISNULL(dett_spesespedizione,0) +   " + vbCrLf + _
		" 									 ISNULL(dett_speseincasso,0) +  " + vbCrLf + _
		" 									 ISNULL(dett_spesefisse,0)+  " + vbCrLf + _
		" 									 ISNULL(dett_spesealtre,0),2) ,   " + vbCrLf + _
		" 			dett_totale_spese_iva= ROUND(ISNULL(dett_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesespedizione_iva_id),0)/100 +  " + vbCrLf + _
		" 										ISNULL(dett_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_speseincasso_iva_id),0)/100 +  " + vbCrLf + _
		" 										ISNULL(dett_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesefisse_iva_id),0)/100 +  " + vbCrLf + _
		" 										ISNULL(dett_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesealtre_iva_id),0)/100,2)   " + vbCrLf + _
		" 		WHERE dett_cart_id=@sc_id " + vbCrLf + _
		" 	END  " + vbCrLf + _
		" 	  " + vbCrLf + _
		" 	--calcolo dei totali dei dettagli sulla testata della shopping cart " + vbCrLf + _
		" 	UPDATE gtb_shopping_cart  " + vbCrLf + _
		" 	SET sc_totale=(SELECT SUM(dett_totale) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale IS NOT NULL) ,  " + vbCrLf + _
		" 		sc_totale_iva=(SELECT SUM(dett_totale_iva) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_iva IS NOT NULL) ,   " + vbCrLf + _
		" 		sc_dett_totale_spese=(SELECT SUM(dett_totale_spese) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_spese IS NOT NULL) ,   " + vbCrLf + _
		" 		sc_dett_totale_spese_iva=(SELECT SUM(dett_totale_spese_iva) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_spese_iva IS NOT NULL)   " + vbCrLf + _
		" 	WHERE sc_id=@sc_id  " + vbCrLf + _
		" 	 " + vbCrLf + _
		" 	--calcolo dei totali generali della shopping cart " + vbCrLf + _
		" 	UPDATE gtb_shopping_cart  " + vbCrLf + _
		" 	SET sc_totale_spese=ROUND(ISNULL(sc_spesespedizione,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_speseincasso,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_spesefisse,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_spesealtre,0) +  " + vbCrLf + _
		" 							  ISNULL(sc_dett_totale_spese,0),2) ,   " + vbCrLf + _
		" 		sc_totale_spese_iva=ROUND(ISNULL(sc_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesespedizione_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_speseincasso_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesefisse_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesealtre_iva_id),0)/100 +  " + vbCrLf + _
		" 								  ISNULL(sc_dett_totale_spese_iva,0),2)   " + vbCrLf + _
		" 	WHERE sc_id=@sc_id  " + vbCrLf + _
		" END; " + vbCrLf + _
		" " + vbCrLf + _
		" " + vbCrLf + _
		"	DECLARE rs CURSOR  " + vbCrLf + _
		"	READ_ONLY " + vbCrLf + _
		"	FOR SELECT ord_id FROM gtb_ordini WHERE year(ord_data) = 2011 " + vbCrLf + _
		"	DECLARE @ord_id int " + vbCrLf + _
		" " + vbCrLf + _
		"	OPEN rs " + vbCrLf + _
		" " + vbCrLf + _
		"	FETCH NEXT FROM rs INTO @ord_id " + vbCrLf + _
		" " + vbCrLf + _
		"	WHILE (@@fetch_status <> -1) " + vbCrLf + _
		"	BEGIN " + vbCrLf + _
		"		IF (@@fetch_status <> -2) " + vbCrLf + _
		"		BEGIN " + vbCrLf + _
		"			EXECUTE gsp_totale_ordini @ord_id " + vbCrLf + _
		"		END " + vbCrLf + _
		"		FETCH NEXT FROM rs INTO @ord_id " + vbCrLf + _
		"	END " + vbCrLf + _
		"	  " + vbCrLf + _
		"	CLOSE rs " + vbCrLf + _
		"	DEALLOCATE rs; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 150
'...........................................................................................
'	Andrea, 3/02/2011
'...........................................................................................
function Aggiornamento__B2B__150(conn)
	Aggiornamento__B2B__150 = _
		"CREATE NONCLUSTERED INDEX [IX_gtb_art_foto_fo_articolo_id] ON [dbo].gtb_art_foto " + vbCrLf +_
		"( fo_articolo_id ASC ) ON [PRIMARY] "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 151
'...........................................................................................
'	Giacomo, 3/02/2011
' 	aggiungo la tabella dei profili per i clienti
'...........................................................................................
function Aggiornamento__B2B__151(conn)
	Aggiornamento__B2B__151 = _
			" ALTER TABLE gtb_dettagli_ord ADD " + _
			"	det_codice " + SQL_CharField(Conn, 255) + " NULL " + _
			"; " + _
			" ALTER TABLE gtb_marche ADD " + _
			" 	mar_anagrafica_id int NULL " + _
			"; " + _
			SQL_AddForeignKey(conn, "gtb_marche", "mar_anagrafica_id", "gtb_rivenditori", "riv_id", false, "") + _
			" " + _
			"CREATE TABLE " + SQL_Dbo(conn) + "gtb_profili(" + _
			"	pro_id " + SQL_PrimaryKey(conn, "gtb_profili") + ", " + _
			SQL_MultiLanguageFieldComplete(conn, "pro_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "," + _
			" 	pro_pagina_id int NULL, " + _
			" 	pro_insAdmin_id int NULL, " + _
			" 	pro_insData DATETIME NULL, " + _
			" 	pro_modAdmin_id int NULL, " + _
			" 	pro_modData DATETIME NULL " + _			
			"); " + _
			SQL_AddForeignKey(conn, "gtb_profili", "pro_insAdmin_id", "tb_admin", "ID_admin", false, "") + _
			SQL_AddForeignKey(conn, "gtb_profili", "pro_modAdmin_id", "tb_admin", "ID_admin", false, "FK_gtb_profili__tb_admin_2") + _
			" " + _
			" ALTER TABLE gtb_rivenditori ADD " + _
			" 	riv_profilo_id int NULL " + _
			"; " + _
			SQL_AddForeignKey(conn, "gtb_rivenditori", "riv_profilo_id", "gtb_profili", "pro_id", false, "") 
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 152
'...........................................................................................
'	Giacomo, 7/02/2011
'...........................................................................................
function Aggiornamento__B2B__152(conn)
	Aggiornamento__B2B__152 = _
			" ALTER TABLE gtb_profili ADD " + _
			" 	pro_codice " + SQL_CharField(Conn, 100) + " NULL, " + _
			" 	pro_rubrica_id int NULL " + _
			"; " + _
			SQL_AddForeignKey(conn, "gtb_profili", "pro_rubrica_id", "tb_rubriche", "id_rubrica", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 153
'...........................................................................................
' Giacomo 18/04/2011
'...........................................................................................
' aggiunge parametro per gestire gli avvisi degli impegni
'...........................................................................................
function Aggiornamento__B2B__153(conn)
	Aggiornamento__B2B__153 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__153(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "ART_COD_INT_UNIVOCO", _
									0, _
									"Durante il salvataggio di un articolo viene controllato che il codice interno, art_cod_int, sia univoco.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									1, null, null, null, null)
	end if
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 154
'...........................................................................................
'Giacomo - 10/05/2011
'...........................................................................................
'
'...........................................................................................
function Aggiornamento__B2B__154(conn)
	Aggiornamento__B2B__154 = _
		DropObject(conn, "fn_articolo_lista_codici_varianti", "FUNCTION") + vbCrLf + _
		" CREATE FUNCTION [dbo].[fn_articolo_lista_codici_varianti]  " + vbCrLf + _
		"  (   " + vbCrLf + _
		"		@ArtId int " + vbCrLf + _
		"  )   " + vbCrLf + _
		" RETURNS nvarchar(3000)  " + vbCrLf + _
		" AS   " + vbCrLf + _
		" BEGIN   " + vbCrLf + _
		"		DECLARE @cod_int nvarchar(255), @cod_pro nvarchar(255), @cod_alt nvarchar(255) " + vbCrLf + _
		"		DECLARE @codici nvarchar(3000) " + vbCrLf + _
		"	  SELECT @codici = ''  " + vbCrLf + _
		" " + vbCrLf + _
		"		DECLARE RS CURSOR FOR " + vbCrLf + _
		"			SELECT rel_cod_int, rel_cod_alt, rel_cod_pro FROM grel_art_valori WHERE rel_art_id= @ArtId " + vbCrLf + _
		"		OPEN RS " + vbCrLf + _
		" " + vbCrLf + _	
		"		FETCH NEXT FROM RS INTO @cod_int, @cod_pro, @cod_alt " + vbCrLf + _
		" " + vbCrLf + _
		"	  WHILE (@@fetch_status <> -1) " + vbCrLf + _
		"			BEGIN " + vbCrLf + _
		"				  IF (@@fetch_status <> -2) " + vbCrLf + _
		"				  BEGIN " + vbCrLf + _
		"						IF(IsNull(@cod_int, '') <> '') " + vbCrLf + _
		"							 SET @codici = @codici + ' ' + @cod_int " + vbCrLf + _
		"						IF(IsNull(@cod_pro, '') <> '') " + vbCrLf + _
		"							 SET @codici = @codici + ' ' + @cod_pro " + vbCrLf + _
		"						IF(IsNull(@cod_alt, '') <> '') " + vbCrLf + _
		"							 SET @codici = @codici + ' ' + @cod_alt " + vbCrLf + _
		"				  END " + vbCrLf + _
		"				  FETCH NEXT FROM RS INTO @cod_int, @cod_pro, @cod_alt " + vbCrLf + _
		"			END " + vbCrLf + _
		" " + vbCrLf + _
		"		CLOSE RS " + vbCrLf + _
		"	  DEALLOCATE RS " + vbCrLf + _
		"		RETURN  @codici   " + vbCrLf + _
		" END "
end function
'********************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 155
'...........................................................................................
'	Giacomo, 19/05/2011
'...........................................................................................
function Aggiornamento__B2B__155(conn)
	Aggiornamento__B2B__155 = _
			" ALTER TABLE gtb_rivenditori ADD " + _
			" 	riv_azienda_capogruppo_id int NULL " + _
			"; " + _
			SQL_AddForeignKey(conn, "gtb_rivenditori", "riv_azienda_capogruppo_id", "gtb_rivenditori", "riv_id", false, "")
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 156
'...........................................................................................
'	Giacomo, 24/05/2011 - Aggiunta campi nuova lingua per il B2B
'...........................................................................................
function Aggiornamento__B2B__156(conn, lingua_abbr)
	Aggiornamento__B2B__156 = _
		  " ALTER TABLE grel_art_acc ADD " + vbCrLf + _
		  " 	aa_note_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL;" + vbCrLf + _
		  " ALTER TABLE grel_art_ctech ADD " + vbCrLf + _
		  " 	rel_ctech_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE grel_art_valori ADD " + vbCrLf + _
		  " 	rel_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	rel_descr_prezzo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE grel_dett_cart_des_value ADD " + vbCrLf + _
		  " 	rel_des_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	rel_des_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE grel_dettagli_ord_des_value ADD " + vbCrLf + _
		  " 	rel_des_valore_" + lingua_abbr + " " + SQL_CharField(Conn, 250) + " NULL," + vbCrLf + _
		  " 	rel_des_memo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_accessori_tipo ADD " + vbCrLf + _
		  " 	at_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_art_foto ADD " + vbCrLf + _
		  " 	fo_didascalia_" + lingua_abbr + " " + SQL_CharField(Conn, 510) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_articoli ADD " + vbCrLf + _
		  " 	art_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL," + vbCrLf + _
		  " 	art_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	art_composizione_note_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	art_accessori_note_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	art_descr_riassunto_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	art_descr_prezzo_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	art_url_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_carattech ADD " + vbCrLf + _
		  " 	ct_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 510) + " NULL," + vbCrLf + _
		  " 	ct_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 100) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_carattech_raggruppamenti ADD " + vbCrLf + _
		  " 	ctr_titolo_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + vbCrLf + _  
		  " ALTER TABLE gtb_dett_cart ADD " + vbCrLf + _
		  " 	dett_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_dettagli_ord ADD " + vbCrLf + _
		  " 	det_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL;" + vbCrLf + _  
		  " ALTER TABLE gtb_dettagli_ord_des ADD " + vbCrLf + _
		  " 	dod_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	dod_unita_" + lingua_abbr + " " + SQL_CharField(Conn, 50) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_dettagli_ord_tipo ADD " + vbCrLf + _
		  " 	dot_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	dot_descrizione_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL;" + vbCrLf + _		  
		  " ALTER TABLE gtb_marche ADD " + vbCrLf + _
		  " 	mar_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 100) + " NULL," + vbCrLf + _
		  " 	mar_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_modipagamento ADD " + vbCrLf + _
		  " 	mosp_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	mosp_label_spsp_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL," + vbCrLf + _
		  " 	mosp_istruzioni_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _	
		  " ALTER TABLE gtb_spese_spedizione ADD " + vbCrLf + _
		  " 	sp_area_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	sp_condizioni_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_spese_spedizione_articolo ADD " + vbCrLf + _
		  " 	spa_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 255) + " NULL," + vbCrLf + _
		  " 	spa_condizioni_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_stati_ordine ADD " + vbCrLf + _
		  " 	so_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _	
		  " ALTER TABLE gtb_tipologie ADD " + vbCrLf + _
		  " 	tip_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL," + vbCrLf + _
		  " 	tip_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_tipologie_raggruppamenti ADD " + vbCrLf + _
		  " 	rag_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL," + vbCrLf + _
		  " 	rag_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_valori ADD " + vbCrLf + _
		  " 	val_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 100) + " NULL," + vbCrLf + _
		  " 	val_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;" + vbCrLf + _
		  " ALTER TABLE gtb_varianti ADD " + vbCrLf + _
		  " 	var_nome_" + lingua_abbr + " " + SQL_CharField(Conn, 500) + " NULL," + vbCrLf + _
		  " 	var_descr_" + lingua_abbr + " " + SQL_CharField(Conn, 0) + " NULL;"
end function
'*******************************************************************************************




'*******************************************************************************************
'AGGIORNAMENTO B2B 157
'..........................................................................................................................
'Giacomo - 24/05/2011
'..........................................................................................................................
'creazione viste gv_articoli divise per lingua
'..........................................................................................................................
function Aggiornamento__B2B__157(conn)

	DropObject conn,"gv_articoli_it","VIEW"
	DropObject conn,"gv_articoli_en","VIEW"
	DropObject conn,"gv_articoli_fr","VIEW"
	DropObject conn,"gv_articoli_es","VIEW"
	DropObject conn,"gv_articoli_de","VIEW"
	DropObject conn,"gv_articoli_cn","VIEW"
	DropObject conn,"gv_articoli_ru","VIEW"
	DropObject conn,"gv_articoli_pt","VIEW"

	Dim Agg,Agg_it,Agg_en,Agg_es,Agg_fr,Agg_de,Agg_cn,Agg_ru,Agg_pt
	
	Agg = _
		" SELECT art_id, art_nome_it, art_nome_en, art_descr_it, art_descr_en, art_cod_int, art_cod_pro, art_cod_alt, art_prezzo_base, " + vbCrLf + _
		"		art_scontoQ_id, art_giacenza_min, art_lotto_riordino, art_qta_min_ord, art_NovenSingola, art_se_accessorio, art_ha_accessori, " + vbCrLf + _
		"		art_se_bundle, art_se_confezione, art_in_bundle, art_in_confezione, art_varianti, art_disabilitato, art_tipologia_id, "  + vbCrLf + _
		"		art_marca_id, art_note, art_iva_id, art_composizione_note_it, art_composizione_note_en, art_accessori_note_it, art_accessori_note_en, " + vbCrLf + _
		"		art_external_id, art_raggruppamento_id, art_insData, art_insAdmin_id, art_modData, art_modAdmin_id, art_non_vendibile, " + vbCrLf + _
		"		art_applicativo_id, art_unico, art_descr_riassunto_it, art_descr_riassunto_en, art_descr_prezzo_it, art_descr_prezzo_en, " + vbCrLf + _
		"		art_spedizione_id, art_url_it, art_url_en, art_ordine, art_dettagli_ord_tipo_id, art_qta_max_ord, rel_id, rel_art_id, rel_prezzo, " + vbCrLf + _
		"		rel_giacenza_min, rel_lotto_riordino, rel_qta_min_ord, rel_cod_int, rel_cod_pro, rel_cod_alt, rel_disabilitato, rel_scontoQ_id, " + vbCrLf + _
		"		rel_ordine, rel_external_id, rel_var_euro, rel_var_sconto, rel_prezzo_indipendente, rel_foto_id, rel_descr_it, rel_descr_en, " + vbCrLf + _
		"		rel_insData, rel_insAdmin_id, rel_modData, rel_modAdmin_id, rel_non_vendibile, rel_descr_prezzo_it, rel_descr_prezzo_en, " + vbCrLf + _
		"		mar_id, mar_nome_it, mar_nome_en, mar_logo, mar_sito, mar_descr_it, mar_descr_en, mar_codice, mar_generica, mar_anagrafica_id, " + vbCrLf + _
		"		iva_id, iva_nome, iva_valore, iva_ordine, tip_id, tip_nome_it, tip_nome_en, tip_logo, tip_foto, tip_codice, tip_descr_it, tip_descr_en, " + vbCrLf + _
		"		tip_foglia, tip_livello, tip_padre_id, tip_ordine, tip_ordine_assoluto, tip_external_id, tip_tipologia_padre_base, tip_visibile, " + vbCrLf + _
		"		tip_albero_visibile, tip_tipologie_padre_lista, spa_id, spa_nome_it, spa_nome_en, spa_condizioni_it, spa_condizioni_en, " + vbCrLf + _
		"		spa_annullamento_qta, spa_annullamento_importo, spa_importo_spese " + vbCrLf + _
		" FROM gtb_articoli" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLf + _
		"		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id " + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";"
		
	Agg_it=" CREATE VIEW " & SQL_Dbo(conn) & "gv_articoli_it AS " + vbCrLf + Agg
	Agg_en=" CREATE VIEW " & SQL_Dbo(conn) & "gv_articoli_en AS " + vbCrLf + Agg
	Agg_cn=	_
		" CREATE VIEW " & SQL_Dbo(conn) & "gv_articoli_cn AS " + vbCrLf + _
		" SELECT art_id, art_nome_it, art_nome_en, art_nome_cn, art_descr_it, art_descr_en, art_descr_cn, art_cod_int, art_cod_pro, art_cod_alt, art_prezzo_base, " + vbCrLf + _
		"		art_scontoQ_id, art_giacenza_min, art_lotto_riordino, art_qta_min_ord, art_NovenSingola, art_se_accessorio, art_ha_accessori, " + vbCrLf + _
		"		art_se_bundle, art_se_confezione, art_in_bundle, art_in_confezione, art_varianti, art_disabilitato, art_tipologia_id, "  + vbCrLf + _
		"		art_marca_id, art_note, art_iva_id, art_composizione_note_it, art_composizione_note_en, art_composizione_note_cn, art_accessori_note_it, art_accessori_note_en, " + vbCrLf + _
		"		art_accessori_note_cn, art_external_id, art_raggruppamento_id, art_insData, art_insAdmin_id, art_modData, art_modAdmin_id, art_non_vendibile, " + vbCrLf + _
		"		art_applicativo_id, art_unico, art_descr_riassunto_it, art_descr_riassunto_en, art_descr_riassunto_cn, art_descr_prezzo_it, art_descr_prezzo_en, " + vbCrLf + _
		"		art_descr_prezzo_cn, art_spedizione_id, art_url_it, art_url_en, art_url_cn, art_ordine, art_dettagli_ord_tipo_id, art_qta_max_ord, rel_id, rel_art_id, rel_prezzo, " + vbCrLf + _
		"		rel_giacenza_min, rel_lotto_riordino, rel_qta_min_ord, rel_cod_int, rel_cod_pro, rel_cod_alt, rel_disabilitato, rel_scontoQ_id, " + vbCrLf + _
		"		rel_ordine, rel_external_id, rel_var_euro, rel_var_sconto, rel_prezzo_indipendente, rel_foto_id, rel_descr_it, rel_descr_en, " + vbCrLf + _
		"		rel_descr_cn, rel_insData, rel_insAdmin_id, rel_modData, rel_modAdmin_id, rel_non_vendibile, rel_descr_prezzo_it, rel_descr_prezzo_en, rel_descr_prezzo_cn, " + vbCrLf + _
		"		mar_id, mar_nome_it, mar_nome_en, mar_nome_cn, mar_logo, mar_sito, mar_descr_it, mar_descr_en, mar_descr_cn, mar_codice, mar_generica, mar_anagrafica_id, " + vbCrLf + _
		"		iva_id, iva_nome, iva_valore, iva_ordine, tip_id, tip_nome_it, tip_nome_en, tip_nome_cn, tip_logo, tip_foto, tip_codice, tip_descr_it, tip_descr_en, " + vbCrLf + _
		"		tip_descr_cn, tip_foglia, tip_livello, tip_padre_id, tip_ordine, tip_ordine_assoluto, tip_external_id, tip_tipologia_padre_base, tip_visibile, " + vbCrLf + _
		"		tip_albero_visibile, tip_tipologie_padre_lista, spa_id, spa_nome_it, spa_nome_en, spa_nome_cn, spa_condizioni_it, spa_condizioni_en, " + vbCrLf + _
		"		spa_condizioni_cn, spa_annullamento_qta, spa_annullamento_importo, spa_importo_spese " + vbCrLf + _
		" FROM gtb_articoli" + vbCrLf + _
		"		INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id" + vbCrLf + _
		"		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id" + vbCrLf + _
		"		INNER JOIN gtb_iva ON gtb_articoli.art_iva_id = gtb_iva.iva_id " + vbCrLf + _
		"		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id" + vbCrLf + _
		"		INNER JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";"
		
	Agg_es = Replace(Agg_cn,"_cn","_es")
	Agg_fr = Replace(Agg_cn,"_cn","_fr")
	Agg_ru = Replace(Agg_cn,"_cn","_ru")
	Agg_pt = Replace(Agg_cn,"_cn","_pt")
	Agg_de = Replace(Agg_cn,"_cn","_de")

	Aggiornamento__B2B__157 = _
		DropObject(conn,"gv_articoli_it","VIEW") + _
		DropObject(conn,"gv_articoli_en","VIEW") + _
		DropObject(conn,"gv_articoli_fr","VIEW") + _
		DropObject(conn,"gv_articoli_de","VIEW") + _
		DropObject(conn,"gv_articoli_es","VIEW") + _
		DropObject(conn,"gv_articoli_ru","VIEW") + _
		DropObject(conn,"gv_articoli_pt","VIEW") + _
		DropObject(conn,"gv_articoli_cn","VIEW") + _
		Agg_it + Agg_en  + Agg_fr + Agg_de + Agg_es + Agg_ru + Agg_pt + Agg_cn
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 158
'..........................................................................................................................
'Matteo - 27/05/2011
'..........................................................................................................................
' modifica funzioni fn_listino_vendita_articoli: accorcia le righe per permettere il rebuild
'..........................................................................................................................
function Aggiornamento__B2B__158(conn)
		if cIntero(DB_SQL_version(conn)) >= 9 then
			Aggiornamento__B2B__158 = _
			DropObject(conn, "fn_listino_vendita_articoli", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"				 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"				 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iav, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"				 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid  " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				    AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"					AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) "
		else
			Aggiornamento__B2B__158 = ""
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 159
'...........................................................................................
'	Giacomo, 14/07/2011
'...........................................................................................
'   aggiungo ulteriore campo immagine a gtb_marche
'...........................................................................................
function Aggiornamento__B2B__159(conn)
	Aggiornamento__B2B__159 = _	
		" ALTER TABLE gtb_marche ADD " + _
		"	mar_img " + SQL_CharField(Conn, 510) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 160
'...........................................................................................
'	Nicola, 28/09/2011
'...........................................................................................
'   aggiungo campo per ritorno dati sull'immagine principale.
'...........................................................................................
function Aggiornamento__B2B__160(conn)
	Aggiornamento__B2B__160 = _	
		" ALTER TABLE gtb_articoli ADD " + _
		"	art_foto_thumb " + SQL_CharField(Conn, 250) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 161
'..........................................................................................................................
'Nicola - 06/10/2011
'..........................................................................................................................
' ricostruisce funzioni sql per restituire il listino in vigore per gli articoli (prezzo minimo)
' c'era un problema di lunghezza di riga nella "rigenerazione" della funzione
'..........................................................................................................................
function Aggiornamento__B2B__161(conn)
		if cIntero(DB_SQL_version(conn)) >= 9 then
			Aggiornamento__B2B__161 = _
			DropObject(conn, "fn_listino_vendita_articoli", "FUNCTION") + vbcrLf + _
			DropObject(conn, "fn_listino_vendita_varianti", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iav, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid  " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) " + _
			" ; " + _
			" CREATE FUNCTION dbo.fn_listino_vendita_varianti( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT gv_articoli.*, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo) AS prezzo, " + vbCrLF + _
			" 		   COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore) AS iva, " + vbCrLF + _
			" 		   COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id) AS ivaId, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_scontoQ_id, cliente.prz_scontoQ_id, base.prz_scontoQ_id, rel_scontoQ_id) AS scontoQId, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0) AS listinoId, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) AS visibile, " + vbCrLF + _
			" 		   COALESCE(offerte.prz_non_vendibile, cliente.prz_non_vendibile, base.prz_non_vendibile, 0) AS nonVendibile " + vbCrLF + _
			"	FROM gv_articoli " + vbCrLF + _
			" 		INNER JOIN gtb_prezzi base ON gv_articoli.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 		LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 		LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 						AND ISNULL(offerte.prz_visibile, 0) = 1 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 						AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 		LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id " + vbCrLF + _
			" 		LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 		LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"	WHERE ISNULL(art_disabilitato,0) = 0 " + vbCrLf + _
			"		AND ISNULL(rel_disabilitato,0) = 0 " + vbCrLF + _
			"		AND tip_visibile = 1 " + vbCrLf + _
			"		AND tip_albero_visibile = 1 " + vbCrLF + _
			"		AND COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) = 1 " + vbCrLF + _
			" ) "
		else
			Aggiornamento__B2B__161 = ""
		end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 162
'...........................................................................................
'	Giacomo, 27/10/2011
'...........................................................................................
'   aggiungo campi per gestione giacenze e per gestione spese di spedizione
'...........................................................................................
function Aggiornamento__B2B__162(conn)
	Aggiornamento__B2B__162 = _	
		" ALTER TABLE grel_giacenze ADD " + _
		"	gia_ordinato_data_arrivo smalldatetime NULL; " + _
		" ALTER TABLE gtb_rivenditori ADD " + _
		"	riv_spese_spedizione_id int NULL; " + _
		" ALTER TABLE gtb_spese_spedizione ADD " + _
		"	sp_percentuale real NULL; " + _
		SQL_AddForeignKey(conn, "gtb_rivenditori", "riv_spese_spedizione_id", "gtb_spese_spedizione", "sp_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 163
'...........................................................................................
' Giacomo 13/12/2011
'...........................................................................................
' aggiunge parametro
'...........................................................................................
function Aggiornamento__B2B__163(conn)
	Aggiornamento__B2B__163 = "SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__163(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "B2B_ID_PAG_SPEDIZ_CREDENZ_ACCESSO", _
									0, _
									"pagina che viene spedita ad un contatto, contenente le credenziali di accesso.", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)						
		AggiornamentoSpeciale__B2B__163 = " SELECT * FROM AA_Versione "
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 164
'...........................................................................................
'	Giacomo, 16/12/2011
'...........................................................................................
'   creo tabelle tipologie porti e trasportatori, e relative chiavi esterne su ordini e rivenditori
'...........................................................................................
function Aggiornamento__B2B__164(conn)
	Aggiornamento__B2B__164 = _	
		" CREATE TABLE " + SQL_Dbo(conn) + "gtb_porti(" + _
		"	prt_id " + SQL_PrimaryKey(conn, "gtb_porti") + ", " + _
		SQL_MultiLanguageFieldComplete(conn, "prt_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "," + _
		" 	prt_con_spese bit NULL, " + _
		" 	prt_con_vettore bit NULL " + _
		"); " + _
		" CREATE TABLE " + SQL_Dbo(conn) + "gtb_trasportatori(" + _
		"	tra_id " + SQL_PrimaryKey(conn, "gtb_trasportatori") + ", " + _
		SQL_MultiLanguageFieldComplete(conn, "tra_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "," + _
		" 	tra_codice " + SQL_CharField(Conn, 255) + " NULL, " + _
		SQL_MultiLanguageFieldComplete(conn, "tra_descrizione_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
		"); " + _
		" ALTER TABLE gtb_rivenditori ADD " + _
		"	riv_porto_default_id int NULL; " + _
		" ALTER TABLE gtb_rivenditori ADD " + _
		"	riv_trasportatore_default_id int NULL; " + _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_porto_id int NULL; " + _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_trasportatore_id int NULL; " + _
		SQL_AddForeignKey(conn, "gtb_rivenditori", "riv_porto_default_id", "gtb_porti", "prt_id", false, "") + _
		SQL_AddForeignKey(conn, "gtb_rivenditori", "riv_trasportatore_default_id", "gtb_trasportatori", "tra_id", false, "") + _
		SQL_AddForeignKey(conn, "gtb_ordini", "ord_porto_id", "gtb_porti", "prt_id", false, "") + _
		SQL_AddForeignKey(conn, "gtb_ordini", "ord_trasportatore_id", "gtb_trasportatori", "tra_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 165
'...........................................................................................
'	Giacomo, 20/12/2011
'...........................................................................................
'   aggiungo colonne dati inserimento e modifica su magazzini e listini
'...........................................................................................
function Aggiornamento__B2B__165(conn)
	Aggiornamento__B2B__165 = _	
		" ALTER TABLE gtb_magazzini ADD " + _
	    "   mag_insData	datetime NULL, " + _
	    "   mag_insAdmin_id	int	NULL, " + _
	    "   mag_modData	datetime NULL, " + _
	    "   mag_modAdmin_id	int NULL ; " + _
	    " ALTER TABLE gtb_listini ADD " + _
	    "   listino_insData	datetime NULL, " + _
	    "   listino_insAdmin_id	int	NULL, " + _
	    "   listino_modData	datetime NULL, " + _
	    "   listino_modAdmin_id	int NULL ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 166
'...........................................................................................
'	Giacomo, 29/12/2011
'...........................................................................................
'   aggiungo colonne per chiavi esterne per trasportatore e porto sulla shopping cart
'...........................................................................................
function Aggiornamento__B2B__166(conn)
	Aggiornamento__B2B__166 = _	
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_porto_id int NULL; " + _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_trasportatore_id int NULL; " + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_porto_id", "gtb_porti", "prt_id", false, "") + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_trasportatore_id", "gtb_trasportatori", "tra_id", false, "")	
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 167
'...........................................................................................
'	Giacomo, 10/01/2012
'...........................................................................................
'   aggiungo campo nome ai listini
'...........................................................................................
function Aggiornamento__B2B__167(conn)
	Aggiornamento__B2B__167 = _	
		" ALTER TABLE gtb_listini ADD " + _
		"	listino_nome_it " + SQL_CharField(Conn, 500) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 168
'...........................................................................................
'	Giacomo, 12/01/2012
'...........................................................................................
'   aggiungo campi
'...........................................................................................
function Aggiornamento__B2B__168(conn)
	Aggiornamento__B2B__168 = _	
		" ALTER TABLE gtb_porti ADD " + _
		"	prt_scelta_sede bit NULL; " + _
		" ALTER TABLE gtb_listini ADD " + _
		"	listino_importato bit NULL, " + _
		"	listino_dataImport datetime NULL; "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 169
'...........................................................................................
' Giacomo 13/01/2012
'...........................................................................................
' aggiunge parametro per gestire la dipendenza tra i prezzi dei listini
'...........................................................................................
function Aggiornamento__B2B__169(conn)
	Aggiornamento__B2B__169 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__169(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "LISTINI_PREZZI_INDIPENDENTI", _
									0, _
									"Vengono disattivati i ricalcoli automatici dei prezzi dei listini (vengono ignorate le variazioni).", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 170
'...........................................................................................
'	Giacomo, 23/01/2012
'...........................................................................................
'   modifica vista gv_listino_offerte - tolto INNER JOIN con spese_spedizione_articolo perche' rallentava la vista
'...........................................................................................
function Aggiornamento__B2B__170(conn)
	Aggiornamento__B2B__170 = _	
		DropObject(conn, "gv_listino_offerte", "VIEW") + _
		" CREATE VIEW dbo.gv_listino_offerte AS " + vbCrLf + _
		"    SELECT * FROM gtb_articoli " + vbCrLF + _
		"        INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id " + vbCrLF + _
		"        INNER JOIN gtb_prezzi ON grel_art_valori.rel_id = gtb_prezzi.prz_variante_id " + vbCrLF + _
		"        INNER JOIN gtb_listini ON gtb_prezzi.prz_listino_id = gtb_listini.listino_id " + vbCrLF + _
		"        INNER JOIN gtb_iva ON gtb_prezzi.prz_iva_id = gtb_iva.iva_id " + vbCrLF + _
		"        INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id " + vbCrLF + _
		"    WHERE ISNULL(gtb_articoli.art_disabilitato, 0) = 0 " + vbCrLF + _
		"          AND ISNULL(grel_art_valori.rel_disabilitato, 0)=0 " + vbCrLF + _
		"          AND tip_visibile=1 " + vbCrLF + _
		"          AND tip_albero_visibile=1 " + vbCrLF + _
		"          AND ISNULL(listino_offerte, 0)=1 " + vbCrLF + _
		"          AND ISNULL(prz_visibile, 0)=1 " + vbCrLF + _
		"          AND ( " + vbCrLF + _
		"               ( GETDATE() BETWEEN listino_dataCreazione AND ISNULL(listino_dataScadenza, GETDATE())+1 ) " + vbCrLF + _
		"               OR " + vbCrLF + _
		"               ( listino_dataCreazione IS NULL AND " + vbCrLF + _
		"                 listino_dataScadenza IS NULL AND " + vbCrLF + _
		"                 GETDATE() BETWEEN ISNULL(prz_offerta_dal, GETDATE()-1) AND ISNULL(prz_offerta_al, GETDATE())+1 " + vbCrLF + _
		"               )  " + vbCrLF + _
		"              ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 171
'...........................................................................................
'	Giacomo, 30/01/2012
'...........................................................................................
'   aggiungo campi
'...........................................................................................
function Aggiornamento__B2B__171(conn)
	Aggiornamento__B2B__171 = _	
		" ALTER TABLE gtb_porti ADD " + _
		"	prt_attivo bit NULL, " + _
		SQL_MultiLanguageFieldComplete(conn, "prt_descrizione_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + ";" + _
		" ALTER TABLE gtb_trasportatori ADD " + _
		"	tra_attivo bit NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 172
'...........................................................................................
'	Giacomo, 01/02/2012
'...........................................................................................
'   aggiungo campo
'...........................................................................................
function Aggiornamento__B2B__172(conn)
	Aggiornamento__B2B__172 = _	
		" ALTER TABLE gtb_porti ADD " + _
		"	prt_codice " + SQL_CharField(Conn, 500) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 173
'...........................................................................................
'	Giacomo, 02/02/2012
'...........................................................................................
'   aggiungo campi
'...........................................................................................
function Aggiornamento__B2B__173(conn)
	Aggiornamento__B2B__173 = _	
		" ALTER TABLE gtb_dettagli_ord ADD " + _
		"	det_qta_evasa int NULL," + _
		"	det_data_evasione smalldatetime NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 174
'..........................................................................................................................
'Nicola - 10/02/2012
'..........................................................................................................................
' modifica vista listino per rimozione bug gestione sconti quantità
'..........................................................................................................................
function Aggiornamento__B2B__174(conn)
		Aggiornamento__B2B__174 = _
			DropObject(conn, "fn_listino_vendita_varianti", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_varianti( " + vbCrLF + _
			"   @listinoBaseId int, " + vbCrLF + _
			"   @listinoOfferteId int, " + vbCrLF + _
			"   @listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			"   SELECT gv_articoli.*, " + vbCrLF + _
			"          COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo) AS prezzo, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore) AS iva, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id) AS ivaId, " + vbCrLF + _
			"          (CASE WHEN offerte.prz_id IS NOT null THEN IsNull(offerte.prz_scontoQ_id, 0) " + vbCrLf + _
			"                WHEN cliente.prz_id IS NOT null THEN IsNull(cliente.prz_scontoQ_id,0) " + vbCrLf + _
			"                ELSE IsNull(base.prz_scontoQ_id,0) " + vbCrLF + _
			"                END) AS scontoQId, " + vbCrLF + _
			"          COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0) AS listinoId, " + vbCrLF + _
			"          COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) AS visibile, " + vbCrLF + _
			"          COALESCE(offerte.prz_non_vendibile, cliente.prz_non_vendibile, base.prz_non_vendibile, 0) AS nonVendibile " + vbCrLF + _
			"   FROM gv_articoli " + vbCrLF + _
			"        INNER JOIN gtb_prezzi base ON gv_articoli.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			"        LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"                                        AND ISNULL(offerte.prz_visibile, 0) = 1 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"                                        AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			"        LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			"        LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"   WHERE ISNULL(art_disabilitato,0) = 0 " + vbCrLf + _
			"         AND ISNULL(rel_disabilitato,0) = 0 " + vbCrLF + _
			"         AND tip_visibile = 1 " + vbCrLf + _
			"         AND tip_albero_visibile = 1 " + vbCrLF + _
			"         AND COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) = 1 " + vbCrLF + _
			" ) "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 175
'..........................................................................................................................
'Giacomo - 14/02/2012
'..........................................................................................................................
' modifica viste listino per rimozione bug gestione "in offerta"
'..........................................................................................................................
function Aggiornamento__B2B__175(conn)
		Aggiornamento__B2B__175 = _
			DropObject(conn, "fn_listino_vendita_varianti", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_varianti( " + vbCrLF + _
			"   @listinoBaseId int, " + vbCrLF + _
			"   @listinoOfferteId int, " + vbCrLF + _
			"   @listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			"   SELECT gv_articoli.*, " + vbCrLF + _
			"          COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo) AS prezzo, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore) AS iva, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id) AS ivaId, " + vbCrLF + _
			"          (CASE WHEN offerte.prz_id IS NOT null THEN IsNull(offerte.prz_scontoQ_id, 0) " + vbCrLf + _
			"                WHEN cliente.prz_id IS NOT null THEN IsNull(cliente.prz_scontoQ_id,0) " + vbCrLf + _
			"                ELSE IsNull(base.prz_scontoQ_id,0) " + vbCrLF + _
			"                END) AS scontoQId, " + vbCrLF + _
			"          COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0) AS listinoId, " + vbCrLF + _
			"          COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) AS visibile, " + vbCrLF + _
			"          COALESCE(offerte.prz_non_vendibile, cliente.prz_non_vendibile, base.prz_non_vendibile, 0) AS nonVendibile " + vbCrLF + _
			"   FROM gv_articoli " + vbCrLF + _
			"        INNER JOIN gtb_prezzi base ON gv_articoli.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			"        LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"                                        AND ISNULL(offerte.prz_visibile, 0) = 1  AND " + vbCrLF + _
			"										(GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) ) " + vbCrLF + _
			"        LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			"        LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"   WHERE ISNULL(art_disabilitato,0) = 0 " + vbCrLf + _
			"         AND ISNULL(rel_disabilitato,0) = 0 " + vbCrLF + _
			"         AND tip_visibile = 1 " + vbCrLf + _
			"         AND tip_albero_visibile = 1 " + vbCrLF + _
			"         AND COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) = 1 " + vbCrLF + _
			" ); " + vbCrLF + _
			DropObject(conn, "fn_listino_vendita_articoli", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iva, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid, " + vbCrLF + _
			"		   (SELECT (CASE WHEN ISNULL(offerte.prz_id, 0) = 0 THEN 0 ELSE 1 END) " + vbCrLF + _
			"			FROM grel_art_valori LEFT JOIN gtb_prezzi offerte ON grel_art_valori.rel_id = offerte.prz_variante_id " + vbCrLF + _
			"				AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"				AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"				AND (GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) ) " + vbCrLF + _
            "			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS in_offerta " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) "			
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 176
'..........................................................................................................................
'Giacomo - 16/02/2012
'..........................................................................................................................
' modifica vista listino vendita articoli (correzione colonna in_offerta)
'..........................................................................................................................
function Aggiornamento__B2B__176(conn)
		Aggiornamento__B2B__176 = _
			DropObject(conn, "fn_listino_vendita_articoli", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iva, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid, " + vbCrLF + _
			"		   (SELECT TOP 1 (CASE WHEN ISNULL(offerte.prz_id, 0) = 0 THEN 0 ELSE 1 END) " + vbCrLF + _
			"			FROM grel_art_valori LEFT JOIN gtb_prezzi offerte ON grel_art_valori.rel_id = offerte.prz_variante_id " + vbCrLF + _
			"				AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"				AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"				AND (GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) ) " + vbCrLF + _
            "			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ORDER BY offerte.prz_id DESC) AS in_offerta " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) "			
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 177
'...........................................................................................
'	Giacomo, 17/02/2012
'...........................................................................................
'   aggiungo campo
'...........................................................................................
function Aggiornamento__B2B__177(conn)
	Aggiornamento__B2B__177 = _	
		" ALTER TABLE gtb_modipagamento ADD " + _
		"	mosp_codice " + SQL_CharField(Conn, 500) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 178
'...........................................................................................
'	Giacomo, 21/02/2012
'...........................................................................................
'   aggiungo campo
'...........................................................................................
function Aggiornamento__B2B__178(conn)
	Aggiornamento__B2B__178 = _	
		" ALTER TABLE gtb_spese_spedizione ADD " + _
		"	sp_codice " + SQL_CharField(Conn, 500) + " NULL;"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 179
'..........................................................................................................................
'Nicola - 23/02/2012
'..........................................................................................................................
' modifica viste listino aggiunta colonna "In_offerta"
'..........................................................................................................................
function Aggiornamento__B2B__179(conn)
		Aggiornamento__B2B__179 = _
			DropObject(conn, "fn_listino_vendita_varianti", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_varianti( " + vbCrLF + _
			"   @listinoBaseId int, " + vbCrLF + _
			"   @listinoOfferteId int, " + vbCrLF + _
			"   @listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			"   SELECT gv_articoli.*, " + vbCrLF + _
			"          COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo) AS prezzo, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore) AS iva, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id) AS ivaId, " + vbCrLF + _
			"          (CASE WHEN offerte.prz_id IS NOT null THEN IsNull(offerte.prz_scontoQ_id, 0) " + vbCrLf + _
			"                WHEN cliente.prz_id IS NOT null THEN IsNull(cliente.prz_scontoQ_id,0) " + vbCrLf + _
			"                ELSE IsNull(base.prz_scontoQ_id,0) " + vbCrLF + _
			"                END) AS scontoQId, " + vbCrLF + _
			"          COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0) AS listinoId, " + vbCrLF + _
			"          COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) AS visibile, " + vbCrLF + _
			"          COALESCE(offerte.prz_non_vendibile, cliente.prz_non_vendibile, base.prz_non_vendibile, 0) AS nonVendibile, " + vbCrLF + _
			"          (CASE WHEN COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0)=@listinoOfferteId THEN 1 ELSE 0 END) AS in_offerta " + vbCrLf + _
			"   FROM gv_articoli " + vbCrLF + _
			"        INNER JOIN gtb_prezzi base ON gv_articoli.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			"        LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"                AND ISNULL(offerte.prz_visibile, 0) = 1  AND " + vbCrLF + _
			"				(GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) ) " + vbCrLF + _
			"        LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			"        LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"   WHERE ISNULL(art_disabilitato,0) = 0 " + vbCrLf + _
			"         AND ISNULL(rel_disabilitato,0) = 0 " + vbCrLF + _
			"         AND tip_visibile = 1 " + vbCrLf + _
			"         AND tip_albero_visibile = 1 " + vbCrLF + _
			"         AND COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) = 1 " + vbCrLF + _
			" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 180
'...........................................................................................
'	Giacomo, 27/02/2012
'...........................................................................................
'   aggiungo campo e aggiungo parametro
'...........................................................................................
function Aggiornamento__B2B__180(conn)
	Aggiornamento__B2B__180 = _	
		" ALTER TABLE gtb_rivenditori ADD " + _
		"	riv_sconto_ordine real NULL; "
end function
	
function AggiornamentoSpeciale__B2B__180(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "ID_CATEGORIA_BASE_CATALOGO", _
									0, _
									"Id della categoria base per gli articoli del catalogo in vendita.", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 181
'...........................................................................................
'	Nicola, 27/02/2012
'...........................................................................................
'   aggiungo campi su shopping cart ed ordine per gestione sconti piede + revisione trigger 
'	e stored procedure
'...........................................................................................
function Aggiornamento__B2B__181(conn)
	Aggiornamento__B2B__181 = _	
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_sconto_cliente money NULL, " + _
		"	sc_sconto_web money NULL; " + _
		" ALTER TABLE gtb_ordini ADD " + _
		" 	ord_sconto_cliente money NULL, " + _
		"	ord_sconto_web money NULL; "
end function
	
function AggiornamentoSpeciale__B2B__181(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "PERCENTUALE_SCONTO_WEB", _
									0, _
									"Eventuale sconto web applicato a tutti gli ordini.", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 182
'...........................................................................................
'	Giacomo, 22/03/2012
'...........................................................................................
'   aggiungo campo su gtb_ordini
'...........................................................................................
function Aggiornamento__B2B__182(conn)
	Aggiornamento__B2B__182 = _	
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_exported bit NULL, " + _
		" 	ord_export_data smalldatetime NULL, " + _
		"	ord_imported bit NULL, " + _
		" 	ord_import_data smalldatetime NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 183
'...........................................................................................
'	Giacomo, 22/03/2012
'...........................................................................................
'   aggiungo campo id_utente
'...........................................................................................
function Aggiornamento__B2B__183(conn)
	Aggiornamento__B2B__183 = _	
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_ut_id INT NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 184
'...........................................................................................
' Giacomo 02/04/2012
'...........................................................................................
' aggiunge parametro per gestire lo sconto a piede dell'ordine
'...........................................................................................
function Aggiornamento__B2B__184(conn)
	Aggiornamento__B2B__184 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__184(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "SCONTO_PIEDE_ESCLUDI_OFFERTE", _
									0, _
									"Lo sconto a piede dell'ordine non viene applicato agli articoli in offerta.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 185
'..........................................................................................................................
'Giacomo - 04/04/2012
'..........................................................................................................................
' modifica vista listino vendita articoli (aggiunta colonna in_promozione)
'..........................................................................................................................
function Aggiornamento__B2B__185(conn)
		Aggiornamento__B2B__185 = _
			DropObject(conn, "fn_listino_vendita_articoli", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_articoli( " + vbCrLF + _
			" 	@listinoBaseId int, " + vbCrLF + _
			"	@listinoOfferteId int, " + vbCrLF + _
			"	@listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			" 	SELECT *,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 						AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 						AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS prezzo,  " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 						AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 						AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			" 			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS iva, " + vbCrLF + _
			" 		   (SELECT MIN(COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id)) " + vbCrLF + _
			" 			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 						AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 						AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id  " + vbCrLF + _
			" 				 LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			" 				 LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS ivaid, " + vbCrLF + _
			"		   (SELECT TOP 1 (CASE WHEN ISNULL(offerte.prz_id, 0) = 0 THEN 0 ELSE 1 END) " + vbCrLF + _
			"			FROM grel_art_valori LEFT JOIN gtb_prezzi offerte ON grel_art_valori.rel_id = offerte.prz_variante_id " + vbCrLF + _
			"				AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"				AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"				AND (GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) ) " + vbCrLF + _
            "			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ORDER BY offerte.prz_id DESC) AS in_offerta, " + vbCrLF + _
			"		   (SELECT COALESCE(offerte.prz_promozione, cliente.prz_promozione, base.prz_promozione) " + vbCrLF + _
			"			FROM grel_art_valori INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			"			LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"	 								 AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			"	 								 AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) " + vbCrLF + _
			"	 		LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			"			WHERE grel_art_valori.rel_art_id = gtb_articoli.art_id ) AS in_promozione " + vbCrLF + _
			" 	FROM gtb_articoli  " + vbCrLF + _
			" 		INNER JOIN gtb_marche ON gtb_articoli.art_marca_id = gtb_marche.mar_id  " + vbCrLF + _
			" 		INNER JOIN gtb_tipologie ON gtb_articoli.art_tipologia_id = gtb_tipologie.tip_id  " + vbCrLF + _
			" 	WHERE ISNULL(gtb_articoli.art_disabilitato,0) = 0 " + vbCrLF + _
			" 		AND tip_visibile = 1 " + vbCrLF + _
			" 		AND tip_albero_visibile = 1  " + vbCrLF + _
			" 		AND (SELECT MAX(COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END))  " + vbCrLF + _
			" 				FROM grel_art_valori " + vbCrLF + _
			" 					INNER JOIN gtb_prezzi base ON grel_art_valori.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			" 				 					AND ISNULL(offerte.prz_visibile, 0) = 1 " + vbCrLF + _
			" 				 					AND GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1)  " + vbCrLF + _
			" 					LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId  " + vbCrLF + _
			" 				WHERE rel_art_id = gtb_articoli.art_id) = 1  " + vbCrLF + _
			" ) "			
end function
'*******************************************************************************************

'*******************************************************************************************
' AGGIORNAMENTO B2B 186
'..........................................................................................................................
' Giacomo - 04/04/2012
'..........................................................................................................................
' modifica viste listino (aggiunta colonna in_promozione)
'..........................................................................................................................
function Aggiornamento__B2B__186(conn)
		Aggiornamento__B2B__186 = _
			DropObject(conn, "fn_listino_vendita_varianti", "FUNCTION") + vbcrLf + _
			" CREATE FUNCTION dbo.fn_listino_vendita_varianti( " + vbCrLF + _
			"   @listinoBaseId int, " + vbCrLF + _
			"   @listinoOfferteId int, " + vbCrLF + _
			"   @listinoClienteId int " + vbCrLF + _
			" ) " + vbCrLF + _
			" RETURNS TABLE AS " + vbCrLF + _
			" RETURN (  " + vbCrLF + _
			"   SELECT gv_articoli.*, " + vbCrLF + _
			"          COALESCE(offerte.prz_prezzo, cliente.prz_prezzo, base.prz_prezzo, rel_prezzo) AS prezzo, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_valore, clienteIva.iva_valore, baseIva.iva_valore) AS iva, " + vbCrLF + _
			"          COALESCE(offerteIva.iva_id, clienteIva.iva_id, baseIva.iva_id) AS ivaId, " + vbCrLF + _
			"          (CASE WHEN offerte.prz_id IS NOT null THEN IsNull(offerte.prz_scontoQ_id, 0) " + vbCrLf + _
			"                WHEN cliente.prz_id IS NOT null THEN IsNull(cliente.prz_scontoQ_id,0) " + vbCrLf + _
			"                ELSE IsNull(base.prz_scontoQ_id,0) " + vbCrLF + _
			"                END) AS scontoQId, " + vbCrLF + _
			"          COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0) AS listinoId, " + vbCrLF + _
			"          COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) AS visibile, " + vbCrLF + _
			"          COALESCE(offerte.prz_non_vendibile, cliente.prz_non_vendibile, base.prz_non_vendibile, 0) AS nonVendibile, " + vbCrLF + _
			"          (CASE WHEN COALESCE(offerte.prz_listino_id, cliente.prz_listino_id, base.prz_listino_id, 0)=@listinoOfferteId THEN 1 ELSE 0 END) AS in_offerta, " + vbCrLf + _
			"          COALESCE(offerte.prz_promozione, cliente.prz_promozione, base.prz_promozione) AS in_promozione " + vbCrLF + _
			"   FROM gv_articoli " + vbCrLF + _
			"        INNER JOIN gtb_prezzi base ON gv_articoli.rel_id = base.prz_variante_id AND base.prz_listino_id = @listinoBaseId " + vbCrLF + _
			"        LEFT JOIN gtb_iva baseIva ON base.prz_iva_id = baseIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi offerte ON base.prz_variante_id = offerte.prz_variante_id AND offerte.prz_listino_id = @listinoOfferteId " + vbCrLF + _
			"                AND ISNULL(offerte.prz_visibile, 0) = 1  AND " + vbCrLF + _
			"				(GETDATE() BETWEEN IsNull(offerte.prz_offerta_dal, GETDATE()-1) AND IsNull(offerte.prz_offerta_al, GETDATE() + 1) ) " + vbCrLF + _
			"        LEFT JOIN gtb_iva offerteIva ON offerte.prz_iva_id = offerteIva.iva_id " + vbCrLF + _
			"        LEFT JOIN gtb_prezzi cliente ON base.prz_variante_id = cliente.prz_variante_id AND cliente.prz_listino_id = @listinoClienteId " + vbCrLF + _
			"        LEFT JOIN gtb_iva clienteIva ON cliente.prz_iva_id = clienteIva.iva_id " + vbCrLF + _
			"   WHERE ISNULL(art_disabilitato,0) = 0 " + vbCrLf + _
			"         AND ISNULL(rel_disabilitato,0) = 0 " + vbCrLF + _
			"         AND tip_visibile = 1 " + vbCrLf + _
			"         AND tip_albero_visibile = 1 " + vbCrLF + _
			"         AND COALESCE(offerte.prz_visibile, cliente.prz_visibile, base.prz_visibile, CASE WHEN rel_disabilitato = 1 THEN 0 ELSE 1 END) = 1 " + vbCrLF + _
			" ); "
end function
'*******************************************************************************************


'*******************************************************************************************
' AGGIORNAMENTO B2B 187
'...........................................................................................
' Nicola 13/04/2012
'...........................................................................................
' aggiunge iva alle spese di spedizione
'...........................................................................................
function Aggiornamento__B2B__187(conn)
	Aggiornamento__B2B__187 = _
		" ALTER TABLE gtb_spese_spedizione ADD sp_iva_id INT NULL ; " + _
		SQL_AddForeignKey(conn, "gtb_spese_spedizione", "sp_iva_id", "gtb_iva", "iva_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
' AGGIORNAMENTO B2B 188
'...........................................................................................
' Giacomo 18/06/2012
'...........................................................................................
' ricrea vista dettaglio ordini
'...........................................................................................
function Aggiornamento__B2B__188(conn)
	Aggiornamento__B2B__188 = _
		DropObject(conn, "gv_dettagli_ord","VIEW") + _
		" CREATE VIEW dbo.gv_dettagli_ord AS " + vbCrLF + _
		"	SELECT * FROM gtb_dettagli_ord " + vbCrLF + _
		"		LEFT JOIN grel_art_valori ON gtb_dettagli_ord.det_art_var_id = grel_art_valori.rel_id " + vbCrLF + _
		"		LEFT JOIN gtb_articoli ON grel_art_valori.rel_art_id = gtb_articoli.art_id " + vbCrLF + _
		"		LEFT JOIN gtb_spese_spedizione_articolo ON gtb_articoli.art_spedizione_id = gtb_spese_spedizione_articolo.spa_id " + vbCrLf + _
		";"
end function
'*******************************************************************************************


'*******************************************************************************************
' AGGIORNAMENTO B2B 189
'...........................................................................................
' Giacomo 18/06/2012
'...........................................................................................
' aggiunge campo visibilità sul rivenditore
'...........................................................................................
function Aggiornamento__B2B__189(conn)
	Aggiornamento__B2B__189 = _
		" ALTER TABLE gtb_rivenditori ADD riv_attivo bit NULL; " + _
		" DROP VIEW dbo.gv_rivenditori ;" + vbCrLf + _
		" CREATE VIEW dbo.gv_rivenditori AS " + vbCrLf + _
		"     SELECT riv_id, riv_listino_id, riv_lstcod_id, riv_valuta_id, riv_agente_id, riv_codice, riv_modopagamento_id, riv_profilo_id, " + vbCrLf + _
		"            riv_azienda_capogruppo_id, riv_spese_spedizione_id, riv_porto_default_id, riv_trasportatore_default_id, riv_sconto_ordine, " + vbCrLf + _
		"            ISNULL(riv_attivo, 1) AS riv_attivo,  " + vbCrLf + _
		"			 tb_utenti.*, tb_indirizzario.*, gtb_valute.* " + vbCrLf + _
		"		FROM gtb_rivenditori " + vbCrLf + _
		"		INNER JOIN tb_Utenti ON gtb_rivenditori.riv_id = tb_utenti.ut_ID " + vbCrLf + _
		"		INNER JOIN tb_Indirizzario ON tb_utenti.ut_NextCom_ID = tb_indirizzario.IDElencoIndirizzi " + vbCrLf + _
		"		INNER JOIN gtb_valute ON gtb_rivenditori.riv_valuta_id = gtb_valute.valu_id"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 190
'...........................................................................................
' Nicola 17/09/2012
'...........................................................................................
' aggiunge parametro per gestire l'attivazione delle righe multiple in ordine per lo stesso articolo 
'...........................................................................................
function Aggiornamento__B2B__190(conn)
	Aggiornamento__B2B__190 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__190(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "ORDINE_RIGHE_ARTICOLO_MULTIPLE", _
									0, _
									"Attiva la gestione delle righe d'ordine multiple dello stesso articolo.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 191
'...........................................................................................
' Giacomo 21/09/2012
'...........................................................................................
' aggiunge parametro per gestire la possibilità di inserire articoli con il prezzo a zero
'...........................................................................................
function Aggiornamento__B2B__191(conn)
	Aggiornamento__B2B__191 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__191(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "INIBISCI_PREZZO_A_ZERO", _
									0, _
									"Rende impossibile inserire un articolo con prezzo a zero.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 192
'...........................................................................................
'	Giacomo, 06/03/2013
'...........................................................................................
'   creo tabella ordini tipo consegna
'...........................................................................................
function Aggiornamento__B2B__192(conn)
	Aggiornamento__B2B__192 = _	
		" CREATE TABLE " + SQL_Dbo(conn) + "gtb_tipo_consegna(" + _
		"	tco_id " + SQL_PrimaryKey(conn, "gtb_tipo_consegna") + ", " + _
		" 	tco_ordine int NULL, " + _
		SQL_MultiLanguageFieldComplete(conn, "tco_nome_<lingua> " + SQL_CharField(Conn, 255) + " NULL ") + "," + _
		SQL_MultiLanguageFieldComplete(conn, "tco_descrizione_<lingua> " + SQL_CharField(Conn, 0) + " NULL ") + _
		"); " + _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_tipo_consegna_id int NULL; " + _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_tipo_consegna_id int NULL; " + _
		" ALTER TABLE gtb_porti ADD " + _
		"	prt_scelta_modalita_consegna bit NULL; " + _
		SQL_AddForeignKey(conn, "gtb_ordini", "ord_tipo_consegna_id", "gtb_tipo_consegna", "tco_id", false, "") + _
		SQL_AddForeignKey(conn, "gtb_shopping_cart", "sc_tipo_consegna_id", "gtb_tipo_consegna", "tco_id", false, "")
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 193
'...........................................................................................
' Giacomo 10/03/2014
'...........................................................................................
' aggiunge parametro per gestire la possibilità di avere la descrizione modificabile con CKEditor
'...........................................................................................
function Aggiornamento__B2B__193(conn)
	Aggiornamento__B2B__193 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__193(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "B2B_ABILITA_DESCRIZIONE_HTML", _
									0, _
									"Attiva CKEditor per il campo descrizione estesa degli articoli.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 194
'...........................................................................................
'	Nicola, 25/03/2013
'...........................................................................................
'   aggiungo riferimento ordine su dettaglio shopping cart
'...........................................................................................
function Aggiornamento__B2B__194(conn)
	Aggiornamento__B2B__194 = _	
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_ord_origine_id int NULL; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 195
'...........................................................................................
'	Nicola, 29/09/2014
'...........................................................................................
'   aggiungo campi shopping cart ed ordine per B2B acquisti Ideallux e dati per preventivo
'...........................................................................................
function Aggiornamento__B2B__195(conn)
	Aggiornamento__B2B__195 = _	
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_is_preventivo bit NULL, " + _
		" 	sc_codice nvarchar(100) NULL; " + _
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_is_preventivo bit NULL"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 196
'...........................................................................................
'aggiunge trigger in update per i prezzi dei listini derivati
'...........................................................................................
function Aggiornamento__B2B__196(conn)
	Aggiornamento__B2B__196 = _ 
		" IF  EXISTS (SELECT * FROM sys.triggers WHERE object_id = OBJECT_ID(N'[dbo].[TRG_gtb_prezzi_FOR_UPDATE]')) " & vbCrLF + _
		" 		DROP TRIGGER [dbo].[TRG_gtb_prezzi_FOR_UPDATE] " & vbCrLF + _
		" ; " + _
		" CREATE TRIGGER [dbo].[TRG_gtb_prezzi_FOR_UPDATE] " & vbCrlf & _
		"  	ON [dbo].[gtb_prezzi] " & vbCrlf & _
		"  	AFTER UPDATE " & vbCrlf & _
 		" AS " & vbCrlf & _
 		"  	UPDATE gtb_prezzi " & vbCrlf & _
 		"  	SET " & vbCrlf & _
 		"  		gtb_prezzi.prz_prezzo = INS.prz_prezzo, " & vbCrlf & _
 		"  		gtb_prezzi.prz_visibile = INS.prz_visibile, " & vbCrlf & _
 		"  		gtb_prezzi.prz_promozione = INS.prz_promozione, " & vbCrlf & _
 		"  		gtb_prezzi.prz_variante_id = INS.prz_variante_id, " & vbCrlf & _
 		"  		gtb_prezzi.prz_scontoQ_id = INS.prz_scontoQ_id, " & vbCrlf & _
 		"  		gtb_prezzi.prz_iva_id = INS.prz_iva_id, " & vbCrlf & _
 		"  		gtb_prezzi.prz_var_euro = INS.prz_var_euro, " & vbCrlf & _
 		"  		gtb_prezzi.prz_var_sconto = INS.prz_var_sconto, " & vbCrlf & _
 		"  		gtb_prezzi.prz_non_vendibile = INS.prz_non_vendibile, " & vbCrlf & _
 		"  		gtb_prezzi.prz_offerta_dal = INS.prz_offerta_dal, " & vbCrlf & _
 		"  		gtb_prezzi.prz_offerta_al =  INS.prz_offerta_al " & vbCrlf & _
 		"  	FROM (inserted INS INNER JOIN deleted DEL ON INS.prz_id = DEL.prz_id) " & vbCrlf & _
 		"  		INNER JOIN gtb_listini L_ancestor ON DEL.prz_listino_id = L_ancestor.listino_id  " & vbCrlf & _
 		"  		INNER JOIN gtb_listini L_child ON L_ancestor.listino_id = L_child.listino_ancestor_id " & vbCrlf & _
 		"  	WHERE gtb_prezzi.prz_listino_id = L_child.listino_id AND " & vbCrlf & _
 		"  		gtb_prezzi.prz_prezzo = DEL.prz_prezzo AND " & vbCrlf & _
 		" 		gtb_prezzi.prz_visibile = DEL.prz_visibile AND " & vbCrlf & _
 		" 		gtb_prezzi.prz_promozione = DEL.prz_promozione AND " & vbCrlf & _
 		" 		gtb_prezzi.prz_variante_id = DEL.prz_variante_id AND " & vbCrlf & _
 		" 		ISNULL(gtb_prezzi.prz_scontoQ_id, 0) = ISNULL(DEL.prz_scontoQ_id, 0) AND " & vbCrlf & _
 		" 		ISNULL(gtb_prezzi.prz_iva_id, 0) = ISNULL(DEL.prz_iva_id, 0) AND " & vbCrlf & _
 		" 		gtb_prezzi.prz_var_euro = DEL.prz_var_euro AND " & vbCrlf & _
 		" 		gtb_prezzi.prz_var_sconto = DEL.prz_var_sconto AND " & vbCrlf & _
 		" 		ISNULL(gtb_prezzi.prz_non_vendibile, 0) = ISNULL(DEL.prz_non_vendibile, 0) AND " & vbCrlf & _
 		" 		ISNULL(gtb_prezzi.prz_offerta_dal, GETDATE()) = ISNULL(DEL.prz_offerta_dal, GETDATE()) AND " & vbCrlf & _
 		" 		ISNULL(gtb_prezzi.prz_offerta_al, GETDATE()) = ISNULL(DEL.prz_offerta_al, GETDATE()) " & vbCrlf & _
		";"
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 197
'...........................................................................................
' Nicola 29/10/2014
'...........................................................................................
' aggiunge parametro per gestire la creazione delle pagine offerte/filtri per i marchi
'...........................................................................................
function Aggiornamento__B2B__197(conn)
	Aggiornamento__B2B__197 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__197(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "B2B_PAGINA_OFFERTE_MARCHIO", _
									0, _
									"Attiva la generazione dei filtri per i prodotti in base al marchio.", _
									"", _
									adGUID, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 198
'...........................................................................................
'	Nicola, 31/10/2014
'...........................................................................................
'   aggiungo campi shopping cart ed ordine per gestione campi colli, pesi e dimensioni
'...........................................................................................
function Aggiornamento__B2B__198(conn)
	Aggiornamento__B2B__198 = _	
		" ALTER TABLE grel_art_valori ADD " + _
		"	rel_peso_netto real NULL, " + _
		"	rel_peso_lordo real NULL, " + _
		"	rel_colli_num INT null, " + _
		"	rel_collo_pezzi_per INT NULL, " + _
		"	rel_collo_width real NULL, " + _
		"	rel_collo_height real NULL, " + _
		"	rel_collo_lenght real NULL, " + _
		"	rel_collo_volume real NULL " + _
		" ; " + _
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_tot_colli INT NULL, " + _
		"	dett_tot_peso_netto real NULL, " + _
		"	dett_tot_peso_lordo real NULL, " + _
		"	dett_tot_volume real NULL, " + _
		" 	dett_rif_spedizione nvarchar(50) " + _
		" ; " + _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_dett_tot_colli INT NULL, " + _
		"	sc_colli INT NULL, " + _
		"	sc_totale_colli INT NULL, " + _
		"	sc_dett_tot_peso_netto real NULL, " + _
		"	sc_peso_netto real NULL, " + _
		"	sc_totale_peso_netto real NULL, " + _
		"	sc_dett_tot_peso_lordo real NULL, " + _
		"	sc_peso_lordo real NULL, " + _
		"	sc_totale_peso_lordo real NULL, " + _
		"	sc_dett_tot_volume real NULL, " + _
		"	sc_volume real NULL, " + _
		"	sc_totale_volume real NULL" + _
		" ; " + _
		" ALTER TABLE gtb_dettagli_ord ADD " + _
		"	det_tot_colli INT NULL, " + _
		"	det_tot_peso_netto real NULL, " + _
		"	det_tot_peso_lordo real NULL, " + _
		"	det_tot_volume real NULL, " + _
		" 	det_rif_spedizione nvarchar(50) " + _
		" ; " + _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_dett_tot_colli INT NULL, " + _
		"	ord_colli INT NULL, " + _
		"	ord_totale_colli INT NULL, " + _
		"	ord_dett_tot_peso_netto real NULL, " + _
		"	ord_peso_netto real NULL, " + _
		"	ord_totale_peso_netto real NULL, " + _
		"	ord_dett_tot_peso_lordo real NULL, " + _
		"	ord_peso_lordo real NULL, " + _
		"	ord_totale_peso_lordo real NULL, " + _
		"	ord_dett_tot_volume real NULL, " + _
		"	ord_volume real NULL, " + _
		"	ord_totale_volume real NULL" + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 199
'...........................................................................................
'	Nicola, 31/10/2014
'...........................................................................................
' 	aggiorna trigger e stored procedure per aggiunta campi volumi, colli e peso
'...........................................................................................
function Aggiornamento__B2B__199(conn)
	Aggiornamento__B2B__199 = _	
		"DROP PROCEDURE gsp_totale_ordini ; " + _
		"DROP PROCEDURE gsp_totale_shopping_cart ; " + _
		"DROP TRIGGER gtb_dett_cart_update ; " + _
		"DROP TRIGGER [gtb_shopping_cart_update] ; " + _
		"DROP TRIGGER [gtb_dettagli_ord_update] ; " + _
		"DROP TRIGGER [gtb_ordini_update] ; " + _
		ReadFileContent(Server.MapPath("subscripts/Aggiornamento_B2B_199.sql"))
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 200
'...........................................................................................
'	Giacomo, 18/11/2014
'...........................................................................................
'   aggiungo campi su gItb_articoli per gestione campi colli, pesi e dimensioni
'...........................................................................................
function Aggiornamento__B2B__200(conn)
	Aggiornamento__B2B__200 = _	
		" ALTER TABLE gItb_articoli ADD " + _
		"	Iart_peso_netto real NULL, " + _
		"	Iart_peso_lordo real NULL, " + _
		"	Iart_colli_num INT null, " + _
		"	Iart_collo_pezzi_per INT NULL, " + _
		"	Iart_collo_width real NULL, " + _
		"	Iart_collo_height real NULL, " + _
		"	Iart_collo_lenght real NULL, " + _
		"	Iart_collo_volume real NULL " + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 201
'...........................................................................................
' Giacomo 21/11/2014
'...........................................................................................
' aggiunge parametro per attivare i campi per gestione campi colli, pesi e dimensioni articoli
'...........................................................................................
function Aggiornamento__B2B__201(conn)
	Aggiornamento__B2B__201 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__201(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "ATTIVA_GESTIONE_COLLI-PESI-DIMENSIONI", _
									0, _
									"Vengono attivati i campi riguardanti colli, pesi e dimensioni per gli articoli.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 202
'...........................................................................................
' Giacomo 05/12/2014
'...........................................................................................
' creazione tabelle per fatturazione
'...........................................................................................
function Aggiornamento__B2B__202(conn)
	Aggiornamento__B2B__202 = _
		" CREATE TABLE " & SQL_Dbo(conn) & " gtb_fatture( " & vbCrlf & _
		" 	fa_id " & SQL_PrimaryKey(conn, "gtb_fatture") & ", " & vbCrlf & _
		" 	fa_emittente_id int NULL, " & vbCrlf & _
		" 	fa_emit_NomeElencoIndirizzi " & SQL_CharField(Conn, 255) & " NOT NULL, " & vbCrlf & _
		" 	fa_emit_CognomeElencoIndirizzi " & SQL_CharField(Conn, 255) & " NOT NULL, " & vbCrlf & _
		" 	fa_emit_NomeOrganizzazioneElencoIndirizzi " & SQL_CharField(Conn, 255) & " NOT NULL, " & vbCrlf & _
		" 	fa_emit_IndirizzoElencoIndirizzi " & SQL_CharField(Conn, 255) & " NULL, " & vbCrlf & _
		" 	fa_emit_CittaElencoIndirizzi " & SQL_CharField(Conn, 100) & " NULL, " & vbCrlf & _
		" 	fa_emit_StatoProvElencoIndirizzi " & SQL_CharField(Conn, 50) & " NULL, " & vbCrlf & _
		" 	fa_emit_CAPElencoIndirizzi " & SQL_CharField(Conn, 10) & " NULL, " & vbCrlf & _
		" 	fa_emit_CountryElencoIndirizzi " & SQL_CharField(Conn, 50) & " NULL, " & vbCrlf & _
		" 	fa_emit_CF " & SQL_CharField(Conn, 20) & " NULL, " & vbCrlf & _
		" 	fa_emit_partita_iva " & SQL_CharField(Conn, 20) & " NULL, " & vbCrlf & _
		" 	fa_intestatario_id int NULL, " & vbCrlf & _
		"	fa_int_esente_iva bit NULL, " & vbCrlf & _
		" 	fa_int_NomeElencoIndirizzi " & SQL_CharField(Conn, 255) & " NOT NULL, " & vbCrlf & _
		" 	fa_int_CognomeElencoIndirizzi " & SQL_CharField(Conn, 255) & " NOT NULL, " & vbCrlf & _
		" 	fa_int_NomeOrganizzazioneElencoIndirizzi " & SQL_CharField(Conn, 255) & " NOT NULL, " & vbCrlf & _
		" 	fa_int_IndirizzoElencoIndirizzi " & SQL_CharField(Conn, 255) & " NULL, " & vbCrlf & _
		" 	fa_int_CittaElencoIndirizzi " & SQL_CharField(Conn, 100) & " NULL, " & vbCrlf & _
		" 	fa_int_StatoProvElencoIndirizzi " & SQL_CharField(Conn, 50) & " NULL, " & vbCrlf & _
		" 	fa_int_CAPElencoIndirizzi " & SQL_CharField(Conn, 10) & " NULL, " & vbCrlf & _
		" 	fa_int_CountryElencoIndirizzi " & SQL_CharField(Conn, 50) & " NULL, " & vbCrlf & _
		" 	fa_int_CF " & SQL_CharField(Conn, 20) &" NULL, " & vbCrlf & _
		" 	fa_int_partita_iva " & SQL_CharField(Conn, 20) & " NULL, " & vbCrlf & _
		" 	fa_data_fattura smalldatetime NOT NULL, " & vbCrlf & _
		" 	fa_data_scadenza smalldatetime NULL, " & vbCrlf & _
		"	fa_data_invio_a_cliente smalldatetime NULL, " & vbCrlf & _
		"	fa_data_invio_a_commercialista smalldatetime NULL, " & vbCrlf & _
		"	fa_data_pagamento smalldatetime NULL, " & vbCrlf & _
		" 	fa_serie_fattura_id int NULL, " & vbCrlf & _
		" 	fa_numero int NULL, " & vbCrlf & _
		" 	fa_serie " & SQL_CharField(Conn, 10) & " NULL, " & vbCrlf & _
		" 	fa_anno int NOT NULL, " & vbCrlf & _
		" 	fa_note " & SQL_CharField(Conn, 0) & " NULL, " & vbCrlf & _
		" 	fa_totale_imponibile money NULL, " & vbCrlf & _
		" 	fa_totale_iva money NULL, " & vbCrlf & _
		" 	fa_totale money NULL, " & vbCrlf & _
		" 	fa_metodo_pagamento_id int NULL, " & vbCrlf & _
		" 	fa_is_bozza bit NULL, " & vbCrlf & _
		" 	fa_insAdmin_id int NULL, " & vbCrlf & _
		" 	fa_insData smalldatetime NULL, " & vbCrlf & _
		" 	fa_modAdmin_id int NULL, " & vbCrlf & _
		" 	fa_modData smalldatetime NULL " & vbCrlf & _
		" ); " & vbCrlf & _
		" CREATE TABLE " & SQL_Dbo(conn) & " gtb_fatture_dettagli( " & vbCrlf & _
		" 	fad_id " & SQL_PrimaryKey(conn, "gtb_fatture_dettagli") & ", " & vbCrlf & _
		" 	fad_fattura_id int NOT NULL, " & vbCrlf & _
		" 	fad_qta int NULL, " & vbCrlf & _
		" 	fad_prezzo_unitario money NULL, " & vbCrlf & _
		" 	fad_prezzo_listino money NULL, " & vbCrlf & _
		" 	fad_sconto_perc real NULL, " & vbCrlf & _
		" 	fad_sconto money NULL, " & vbCrlf & _
		" 	fad_aliquota_iva_id int NULL, " & vbCrlf & _
		" 	fad_iva real NULL, " & vbCrlf & _
		SQL_MultiLanguageFieldComplete(conn, "fad_desc_<lingua> " & SQL_CharField(Conn, 500) & " NULL ") & ", " & vbCrlf & _
		" 	fad_totale_imponibile money NULL, " & vbCrlf & _
		" 	fad_totale_iva money NULL, " & vbCrlf & _
		" 	fad_totale money NULL, " & vbCrlf & _
		" 	fad_codice_articolo " & SQL_CharField(Conn, 255) & " NULL, " & vbCrlf & _
		" 	fad_nome_tabella_esterna " & SQL_CharField(Conn, 255) & " NULL, " & vbCrlf & _
		"	fad_id_in_tabella_esterna int NULL " & vbCrlf & _
		" ); " & vbCrlf & _
		" CREATE TABLE " & SQL_Dbo(conn) & " gtb_fatture_serie( " & vbCrlf & _
		" 	fs_id " & SQL_PrimaryKey(conn, "gtb_fatture_serie") & ", " & vbCrlf & _
		SQL_MultiLanguageFieldComplete(conn, "fs_nome_<lingua> " & SQL_CharField(Conn, 255) & " NULL ") & ", " & vbCrlf & _
		" 	fs_codice " & SQL_CharField(Conn, 50) & " NULL, " & vbCrlf & _
		" 	fs_ordine int NULL " & vbCrlf & _
		" ); " & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture", "fa_emittente_id", "gtb_agenti", "ag_id", false, "") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture", "fa_intestatario_id", "gtb_rivenditori", "riv_id", false, "") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture", "fa_serie_fattura_id", "gtb_fatture_serie", "fs_id", false, "") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture", "fa_metodo_pagamento_id", "gtb_modipagamento", "mosp_id", false, "") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture", "fa_insAdmin_id", "tb_admin", "ID_admin", false, "") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture", "fa_modAdmin_id", "tb_admin", "ID_admin", false, "_2") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture_dettagli", "fad_fattura_id", "gtb_fatture", "fa_id", true, "") & vbCrlf & _
		SQL_AddForeignKey(conn, "gtb_fatture_dettagli", "fad_aliquota_iva_id", "gtb_iva", "iva_id", false, "") & vbCrlf & _
		DropObject(conn, "gsp_totale_fatture", "PROCEDURE") & vbCrLf & _
		" CREATE PROCEDURE [dbo].[gsp_totale_fatture] " & vbCrLf & _
		" 	@fa_id INT  " & vbCrLf & _
		" AS  " & vbCrLf & _
		" BEGIN  " & vbCrLf & _
		" 	--calcolo dei totali per i dettagli della fattura " & vbCrLf & _
		" 	UPDATE gtb_fatture_dettagli   " & vbCrLf & _
		" 	SET fad_totale = ROUND(ISNULL(fad_prezzo_unitario,0)*ISNULL(fad_qta,0),2) ,   " & vbCrLf & _
		" 		fad_totale_iva = ROUND(ISNULL(fad_prezzo_unitario,0)*ISNULL(fad_qta,0)*ISNULL(fad_iva,0)/100,2) " & vbCrLf & _
		" 	WHERE fad_fattura_id=@fa_id " & vbCrLf & _
		" 	 " & vbCrLf & _
		"    " & vbCrLf & _
		" 	--calcolo dei totali dei dettagli sulla testata della fattura " & vbCrLf & _
		" 	UPDATE gtb_fatture  " & vbCrLf & _
		" 	SET fa_totale=(SELECT SUM(fad_totale) FROM gtb_fatture_dettagli WHERE fad_fattura_id=@fa_id AND fad_totale IS NOT NULL) ,  " & vbCrLf & _
		" 		fa_totale_iva=(SELECT SUM(fad_totale_iva) FROM gtb_fatture_dettagli WHERE fad_fattura_id=@fa_id AND fad_totale_iva IS NOT NULL)  " & vbCrLf & _
		" 	WHERE fa_id=@fa_id  " & vbCrLf & _
		" END; " & vbCrLf & _
		DropObject(conn, "gtb_fatture_insert", "TRIGGER") & vbCrlf & _
		" CREATE TRIGGER [gtb_fatture_insert] " & vbCrlf & _
		" ON [dbo].[gtb_fatture] " & vbCrlf & _
		" AFTER INSERT " & vbCrlf & _
		" AS " & vbCrlf & _
		"  " & vbCrlf & _
		" DECLARE @F_id INT " & vbCrlf & _
		" /*apre recordset con fatture inserite */ " & vbCrlf & _
		" DECLARE rs CURSOR local FAST_FORWARD FOR  " & vbCrlf & _
		" SELECT DISTINCT fa_id FROM inserted " & vbCrlf & _
		"  " & vbCrlf & _
		" OPEN rs " & vbCrlf & _
		" FETCH NEXT FROM rs INTO @F_id " & vbCrlf & _
		" WHILE @@FETCH_STATUS = 0 " & vbCrlf & _
		" BEGIN " & vbCrlf & _
		" 	/* esegue ricalcolo della fattura */ " & vbCrlf & _
		" 	EXEC gsp_totale_fatture @fa_id=@F_id " & vbCrlf & _
		" 	FETCH NEXT FROM rs INTO @F_id " & vbCrlf & _
		" END; " & vbCrlf & _
		DropObject(conn, "gtb_fatture_update", "TRIGGER") & vbCrlf & _
		" CREATE TRIGGER [dbo].[gtb_fatture_update] " & vbCrlf & _
		" ON [dbo].[gtb_fatture] " & vbCrlf & _
		" AFTER UPDATE " & vbCrlf & _
		" AS " & vbCrlf & _
		"  " & vbCrlf & _
		" DECLARE @F_id INT " & vbCrlf & _
		" /* " & vbCrlf & _
		" apre recordset delle fatture modificate " & vbCrlf & _
		" in almeno uno dei campi che concorrono al calcolo dei totali " & vbCrlf & _
		" */ " & vbCrlf & _
		" DECLARE rs CURSOR local FAST_FORWARD FOR " & vbCrlf & _
		"	SELECT DISTINCT inserted.fa_id FROM " & vbCrlf & _
		"	inserted INNER JOIN deleted ON " & vbCrlf & _
		"		inserted.fa_id = deleted.fa_id " & vbCrlf & _
		"		AND (  " & vbCrlf & _
		"			1 = 1 " & vbCrlf & _
		"			) " & vbCrlf & _
		" OPEN rs " & vbCrlf & _
		" FETCH NEXT FROM rs INTO @F_id " & vbCrlf & _
		" WHILE @@FETCH_STATUS = 0 " & vbCrlf & _
		" BEGIN " & vbCrlf & _
		"	/* esegue ricalcolo della fattura */ " & vbCrlf & _
		"	EXEC gsp_totale_fatture @fa_id=@F_id " & vbCrlf & _
		"	FETCH NEXT FROM rs INTO @F_id " & vbCrlf & _
		" END "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 203
'...........................................................................................
' Giacomo 05/12/2014
'...........................................................................................
' aggiunge parametro per attivare la sezione fatture
'...........................................................................................
function Aggiornamento__B2B__203(conn)
	Aggiornamento__B2B__203 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__203(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "ATTIVA_FATTURE", _
									0, _
									"Viene attivata la gestione fatture nel B2B.", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 204
'...........................................................................................
' Giacomo 05/12/2014
'...........................................................................................
' aggiunge parametro per id emittente fatture di default
'...........................................................................................
function Aggiornamento__B2B__204(conn)
	Aggiornamento__B2B__204 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__204(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "FATTURE_ID_EMITTENTE_DEFAULT", _
									0, _
									"Viene attivata la gestione fatture nel B2B.", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									null, null, null, null, null)
	end if
end function

'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 205
'...........................................................................................
'	Nicola 15/12/2014
'...........................................................................................
'   aggiungo campi su gtb_ordini per data di spedizione, evasione, consegna
'...........................................................................................
function Aggiornamento__B2B__205(conn)
	Aggiornamento__B2B__205 = _	
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_spedizione_ordine_data SMALLDATETIME NULL, " + _
		"	ord_evasione_ordine_data SMALLDATETIME NULL, " + _
		"	ord_consegna_ordine_data SMALLDATETIME NULL " + _
		" ; "
end function
'*******************************************************************************************



'*******************************************************************************************
'AGGIORNAMENTO B2B 206
'...........................................................................................
'	Nicola 16/12/2014
'...........................................................................................
'   aggiungo campi su gtb_shopping_cart per data di spedizione, evasione, consegna
'...........................................................................................
function Aggiornamento__B2B__206(conn)
	Aggiornamento__B2B__206 = _	
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_spedizione_ordine_data SMALLDATETIME NULL, " + _
		"	sc_evasione_ordine_data SMALLDATETIME NULL, " + _
		"	sc_consegna_ordine_data SMALLDATETIME NULL " + _
		" ; "
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO B2B 207
'...........................................................................................
' Nicola 05/12/2014
'...........................................................................................
' aggiunge parametro per attivazione / disattivazione pesi e volumi
'...........................................................................................
function Aggiornamento__B2B__207(conn)
	Aggiornamento__B2B__207 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__207(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "ORDINI_GESTIONE_PESI", _
									0, _
									"Viene attivata la gestione dei pesi negli articoli, shopping cart e negli ordini", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									0, null, null, null, null)
		CALL AddParametroSito(conn, "ORDINI_GESTIONE_VOLUMI", _
									0, _
									"Viene attivata la gestione dei volumi negli articoli, shopping cart e negli ordini", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									0, null, null, null, null)
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 208
'...........................................................................................
' Luca 20/04/2015
'...........................................................................................
' aggiungo campo per il codice container che verrà utilizzato per la gestione degli ordini
' a fornitore
'...........................................................................................
function Aggiornamento__B2B__208(conn)
	Aggiornamento__B2B__208 = _
		" ALTER TABLE gtb_dett_cart ADD " + _
		"	dett_cod_spedizione nvarchar(50) NULL ;" + _
		" ALTER TABLE gtb_dettagli_ord ADD " + _
		"	det_cod_spedizione nvarchar(50) NULL ;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 209
'...........................................................................................
' Luca 24/04/2015
'...........................................................................................
' rinomino campo dett_rif_spedizione, perché non utilizzato, in dett_rif_ordine che verrà 
' utilizzato per la gestione degli ordini a fornitore
'...........................................................................................
function Aggiornamento__B2B__209(conn)
	Aggiornamento__B2B__209 = _
	" EXEC sp_rename 'gtb_dett_cart.dett_rif_spedizione', 'dett_rif_ordine', 'COLUMN' ;" + _
	" EXEC sp_rename 'gtb_dettagli_ord.det_rif_spedizione', 'det_rif_ordine', 'COLUMN' ;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 210
'...........................................................................................
' Luca 29/04/2015
'...........................................................................................
' aggiungo i campi ord_valu_id ed ord_valu_cambio, su gtb_ordini, che verranno utilizzati
' per la gestione degli ordini a fornitore
'...........................................................................................
function Aggiornamento__B2B__210(conn)
	Aggiornamento__B2B__210 = _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_valu_id int NULL ;" + _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_valu_cambio money NULL ;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 211
'...........................................................................................
' Luca 19/05/2015
'...........................................................................................
' creo la tabella per la gestione dei pagamenti (paypal, cc)
' aggiungo il flag mosp_pag_immediato su gtb_modipagamento per distinguere i pagamenti
' immediati da quelli differiti
'...........................................................................................
function Aggiornamento__B2B__211(conn)
	Aggiornamento__B2B__211 = _
        " CREATE TABLE dbo.gtb_pagamenti ( " + _
        " 	pag_id int IDENTITY(1,1) NOT NULL, " + _
        "   pag_data smalldatetime NULL, " + _
        "   pag_importo money NULL, " + _
		"   pag_ordine_id int NULL, " + _
        "   pag_stato_ordine_id int NULL, " + _
		"	pag_mosp_id int NULL, " + _
        "   pag_RAW text NULL, " + _
		"	pag_tran_id nvarchar(30) NULL); " + _
		" ALTER TABLE dbo.gtb_pagamenti WITH NOCHECK ADD " + _
        "   CONSTRAINT PK_gtb_pagamenti PRIMARY KEY ( pag_id ); " + _
		" ALTER TABLE gtb_modipagamento ADD " + _
		"	mosp_pag_immediato bit NULL; "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 212
'...........................................................................................
' Luca 22/05/2015
'...........................................................................................
' aggiungo ord_se_saldato su gtb_ordini per identificare se l'ordine è saldato o meno
' aggiungo sc_mosp_id su gtb_shopping_cart per salvare il metodo di pagamento scelto
'...........................................................................................
function Aggiornamento__B2B__212(conn)
	Aggiornamento__B2B__212 = _
		" ALTER TABLE gtb_ordini ADD " + _
		"	ord_se_saldato bit NULL ;" + _
		" ALTER TABLE gtb_shopping_cart ADD " + _
		"	sc_mosp_id int NULL ;"
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 213
'...........................................................................................
' Luca 24/07/2015
'...........................................................................................
' aggiunge parametri per il controllo sulla giacenza (percentuale giacenza e quantità minima
' per effettuare il controllo)
'...........................................................................................
function Aggiornamento__B2B__213(conn)
	Aggiornamento__B2B__213 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__213(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "PERCENTUALE_CONTROLLO_GIACENZA", _
									0, _
									"Percentuale utilizzata per effettuare il controllo sulla giacenza", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									20, null, null, null, null)
		CALL AddParametroSito(conn, "QUANTITA_MINIMA_CONTROLLO_GIACENZA", _
									0, _
									"Quantità minima oltre la quale bisogna effettuare il controllo sulla giacenza", _
									"", _
									adNumeric, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									15, null, null, null, null)
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 214
'...........................................................................................
' Luca 27/07/2015
'...........................................................................................
' aggiungo tabella categorie iva clienti, tabella categorie iva articoli e tabella di
' relazione tra le due precedenti per la nuova gestione dell'iva;
' aggiungo gli id di collegamento anche sulle tabelle di rivenditori e articoli
'...........................................................................................
function Aggiornamento__B2B__214(conn)
	Aggiornamento__B2B__214 = _
		" CREATE TABLE dbo.gtb_civa_riv ( " & _
		" 	cir_id int IDENTITY(1,1) NOT NULL, " & _
		" 	cir_codice nvarchar(20) NULL, " & _
		" 	cir_descrizione nvarchar(50) NULL); " & _
		" ALTER TABLE dbo.gtb_civa_riv " & _
		" 	ADD CONSTRAINT PK_gtb_civa_riv PRIMARY KEY (cir_id); " & _
		" ALTER TABLE dbo.gtb_rivenditori " & _
		" 	ADD riv_civa_id int NULL; " & _
		" CREATE TABLE dbo.gtb_civa_art ( " & _
		" 	cia_id int IDENTITY(1,1) NOT NULL, " & _
		" 	cia_codice nvarchar(20) NULL, " & _
		" 	cia_descrizione nvarchar(50) NULL); " & _
		" ALTER TABLE dbo.gtb_civa_art " & _
		" 	ADD CONSTRAINT PK_gtb_civa_art PRIMARY KEY (cia_id); " & _
		" ALTER TABLE dbo.gtb_articoli " & _
		" 	ADD art_civa_id int NULL; " & _
		" CREATE TABLE dbo.grel_civa ( " & _
		" 	civa_id int IDENTITY(1,1) NOT NULL, " & _
		" 	civa_riv_id int NULL, " & _
		" 	civa_art_id int NULL, " & _
		" 	civa_valore real NULL ); " & _
		" ALTER TABLE dbo.grel_civa " & _
		" 	ADD CONSTRAINT PK_grel_civa PRIMARY KEY (civa_id); "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 215
'...........................................................................................
' Luca 28/07/2015
'...........................................................................................
' elimino la foreign key che collega gtb_dett_cart con gtb_iva
' rinomino dett_iva_id in dett_iva_valore e cambio il tipo del campo da int a real
'...........................................................................................
function Aggiornamento__B2B__215(conn)
	Aggiornamento__B2B__215 = _
	" ALTER TABLE gtb_dett_cart DROP CONSTRAINT FK_gtb_dett_cart_gtb_iva; " & _
	" ALTER TABLE gtb_dett_cart ALTER COLUMN dett_iva_id real; " & _
	" EXEC sp_rename 'gtb_dett_cart.dett_iva_id', 'dett_iva_valore', 'COLUMN' "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 216
'...........................................................................................
' Luca 18/08/2015
'...........................................................................................
' aggiunge parametro contenente il codice del metodo di pagamento PAYPAL
'...........................................................................................
function Aggiornamento__B2B__216(conn)
	Aggiornamento__B2B__216 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__216(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "CODICE_METODO_PAGAMENTO_PAYPAL", _
									0, _
									"Codice del metodo di pagamento PAYPAL", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									"PAYPAL", "PAYPAL", "PAYPAL", "PAYPAL", "PAYPAL")
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 217
'...........................................................................................
' Luca 19/08/2015
'...........................................................................................
' aggiorna i trigger su insert e update dei dettagli della shoppingcart per la nuova 
' gestione dell'iva ed aggiunge la stored procedure per il calcolo dell'iva
'...........................................................................................
function Aggiornamento__B2B__217(conn)
	Aggiornamento__B2B__217 = _
		"DROP VIEW gv_CartDetail; " + _
		"DROP VIEW gv_articoli; " + _
		"DROP TRIGGER gtb_dett_cart_insert; " + _
		"DROP TRIGGER gtb_dett_cart_update; " + _
		"DROP PROCEDURE gsp_totale_shopping_cart; " + _
		ReadFileContent(Server.MapPath("subscripts/Aggiornamento_B2B_217.sql"))
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 218
'...........................................................................................
' Luca 24/08/2015
'...........................................................................................
' aggiunge parametro che indica se è attiva la nuova gestione dell'iva
'...........................................................................................
function Aggiornamento__B2B__218(conn)
	Aggiornamento__B2B__218 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__218(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "NUOVA_GESTIONE_IVA", _
									0, _
									"Indica se è attiva la nuova gestione dell'iva", _
									"", _
									adBoolean, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									0, null, null, null, null)
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 219
'...........................................................................................
' Nicola 09/09/2015
'...........................................................................................
' creo tabella per gestione registrazione acquisti
'...........................................................................................
function Aggiornamento__B2B__219(conn)
	Aggiornamento__B2B__219 = _
        " CREATE TABLE dbo.glog_rivenditori_acquisti ( " + _
        " 	rac_id " + SQL_PrimaryKey(conn, "grel_rivenditori_acquisti") + ", " + _
        "   rac_data_ultimo_acquisto smalldatetime NULL, " + _
        "   rac_rivenditore_id int NULL, " + _
		"	rac_riv_sede_id int NULL, " + _
        "   rac_art_var_id int NULL, " + _
		" 	rac_art_id int NULL " + _
		"	); " + _
		SQL_AddForeignKey(conn, "glog_rivenditori_acquisti", "rac_rivenditore_id", "gtb_rivenditori", "riv_id", true, "") + _
		SQL_AddForeignKey(conn, "glog_rivenditori_acquisti", "rac_riv_sede_id", "tb_indirizzario", "IDElencoIndirizzi", false, "") + _
		SQL_AddForeignKey(conn, "glog_rivenditori_acquisti", "rac_art_var_id", "grel_art_Valori", "rel_id", true, "") + _
		SQL_AddForeignKey(conn, "glog_rivenditori_acquisti", "rac_art_id", "gtb_articoli", "art_id", false, "")
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 220
'...........................................................................................
' Luca 24/09/2015
'...........................................................................................
' aggiunge parametro contenente il codice del metodo di pagamento VIRTUALPAY
'...........................................................................................
function Aggiornamento__B2B__220(conn)
	Aggiornamento__B2B__220 = " SELECT * FROM AA_Versione "
end function

function AggiornamentoSpeciale__B2B__220(conn)
	if cString(GetValueList(conn , NULL, "SELECT sito_nome FROM tb_siti WHERE id_sito = " & NEXTB2B)) <> "" then
		CALL AddParametroSito(conn, "CODICE_METODO_PAGAMENTO_VIRTUALPAY", _
									0, _
									"Codice del metodo di pagamento VIRTUALPAY", _
									"", _
									adVarChar, _
									0, _
									"", _
									1, _
									1, _
									NEXTB2B, _
									"VIRTUALPAY", "VIRTUALPAY", "VIRTUALPAY", "VIRTUALPAY", "VIRTUALPAY")
	end if
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 221
'...........................................................................................
' Luca 13/10/2015
'...........................................................................................
' elimino la foreign key che collega gtb_spese_spedizione con gtb_iva
'...........................................................................................
function Aggiornamento__B2B__221(conn)
	Aggiornamento__B2B__221 = _
	" ALTER TABLE gtb_spese_spedizione DROP CONSTRAINT FK_gtb_spese_spedizione__gtb_iva "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 222
'...........................................................................................
' Luca 21/10/2015
'...........................................................................................
' aggiorna trigger e stored procedure per la risoluzione dei problemi legati alla nuova 
' gestione dell'iva
'...........................................................................................
function Aggiornamento__B2B__222(conn)
	Aggiornamento__B2B__222 = _
		"DROP PROCEDURE gsp_totale_shopping_cart; " + _
		"DROP PROCEDURE gsp_calcola_iva; " + _
		"DROP TRIGGER gtb_dett_cart_update; " + _
		ReadFileContent(Server.MapPath("subscripts/Aggiornamento_B2B_222.sql"))
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 223
'...........................................................................................
' Luca 09/12/2015
'...........................................................................................
' creo le tabella e la relazione per la gestione delle promozioni
' su shoppingcart e ordine inserisco l'id di collegamento alla promozione
'...........................................................................................
function Aggiornamento__B2B__223(conn)
	Aggiornamento__B2B__223 = _
		" CREATE TABLE dbo.gtb_promozioni ( " & _
		" 	promo_id int IDENTITY(1,1) NOT NULL, " & _
		" 	promo_descr nvarchar(255) NULL, " & _
		" 	promo_valore real NULL, " & _
		" 	promo_inizio_validita smalldatetime NULL, " & _
		" 	promo_fine_validita smalldatetime NULL); " & _
		" ALTER TABLE dbo.gtb_promozioni " & _
		" 	ADD CONSTRAINT PK_gtb_promozioni PRIMARY KEY (promo_id); " & _
		" CREATE TABLE dbo.grel_promo_articoli ( " & _
		" 	pa_id int IDENTITY(1,1) NOT NULL, " & _
		" 	pa_promo_id int NULL, " & _
		" 	pa_art_id int NULL); " & _
		" ALTER TABLE dbo.grel_promo_articoli " & _
		" 	ADD CONSTRAINT PK_grel_promo_articoli PRIMARY KEY (pa_id); " & _
		" ALTER TABLE dbo.gtb_shopping_cart ADD " & _
		" 	sc_promo_id int NULL, " & _
		" 	sc_tipo_id int NULL; " & _
		" ALTER TABLE dbo.gtb_dett_cart " & _
		" 	ADD dett_promo_id int NULL; " & _
		" ALTER TABLE dbo.gtb_ordini " & _
		" 	ADD ord_promo_id int NULL; " & _
		" ALTER TABLE dbo.gtb_dettagli_ord " & _
		" 	ADD det_promo_id int NULL; "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 224
'...........................................................................................
' Luca 12/01/2016
'...........................................................................................
' aggiungo i campi per le traduzioni alla tabella delle promozioni
'...........................................................................................
function Aggiornamento__B2B__224(conn)
	Aggiornamento__B2B__224 = _
		" ALTER TABLE dbo.gtb_promozioni DROP COLUMN " & _
		"	promo_descr; " & _
		" ALTER TABLE dbo.gtb_promozioni ADD " & _
		" 	promo_nome_it nvarchar(255) NULL, " & _
		" 	promo_nome_en nvarchar(255) NULL, " & _
		" 	promo_nome_de nvarchar(255) NULL, " & _
		" 	promo_nome_fr nvarchar(255) NULL, " & _
		" 	promo_nome_es nvarchar(255) NULL, " & _
		" 	promo_nome_ru nvarchar(255) NULL, " & _
		" 	promo_nome_cn nvarchar(255) NULL, " & _
		" 	promo_nome_pt nvarchar(255) NULL, " & _
		" 	promo_descrizione_it nvarchar(1000) NULL, " & _
		" 	promo_descrizione_en nvarchar(1000) NULL, " & _
		" 	promo_descrizione_de nvarchar(1000) NULL, " & _
		" 	promo_descrizione_fr nvarchar(1000) NULL, " & _
		" 	promo_descrizione_es nvarchar(1000) NULL, " & _
		" 	promo_descrizione_ru nvarchar(1000) NULL, " & _
		" 	promo_descrizione_cn nvarchar(1000) NULL, " & _
		" 	promo_descrizione_pt nvarchar(1000) NULL; "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 225
'...........................................................................................
' Luca 27/01/2016
'...........................................................................................
' creo tabella per le lettere d'intento
'...........................................................................................
function Aggiornamento__B2B__225(conn)
	Aggiornamento__B2B__225 = _
		" CREATE TABLE dbo.gtb_eccezioni_iva ( " & _
		" 	ei_id int IDENTITY(1,1) NOT NULL, " & _
		" 	ei_descrizione nvarchar(255) NULL, " & _
		" 	ei_riv_id int NULL, " & _
		" 	ei_civa_riv_id int NULL, " & _
		" 	ei_inizio_validita smalldatetime NULL, " & _
		" 	ei_fine_validita smalldatetime NULL); "
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 226
'...........................................................................................
' Luca 28/01/2016
'...........................................................................................
' modifico le procedure che calcolano l'iva in modo da andare a cercare la categoria iva
' del rivenditore prima sulla tabella delle lettere d'intento e poi sui rivenditori
'...........................................................................................
function Aggiornamento__B2B__226(conn)
	Aggiornamento__B2B__226 = _
		"DROP PROCEDURE gsp_totale_shopping_cart; " + _
		"DROP PROCEDURE gsp_calcola_iva; " + _
		ReadFileContent(Server.MapPath("subscripts/Aggiornamento_B2B_226.sql"))
end function
'*******************************************************************************************

'*******************************************************************************************
'AGGIORNAMENTO B2B 227
'...........................................................................................
' Luca 27/04/2016
'...........................................................................................
' aggiungo la data del cambio sulla valuta
'...........................................................................................
function Aggiornamento__B2B__227(conn)
	Aggiornamento__B2B__227 = _
		" ALTER TABLE dbo.gtb_valute ADD " & _
		"	valu_data_cambio smalldatetime NULL; "
end function
'*******************************************************************************************
%>