USE prestigeinternational 

ALTER INDEX PK_rel_contents ON dbo.rel_contents REBUILD
UPDATE STATISTICS rel_contents

ALTER INDEX PK_rel_contents_tags ON dbo.rel_contents_tags REBUILD
UPDATE STATISTICS rel_contents_tags

ALTER INDEX PK_rel_index_pubblicazioni ON dbo.rel_index_pubblicazioni REBUILD
UPDATE STATISTICS rel_index_pubblicazioni

ALTER INDEX IDX_rel_index_url_redirect ON dbo.rel_index_url_redirect REBUILD
ALTER INDEX PK_rel_index_url_redirect ON dbo.rel_index_url_redirect REBUILD
UPDATE STATISTICS rel_index_url_redirect

ALTER INDEX PK_rel_utenti_sito ON dbo.rel_utenti_sito REBUILD
UPDATE STATISTICS rel_utenti_sito

ALTER INDEX PK_Rrel_agenzie_descrittori ON dbo.Rrel_agenzie_descrittori REBUILD
UPDATE STATISTICS Rrel_agenzie_descrittori

ALTER INDEX idx_Rrel_descrittori_realestate ON dbo.Rrel_descrittori_realestate REBUILD
ALTER INDEX PK_Rrel_descrittori_realestate ON dbo.Rrel_descrittori_realestate REBUILD
UPDATE STATISTICS Rrel_descrittori_realestate

--ALTER INDEX IDX_test_agenzie_categoria ON dbo.rtb_agenzie REBUILD
ALTER INDEX PK_rtb_agenzie ON dbo.rtb_agenzie REBUILD
UPDATE STATISTICS rtb_agenzie

ALTER INDEX PK_rtb_Aree ON dbo.rtb_Aree REBUILD
UPDATE STATISTICS dbo.rtb_Aree

ALTER INDEX PK_rtb_categorieRealEstate ON dbo.rtb_categorieRealEstate REBUILD
UPDATE STATISTICS dbo.rtb_categorieRealEstate

ALTER INDEX PK_Rtb_contratti ON dbo.Rtb_contratti REBUILD
UPDATE STATISTICS dbo.Rtb_contratti

ALTER INDEX PK_Rtb_descrittori ON dbo.Rtb_descrittori REBUILD
UPDATE STATISTICS dbo.Rtb_descrittori

ALTER INDEX IX_Rtb_foto ON dbo.Rtb_foto REBUILD
ALTER INDEX PK_Rtb_foto ON dbo.Rtb_foto REBUILD
UPDATE STATISTICS dbo.Rtb_foto

ALTER INDEX IDX_rtb_strutture_agenzia ON dbo.Rtb_strutture REBUILD
ALTER INDEX PK_Rtb_strutture ON dbo.Rtb_strutture REBUILD
ALTER INDEX IDX_rtb_strutture_categoria ON dbo.Rtb_strutture REBUILD
ALTER INDEX IDX_rtb_strutture_contratti ON dbo.Rtb_strutture REBUILD
ALTER INDEX IDX_rtb_strutture_pub_client_id ON dbo.Rtb_strutture REBUILD
UPDATE STATISTICS dbo.Rtb_strutture

ALTER INDEX PK_tb_ADMIN ON dbo.tb_admin REBUILD
UPDATE STATISTICS dbo.tb_admin

ALTER INDEX IX_tb_contents_co_F_ids ON tb_contents REBUILD
ALTER INDEX PK_tb_contents ON tb_contents REBUILD
UPDATE STATISTICS dbo.tb_contents

ALTER INDEX IX_tb_contents_index_geturls ON dbo.tb_contents_index REBUILD
ALTER INDEX IX_tb_contents_index_padre_id ON dbo.tb_contents_index REBUILD
ALTER INDEX IX_tb_contents_index_tipologie_padre_lista ON dbo.tb_contents_index REBUILD
ALTER INDEX IX_tb_contents_index_urls ON dbo.tb_contents_index REBUILD
ALTER INDEX PK_tb_contents_index ON dbo.tb_contents_index REBUILD
UPDATE STATISTICS dbo.tb_contents_index

ALTER INDEX PK_tb_email ON dbo.tb_email REBUILD
UPDATE STATISTICS dbo.tb_email

ALTER INDEX PK_tb_Indirizzario ON dbo.tb_Indirizzario REBUILD
UPDATE STATISTICS dbo.tb_Indirizzario

ALTER INDEX PK_tb_indirizzario_carattech ON dbo.tb_indirizzario_carattech REBUILD
UPDATE STATISTICS dbo.tb_indirizzario_carattech

ALTER INDEX PK_tb_layers ON tb_layers REBUILD
ALTER INDEX IX_tb_layers_id_pag ON dbo.tb_layers REBUILD
UPDATE STATISTICS dbo.tb_layers

ALTER INDEX PK_tb_pages ON tb_pages REBUILD
UPDATE STATISTICS dbo.tb_pages

ALTER INDEX PK_tb_pagineSito ON dbo.tb_pagineSito REBUILD
UPDATE STATISTICS dbo.tb_pagineSito

ALTER INDEX PK_tb_siti_tabelle ON dbo.tb_siti_tabelle REBUILD
UPDATE STATISTICS dbo.tb_siti_tabelle

ALTER INDEX PK_tb_siti_tabelle_pubblicazioni ON dbo.tb_siti_tabelle_pubblicazioni REBUILD
UPDATE STATISTICS dbo.tb_siti_tabelle_pubblicazioni

ALTER INDEX IX_tb_ValoriNumeri ON dbo.tb_ValoriNumeri REBUILD
ALTER INDEX PK_tb_ValoriNumeri ON dbo.tb_ValoriNumeri REBUILD
UPDATE STATISTICS dbo.tb_ValoriNumeri

ALTER INDEX IX_tb_utenti ON dbo.tb_Utenti REBUILD
ALTER INDEX PK_tb_Utenti ON dbo.tb_Utenti REBUILD
UPDATE STATISTICS dbo.tb_Utenti
