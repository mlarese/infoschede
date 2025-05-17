USE prestigeinternational
;

DISABLE TRIGGER [tb_contents_index_delete] ON [tb_contents_index]

;

DECLARE @tabellaAgenzie int
DECLARE @tabellaImmobili int
DECLARE @DATAINIZIO datetime
DECLARE @DATAOGGI datetime
DECLARE @adminId int

set @tabellaAgenzie = 26
set @tabellaImmobili = 11
set @DATAINIZIO = GETDATE()-60
set @DATAOGGI = null --CONVERT(DATETIME, '2014-11-11 00:00:00', 102)
set @adminId = 52


BEGIN TRAN 

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_it, 'it', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_it,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_en, 'en', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_en,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_fr, 'fr', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_fr,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_de, 'de', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_de,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_es, 'es', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_es,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_ru, 'ru', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_ru,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_pt, 'pt', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_pt,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)

INSERT INTO rel_index_url_redirect (riu_idx_id, riu_url, riu_lingua, riu_insData, riu_insAdmin_id, riu_modData, riu_modAdmin_id, riu_co_f_table_id, riu_co_f_key_id)
SELECT (SELECT idx_id FROM v_indice_it WHERE co_F_key_id = rtb_strutture.st_agenzia_id AND co_F_table_id = @tabellaAgenzie),
		idx_link_url_rw_cn, 'cn', GETDATE(), @adminId, GETDATE(), @adminId, co_F_table_id, co_F_key_id
from v_indice INNER JOIN Rtb_strutture ON v_indice.co_F_table_id=@tabellaImmobili AND v_indice.co_F_key_id = Rtb_strutture.st_ID
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND Isnull(idx_link_url_rw_cn,'')<>''
AND isnull(st_visibile,0)=0 
AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
	  OR
	  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
	)


DELETE FROM tb_contents
where co_F_table_id=@tabellaImmobili
AND ISNULL(co_visibile,0)=0
AND co_F_key_id IN (select st_id 
					FROM Rtb_strutture 
					where isnull(st_visibile,0)=0 AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
						AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
							  OR
							  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
							)
					)

DELETE FROM rtb_strutture where isnull(st_visibile,0)=0 AND (st_agenzia_id=2335  OR st_agenzia_id=2336)
						AND  ( ISNULL(st_modData, ISNULL(st_insData, GETDATE()))<@DATAINIZIO
							  OR
							  (@DATAOGGI is not null AND ISNULL(st_modData, ISNULL(st_insData, GETDATE()+5))>@DATAOGGI )
							)
COMMIT TRAN


;

ENABLE TRIGGER  [tb_contents_index_delete] ON [tb_contents_index]