
 CREATE TRIGGER [dbo].[gtb_dettagli_ord_update] 
 ON [dbo].[gtb_dettagli_ord] 
 AFTER UPDATE 
 AS 
  
 DECLARE @O_id INT 
 /* 
 apre recordset degli ordini ai quali è stato modificato un dettaglio 
 in almeno uno dei campi che concorrono al calcolo dei totali 
 */ 
 DECLARE rs CURSOR local FAST_FORWARD FOR  
 	SELECT DISTINCT inserted.det_ord_id FROM 
 	inserted INNER JOIN deleted ON  
 		inserted.det_id = deleted.det_id 
 		AND (  
 			inserted.det_art_var_id <> deleted.det_art_var_id OR 
 			inserted.det_qta <> deleted.det_qta OR 
 			inserted.det_prezzo_unitario <> deleted.det_prezzo_unitario OR 
 			inserted.det_iva <> deleted.det_iva OR 
 			inserted.det_prezzo_listino <> deleted.det_prezzo_listino OR 
 			inserted.det_sconto <> deleted.det_sconto OR 
 			inserted.det_spesespedizione <> deleted.det_spesespedizione OR 
 			inserted.det_speseincasso <> deleted.det_speseincasso OR 
 			inserted.det_spesefisse <> deleted.det_spesefisse OR 
 			inserted.det_spesealtre <> deleted.det_spesealtre OR 
 			inserted.det_spesespedizione_iva <> deleted.det_spesespedizione_iva OR 
 			inserted.det_speseincasso_iva <> deleted.det_speseincasso_iva OR 
 			inserted.det_spesefisse_iva <> deleted.det_spesefisse_iva OR 
 			inserted.det_spesealtre_iva <> deleted.det_spesealtre_iva  OR 
 			inserted.det_tot_colli <> deleted.det_tot_colli OR 
 			inserted.det_tot_volume <> deleted.det_tot_volume OR 
 			inserted.det_tot_peso_lordo <> deleted.det_tot_peso_lordo OR 
 			inserted.det_tot_peso_netto <> deleted.det_tot_peso_netto 
 			) 
  
 OPEN rs 
 FETCH NEXT FROM rs INTO @O_id 
 WHILE @@FETCH_STATUS = 0 
 BEGIN 
 	/* esegue ricalcolo dell'ordine */ 
 	EXEC gsp_totale_ordini @ord_id=@O_id 
 	FETCH NEXT FROM rs INTO @O_id 
 END


;


 CREATE PROCEDURE [dbo].[gsp_totale_ordini] 
 	@ord_id INT  
 AS  
 BEGIN  
 	IF (EXISTS(SELECT det_id  
 			  FROM gtb_dettagli_ord 
				INNER JOIN grel_dettagli_ord_des_value ON gtb_dettagli_ord.det_id = grel_dettagli_ord_des_value.rel_des_dett_ord_id 
				INNER JOIN gtb_dettagli_ord_des ON grel_dettagli_ord_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id 
 			  WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  
 					IsNull(rel_des_valore_it,'') <> '' AND 
 					IsNull(rel_des_valore_it,'') <> '0' AND 
 					IsNull(dod_percentuale_detrazione,0) <> 0 AND 
 					det_ord_id = @ord_id 
 			 )) BEGIN 
 		--ci sono dei descrittori su riga che variano il conteggio della quantità su almeno un dettaglio 
 		--uso un cursore per ogni dettaglio per fare i conti. 
 		DECLARE @det_id INT 
 		DECLARE @det_qta REAL, @detrazione_qta REAL 
 	
 		DECLARE rs CURSOR local FAST_FORWARD FOR  
 		SELECT det_id, det_qta FROM gtb_dettagli_ord WHERE det_ord_id = @ord_id 
 	
 		OPEN rs 
 		FETCH NEXT FROM rs INTO @det_id, @det_qta 
 		WHILE @@FETCH_STATUS = 0 
 		BEGIN 
 			--calcolo quantità in detrazione per ogni singolo dettaglio 
 			SELECT @detrazione_qta = SUM(CAST(IsNull(rel_des_valore_it,'0') AS real) * (CAST(dod_percentuale_detrazione AS real)/100)) 
 				FROM grel_dettagli_ord_des_value INNER JOIN 
 					 gtb_dettagli_ord_des ON grel_dettagli_ord_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id 
 				WHERE rel_des_dett_ord_id = @det_id 
 					  AND IsNull(dod_qta_in_detrazione,0)=1 
 					  AND IsNull(dod_percentuale_detrazione,0)<>0 
 					  AND IsNull(rel_des_valore_it,'') <> ''  
 					  AND IsNull(rel_des_valore_it,'') <> '0' 
   
 			SET @det_qta = @det_qta - @detrazione_qta 
 	
 			--conteggio quantità in base a valori derivati per il singolo dettaglio 
 			UPDATE gtb_dettagli_ord   
 				SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(@det_qta,0),2) ,   
 					det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(@det_qta,0)*ISNULL(det_iva,0)/100,2) ,   
 					det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +   
 											ISNULL(det_speseincasso,0) +  
 											ISNULL(det_spesefisse,0)+  
 											ISNULL(det_spesealtre,0),2) ,   
 					det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 +  
 												ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 +  
 												ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 +  
 												ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)   
 				WHERE det_id=@det_id 
    
 			FETCH NEXT FROM rs INTO @det_id, @det_qta 
 		END 
    
 	END 
 	ELSE  
 	BEGIN  
 		--calcolo normale dei totali per i dettagli dell'ordine 
 		UPDATE gtb_dettagli_ord   
 		SET det_totale= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0),2) ,   
 			det_totale_iva= ROUND(ISNULL(det_prezzo_unitario,0)*ISNULL(det_qta,0)*ISNULL(det_iva,0)/100,2) ,   
 			det_totale_spese= ROUND(ISNULL(det_spesespedizione,0) +   
 									ISNULL(det_speseincasso,0) +  
 									ISNULL(det_spesefisse,0)+  
 									ISNULL(det_spesealtre,0),2) ,   
 			det_totale_spese_iva= ROUND(ISNULL(det_spesespedizione,0)*ISNULL(det_spesespedizione_iva,0)/100 +  
 										ISNULL(det_speseincasso,0)*ISNULL(det_speseincasso_iva,0)/100 +  
 										ISNULL(det_spesefisse,0)*ISNULL(det_spesefisse_iva,0)/100 +  
 										ISNULL(det_spesealtre,0)*ISNULL(det_spesealtre_iva,0)/100,2)   
 		WHERE det_ord_id=@ord_id 
 	END  
 	 
    
 	--calcolo dei totali dei dettagli sulla testata dell'ordine 
 	UPDATE gtb_ordini  
 	SET ord_totale=(SELECT SUM(det_totale) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale IS NOT NULL) ,  
 		ord_totale_iva=(SELECT SUM(det_totale_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_iva IS NOT NULL) ,   
 		ord_det_totale_spese=(SELECT SUM(det_totale_spese) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese IS NOT NULL) ,   
 		ord_det_totale_spese_iva=(SELECT SUM(det_totale_spese_iva) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_totale_spese_iva IS NOT NULL),
 		ord_dett_tot_colli=(SELECT SUM(det_tot_colli) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_tot_colli IS NOT NULL),  
 		ord_dett_tot_peso_netto =(SELECT SUM(det_tot_peso_netto) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_tot_peso_netto IS NOT NULL),
 		ord_dett_tot_peso_lordo=(SELECT SUM(det_tot_peso_lordo) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_tot_peso_lordo IS NOT NULL), 
 		ord_dett_tot_volume=(SELECT SUM(det_tot_volume) FROM gtb_dettagli_ord WHERE det_ord_id=@ord_id AND det_tot_volume IS NOT NULL)
 	WHERE ord_id=@ord_id  
 	 
 	--calcolo dei totali generali dell'ordine 
 	UPDATE gtb_ordini  
 	SET ord_totale_spese=ROUND(ISNULL(ord_spesespedizione,0) +  
 							   ISNULL(ord_speseincasso,0) +  
 							   ISNULL(ord_spesefisse,0) +  
 							   ISNULL(ord_spesealtre,0) +  
 							   ISNULL(ord_det_totale_spese,0),2) ,   
 		ord_totale_spese_iva=ROUND(ISNULL(ord_spesespedizione,0)*ISNULL(ord_spesespedizione_iva,0)/100 +  
 								   ISNULL(ord_speseincasso,0)*ISNULL(ord_speseincasso_iva,0)/100 +  
 								   ISNULL(ord_spesefisse,0)*ISNULL(ord_spesefisse_iva,0)/100 +  
 								   ISNULL(ord_spesealtre,0)*ISNULL(ord_spesealtre_iva,0)/100 +  
 								   ISNULL(ord_det_totale_spese_iva,0),2),
 		ord_totale_colli = IsNull(ord_dett_tot_colli,0) + IsNull(ord_colli,0), 
 		ord_totale_peso_netto = IsNull(ord_dett_tot_peso_netto,0) + IsNull(ord_peso_netto,0), 
 		ord_totale_peso_lordo = IsNull(ord_dett_tot_peso_lordo,0) + IsNull(ord_peso_lordo,0),
 		ord_totale_volume = IsNull(ord_volume,0) + IsNull(ord_dett_tot_volume,0)
 	WHERE ord_id=@ord_id  
 END


;



 CREATE PROCEDURE [dbo].[gsp_totale_shopping_cart]  
 	@sc_id INT  
 AS  
 BEGIN  
 	IF (EXISTS(SELECT dett_id  
 			  FROM gtb_dett_cart 
				INNER JOIN grel_dett_cart_des_value ON gtb_dett_cart.dett_id = grel_dett_cart_des_value.rel_des_dett_cart_id 
				INNER JOIN gtb_dettagli_ord_des ON grel_dett_cart_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id 
 			  WHERE IsNull(gtb_dettagli_ord_des.dod_qta_in_detrazione,0) = 1 AND  
 					IsNull(rel_des_valore_it,'') <> '' AND 
 					IsNull(rel_des_valore_it,'') <> '0' AND 
 					IsNull(dod_percentuale_detrazione,0) <> 0 AND 
 					dett_cart_id = @sc_id 
 			 )) BEGIN 
 		--ci sono dei descrittori su riga che variano il conteggio della quantità su almeno un dettaglio 
 		--uso un cursore per ogni dettaglio per fare i conti. 
 		DECLARE @dett_id INT 
 		DECLARE @dett_qta REAL, @detrazione_qta REAL 
 	 
 		DECLARE rs CURSOR local FAST_FORWARD FOR  
 		SELECT dett_id, dett_qta FROM gtb_dett_cart WHERE dett_cart_id = @sc_id 
 	 
 		OPEN rs 
 		FETCH NEXT FROM rs INTO @dett_id, @dett_qta 
 		WHILE @@FETCH_STATUS = 0 
 		BEGIN 
 			--calcolo quantità in detrazione per ogni singolo dettaglio 
 			SELECT @detrazione_qta = SUM(CAST(IsNull(rel_des_valore_it,'0') AS real) * (CAST(dod_percentuale_detrazione AS real)/100)) 
 				FROM grel_dett_cart_des_value INNER JOIN 
 					 gtb_dettagli_ord_des ON grel_dett_cart_des_value.rel_des_descrittore_id = gtb_dettagli_ord_des.dod_id 
 				WHERE rel_des_dett_cart_id = @dett_id  
 					  AND IsNull(dod_qta_in_detrazione,0)=1 
 					  AND IsNull(dod_percentuale_detrazione,0)<>0 
 					  AND IsNull(rel_des_valore_it,'') <> ''  
 					  AND IsNull(rel_des_valore_it,'') <> '0' 
 	 
 			SET @dett_qta = @dett_qta - @detrazione_qta 
 	 
 			--calcolo normale dei totali per i dettagli della shopping cart 
 			UPDATE gtb_dett_cart 
 				SET dett_totale= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0),2) ,   
 					dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_iva_id),0)/100,2) ,   
 					dett_totale_spese= ROUND(ISNULL(dett_spesespedizione,0) +   
 											 ISNULL(dett_speseincasso,0) +  
 											 ISNULL(dett_spesefisse,0)+  
 											 ISNULL(dett_spesealtre,0),2) ,   
 					dett_totale_spese_iva = ROUND(ISNULL(dett_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesespedizione_iva_id),0)/100 +  
 												  ISNULL(dett_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_speseincasso_iva_id),0)/100 +  
 												  ISNULL(dett_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesefisse_iva_id),0)/100 +  
 												  ISNULL(dett_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesealtre_iva_id),0)/100,2)   
 				WHERE dett_id=@dett_id 
 	 
 			FETCH NEXT FROM rs INTO @dett_id, @dett_qta 
 		END 
 	 
 	END 
 	ELSE  
 	BEGIN 
 		--calcolo dei totali dei dettagli sulla testata della shopping cart 
 		UPDATE gtb_dett_cart   
 		SET dett_totale= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0),2) ,   
 			dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_iva_id),0)/100,2) ,   
 			dett_totale_spese= ROUND(ISNULL(dett_spesespedizione,0) +   
 									 ISNULL(dett_speseincasso,0) +  
 									 ISNULL(dett_spesefisse,0)+  
 									 ISNULL(dett_spesealtre,0),2) ,   
 			dett_totale_spese_iva= ROUND(ISNULL(dett_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesespedizione_iva_id),0)/100 +  
 										ISNULL(dett_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_speseincasso_iva_id),0)/100 +  
 										ISNULL(dett_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesefisse_iva_id),0)/100 +  
 										ISNULL(dett_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = dett_spesealtre_iva_id),0)/100,2)   
 		WHERE dett_cart_id=@sc_id 
 	END  
 	  
 	--calcolo dei totali dei dettagli sulla testata della shopping cart 
 	UPDATE gtb_shopping_cart  
 	SET sc_totale=(SELECT SUM(dett_totale) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale IS NOT NULL) ,  
 		sc_totale_iva=(SELECT SUM(dett_totale_iva) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_iva IS NOT NULL) ,   
 		sc_dett_totale_spese=(SELECT SUM(dett_totale_spese) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_spese IS NOT NULL) ,   
 		sc_dett_totale_spese_iva=(SELECT SUM(dett_totale_spese_iva) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_totale_spese_iva IS NOT NULL),
 		sc_dett_tot_colli =  (SELECT SUM(dett_tot_colli) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_tot_colli IS NOT NULL),
 		sc_dett_tot_peso_netto =  (SELECT SUM(dett_tot_peso_netto) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_tot_peso_netto IS NOT NULL),
 		sc_dett_tot_peso_lordo =  (SELECT SUM(dett_tot_peso_lordo) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_tot_peso_lordo IS NOT NULL),
 		sc_dett_tot_volume =  (SELECT SUM(dett_tot_volume) FROM gtb_dett_cart WHERE dett_cart_id=@sc_id AND dett_tot_volume IS NOT NULL)
 	WHERE sc_id=@sc_id  
 	 
 	--calcolo dei totali generali della shopping cart 
 	UPDATE gtb_shopping_cart  
 	SET sc_totale_spese=ROUND(ISNULL(sc_spesespedizione,0) +  
 							  ISNULL(sc_speseincasso,0) +  
 							  ISNULL(sc_spesefisse,0) +  
 							  ISNULL(sc_spesealtre,0) +  
 							  ISNULL(sc_dett_totale_spese,0),2) ,   
 		sc_totale_spese_iva=ROUND(ISNULL(sc_spesespedizione,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesespedizione_iva_id),0)/100 +  
 								  ISNULL(sc_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_speseincasso_iva_id),0)/100 +  
 								  ISNULL(sc_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesefisse_iva_id),0)/100 +  
 								  ISNULL(sc_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesealtre_iva_id),0)/100 +  
 								  ISNULL(sc_dett_totale_spese_iva,0),2),
 		sc_totale_colli = IsNull(sc_dett_tot_colli,0) + IsNull(sc_colli,0), 
 		sc_totale_peso_netto = IsNull(sc_dett_tot_peso_netto,0) + IsNull(sc_peso_netto,0), 
 		sc_totale_peso_lordo = IsNull(sc_dett_tot_peso_lordo,0) + IsNull(sc_peso_lordo,0),
 		sc_totale_volume = IsNull(sc_volume,0) + IsNull(sc_dett_tot_volume,0)
 	WHERE sc_id=@sc_id  
 END


;



 CREATE TRIGGER [dbo].[gtb_dett_cart_update] 
 ON [dbo].[gtb_dett_cart] 
 AFTER UPDATE 
 AS 
 DECLARE @s_id INT 
 /* 
 apre recordset delle shopping cart alle quali è stato modificato un dettaglio 
 in almeno uno dei campi che concorrono al calcolo dei totali 
 */ 
 DECLARE rs CURSOR local FAST_FORWARD FOR  
 	SELECT DISTINCT inserted.dett_cart_id FROM 
 	inserted INNER JOIN deleted ON  
 		inserted.dett_id = deleted.dett_id 
 		AND (  
 			inserted.dett_art_var_id <> deleted.dett_art_var_id OR 
 			inserted.dett_qta <> deleted.dett_qta OR 
 			inserted.dett_prezzo_unitario <> deleted.dett_prezzo_unitario OR 
 			inserted.dett_iva_id <> deleted.dett_iva_id OR 
 			inserted.dett_prezzo_listino <> deleted.dett_prezzo_listino OR 
 			inserted.dett_sconto <> deleted.dett_sconto OR 
 			inserted.dett_spesespedizione <> deleted.dett_spesespedizione OR 
 			inserted.dett_speseincasso <> deleted.dett_speseincasso OR 
 			inserted.dett_spesefisse <> deleted.dett_spesefisse OR 
 			inserted.dett_spesealtre <> deleted.dett_spesealtre OR 
 			inserted.dett_spesespedizione_iva_id <> deleted.dett_spesespedizione_iva_id OR 
 			inserted.dett_speseincasso_iva_id <> deleted.dett_speseincasso_iva_id OR 
 			inserted.dett_spesefisse_iva_id <> deleted.dett_spesefisse_iva_id OR 
 			inserted.dett_spesealtre_iva_id <> deleted.dett_spesealtre_iva_id OR
 			inserted.dett_tot_colli <> deleted.dett_tot_colli OR 
 			inserted.dett_tot_peso_netto <> deleted.dett_tot_peso_netto OR 
 			inserted.dett_tot_peso_lordo <> deleted.dett_tot_peso_lordo OR 
 			inserted.dett_tot_volume <> deleted.dett_tot_volume
 			) 
  
 OPEN rs 
 FETCH NEXT FROM rs INTO @s_id 
 WHILE @@FETCH_STATUS = 0 
 BEGIN 
 	/* esegue ricalcolo della shopping cart */ 
 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id 
 	FETCH NEXT FROM rs INTO @s_id 
 END

;




 CREATE TRIGGER [dbo].[gtb_shopping_cart_update] 
 ON [dbo].[gtb_shopping_cart] 
 AFTER UPDATE 
 AS 
  
 DECLARE @s_id INT 
 /* 
 apre recordset delle shopping cart modificate 
 in almeno uno dei campi che concorrono al calcolo dei totali 
 */ 
 DECLARE rs CURSOR local FAST_FORWARD FOR  
 	SELECT DISTINCT inserted.sc_id FROM 
 	inserted INNER JOIN deleted ON  
 		inserted.sc_id = deleted.sc_id 
 		AND (  
 			inserted.sc_spesespedizione <> deleted.sc_spesespedizione OR 
 			inserted.sc_speseincasso <> deleted.sc_speseincasso OR 
 			inserted.sc_spesefisse <> deleted.sc_spesefisse OR 
 			inserted.sc_spesealtre <> deleted.sc_spesealtre OR 
 			inserted.sc_spesespedizione_iva_id <> deleted.sc_spesespedizione_iva_id OR 
 			inserted.sc_speseincasso_iva_id <> deleted.sc_speseincasso_iva_id OR 
 			inserted.sc_spesefisse_iva_id <> deleted.sc_spesefisse_iva_id OR 
 			inserted.sc_spesealtre_iva_id <> deleted.sc_spesealtre_iva_id  OR 
 			inserted.sc_colli <> deleted.sc_colli  OR 
 			inserted.sc_peso_netto <> deleted.sc_peso_netto  OR 
 			inserted.sc_peso_lordo <> deleted.sc_peso_lordo  OR 
 			inserted.sc_volume <> deleted.sc_volume 
 			) 
 OPEN rs 
 FETCH NEXT FROM rs INTO @s_id 
 WHILE @@FETCH_STATUS = 0 
 BEGIN 
 	/* esegue ricalcolo della shopping cart */ 
 	EXEC dbo.gsp_totale_shopping_cart @sc_id=@s_id 
 	FETCH NEXT FROM rs INTO @s_id 
 END

;

 CREATE TRIGGER [dbo].[gtb_ordini_update] 
 ON [dbo].[gtb_ordini] 
 AFTER UPDATE 
 AS 
  
 DECLARE @O_id INT 
 /* 
 apre recordset degli ordini modificati 
 in almeno uno dei campi che concorrono al calcolo dei totali 
 */ 
 DECLARE rs CURSOR local FAST_FORWARD FOR  
 	SELECT DISTINCT inserted.ord_id FROM 
 	inserted INNER JOIN deleted ON  
 		inserted.ord_id = deleted.ord_id 
 		AND (  
 			inserted.ord_spesespedizione <> deleted.ord_spesespedizione OR 
 			inserted.ord_speseincasso <> deleted.ord_speseincasso OR 
 			inserted.ord_spesefisse <> deleted.ord_spesefisse OR 
 			inserted.ord_spesealtre <> deleted.ord_spesealtre OR 
 			inserted.ord_spesespedizione_iva <> deleted.ord_spesespedizione_iva OR 
 			inserted.ord_speseincasso_iva <> deleted.ord_speseincasso_iva OR 
 			inserted.ord_spesefisse_iva <> deleted.ord_spesefisse_iva OR 
 			inserted.ord_spesealtre_iva <> deleted.ord_spesealtre_iva OR
 			inserted.ord_colli <> deleted.ord_colli OR 
 			inserted.ord_peso_netto <> deleted.ord_peso_netto OR 
 			inserted.ord_peso_lordo <> deleted.ord_peso_lordo OR 
 			inserted.ord_volume <> deleted.ord_volume 
 			) 
 OPEN rs 
 FETCH NEXT FROM rs INTO @O_id 
 WHILE @@FETCH_STATUS = 0 
 BEGIN 
 	/* esegue ricalcolo dell'ordine */ 
 	EXEC gsp_totale_ordini @ord_id=@O_id 
 	FETCH NEXT FROM rs INTO @O_id 
 END

;


