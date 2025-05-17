/*stored procedure per il ricalcolo dei totali della shoppingcart*/

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
 					dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(@dett_qta,0)*ISNULL(dett_iva_valore,0)/100,2) ,   
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
 			dett_totale_iva= ROUND(ISNULL(dett_prezzo_unitario,0)*ISNULL(dett_qta,0)*ISNULL(dett_iva_valore,0)/100,2) ,   
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
	DECLARE @sc_riv_id INT
	SELECT @sc_riv_id = sc_riv_id FROM gtb_shopping_cart WHERE sc_id=@sc_id
	DECLARE @riv_civa_id INT = (SELECT TOP 1 ei_civa_riv_id FROM gtb_eccezioni_iva WHERE ei_riv_id = @sc_riv_id AND ei_inizio_validita <= GETDATE() AND GETDATE() <= ei_fine_validita)
	IF (@riv_civa_id IS NULL)BEGIN
		SELECT @riv_civa_id = ISNULL(riv_civa_id,0) FROM gtb_rivenditori WHERE riv_id = @sc_riv_id
	END
 	UPDATE gtb_shopping_cart  
 	SET sc_totale_spese=ROUND(ISNULL(sc_spesespedizione,0) +  
 							  ISNULL(sc_speseincasso,0) +  
 							  ISNULL(sc_spesefisse,0) +  
 							  ISNULL(sc_spesealtre,0) +  
 							  ISNULL(sc_dett_totale_spese,0),2) ,   
 		sc_totale_spese_iva=ROUND(ISNULL(sc_spesespedizione,0)*ISNULL((SELECT TOP 1 civa_valore FROM grel_civa WHERE
																	   civa_riv_id = @riv_civa_id AND 
																	   civa_art_id = ISNULL(sc_spesespedizione_iva_id,0)),0)/100 +  
 								  ISNULL(sc_speseincasso,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_speseincasso_iva_id),0)/100 +  
 								  ISNULL(sc_spesefisse,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesefisse_iva_id),0)/100 +  
 								  ISNULL(sc_spesealtre,0)*ISNULL((SELECT iva_valore FROM gtb_iva WHERE iva_id = sc_spesealtre_iva_id),0)/100 +  
 								  ISNULL(sc_dett_totale_spese_iva,0),2),
 		sc_totale_colli = IsNull(sc_dett_tot_colli,0) + IsNull(sc_colli,0), 
 		sc_totale_peso_netto = IsNull(sc_dett_tot_peso_netto,0) + IsNull(sc_peso_netto,0), 
 		sc_totale_peso_lordo = IsNull(sc_dett_tot_peso_lordo,0) + IsNull(sc_peso_lordo,0),
 		sc_totale_volume = IsNull(sc_volume,0) + IsNull(sc_dett_tot_volume,0)
 	WHERE sc_id=@sc_id  
END;


/*stored procedure per il calcolo dell'iva*/

CREATE PROCEDURE [dbo].[gsp_calcola_iva] 
	@dett_id INT 
AS 
BEGIN 
 	DECLARE @dett_cart_id INT
	DECLARE @dett_art_var_id INT
	DECLARE @sc_riv_id INT
	SELECT	@dett_cart_id = dett_cart_id, 
			@dett_art_var_id = dett_art_var_id ,
			@sc_riv_id = sc_riv_id
		FROM gtb_dett_cart INNER JOIN gtb_shopping_cart ON gtb_dett_cart.dett_cart_id = gtb_shopping_cart.sc_id
		WHERE dett_ID = @dett_id
	
	DECLARE @riv_civa_id INT = (SELECT TOP 1 ei_civa_riv_id FROM gtb_eccezioni_iva WHERE ei_riv_id = @sc_riv_id AND ei_inizio_validita <= GETDATE() AND GETDATE() <= ei_fine_validita)
	IF (@riv_civa_id IS NULL)BEGIN
		SELECT @riv_civa_id = ISNULL(riv_civa_id,0) FROM gtb_rivenditori WHERE riv_id = @sc_riv_id
	END

	DECLARE @art_civa_id INT = (SELECT ISNULL(art_civa_id,0) AS art_civa_id FROM gtb_articoli INNER JOIN grel_art_valori ON gtb_articoli.art_id = grel_art_valori.rel_art_id WHERE rel_id = @dett_art_var_id) 
	
	DECLARE @civa_valore REAL = NULL
	SELECT TOP 1 @civa_valore = civa_valore FROM grel_civa WHERE civa_art_id = @art_civa_id AND civa_riv_id = @riv_civa_id
	
	IF (@civa_valore IS NOT NULL)BEGIN		
		UPDATE gtb_dett_cart 
		SET dett_iva_valore = @civa_valore 
		WHERE dett_ID = @dett_id 
	END 
END;