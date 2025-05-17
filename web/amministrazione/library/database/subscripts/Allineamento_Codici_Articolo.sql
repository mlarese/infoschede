/*
	allinea i codici articolo, portandoli prima a 13 cifre e poi a 6 che è la nuova lunghezza richiesta
*/

-------------------------------trasformazione dei codici da 5 a 13 cifre-------------------------------------
UPDATE gtb_articoli
	SET art_cod_int = ('80216960' + art_cod_int)
	WHERE (art_tipologia_id = '1239' OR art_tipologia_id = '1240') AND LEN(art_cod_int) = 5;

UPDATE gItb_articoli
	SET Iart_x_cod_int = ('80216960' + Iart_x_cod_int)
	WHERE LEN(Iart_x_cod_int) = 5;

UPDATE grel_art_valori
	SET rel_cod_int = ('80216960' + rel_cod_int)
	WHERE rel_art_id in (SELECT art_id FROM gtb_articoli WHERE (art_tipologia_id = '1239' OR art_tipologia_id = '1240')) 
	AND LEN(rel_cod_int) = 5;

-------------------------------trasformazione dei codici da 13 a 6 cifre-------------------------------------
UPDATE gtb_articoli
	SET art_cod_int = SUBSTRING(art_cod_int, 8, 6)
	WHERE (art_tipologia_id = '1239' OR art_tipologia_id = '1240') AND LEN(art_cod_int) = 13;

UPDATE gItb_articoli
	SET Iart_x_cod_int = SUBSTRING(Iart_x_cod_int, 8, 6)
	WHERE LEN(Iart_x_cod_int) = 13;

UPDATE grel_art_valori
	SET rel_cod_int = SUBSTRING(rel_cod_int, 8, 6)
	WHERE rel_art_id in (SELECT art_id FROM gtb_articoli WHERE (art_tipologia_id = '1239' OR art_tipologia_id = '1240')) 
	AND LEN(rel_cod_int) = 13;