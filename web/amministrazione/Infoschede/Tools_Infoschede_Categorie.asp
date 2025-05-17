<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<!--#INCLUDE FILE="../library/ClassPhotoGallery.asp" -->
<%
'.................................................................................................
'GESTIONE DEI MODELLI (CATEGORIE FILTRATE)
'.................................................................................................
dim cat_modelli
set cat_modelli = New objCategorie
with cat_modelli
	.tabella = "gtb_tipologie"
	.prefisso = "tip"
	
    'gestione delle relazioni con le tabelle categorizzate
    CALL .AddRelazione("gtb_articoli", "art_id", "art_tipologia_id", true)
	
	.tabellaRelCaratteristiche = "gtb_tip_ctech"
	.chiaveEsternaRelCaratteristiche = "rct_tipologia_id"
	.ordineRelCaratteristiche = "rct_ordine"
	.idCarRelCaratteristiche = "rct_ctech_id"
	
	.tabellaCaratteristiche = "gtb_carattech"
	.idCaratteristiche = "ct_id"
	.nomeCaratteristiche = "ct_nome_it"
	.tipoCaratteristiche = "ct_tipo"
	.tabellaRelCorCaratteristiche = "grel_art_ctech"
	.idArtRelCorCaratteristiche = "rel_art_id"
	.idCarRelCorCaratteristiche = "rel_ctech_id"
	
	.isB2B = true
	.GestioneCategorieMiste = false
	
	'abilitazione indice e contenuti
	.Index = Index
end with


dim cat_ricambi
set cat_ricambi = New objCategorie
with cat_ricambi
	.tabella = "gtb_tipologie"
	.prefisso = "tip"
	
    'gestione delle relazioni con le tabelle categorizzate
    CALL .AddRelazione("gtb_articoli", "art_id", "art_tipologia_id", true)
	
	.tabellaRelCaratteristiche = "gtb_tip_ctech"
	.chiaveEsternaRelCaratteristiche = "rct_tipologia_id"
	.ordineRelCaratteristiche = "rct_ordine"
	.idCarRelCaratteristiche = "rct_ctech_id"
	
	.tabellaCaratteristiche = "gtb_carattech"
	.idCaratteristiche = "ct_id"
	.nomeCaratteristiche = "ct_nome_it"
	.tipoCaratteristiche = "ct_tipo"
	.tabellaRelCorCaratteristiche = "grel_art_ctech"
	.idArtRelCorCaratteristiche = "rel_art_id"
	.idCarRelCorCaratteristiche = "rel_ctech_id"
	
	.isB2B = true
	.GestioneCategorieMiste = false
	
	'abilitazione indice e contenuti
	.Index = Index
end with

dim cat_articoli
set cat_articoli = New objCategorie
with cat_articoli
	.tabella = "gtb_tipologie"
	.prefisso = "tip"
	
    'gestione delle relazioni con le tabelle categorizzate
    CALL .AddRelazione("gtb_articoli", "art_id", "art_tipologia_id", true)
	
	.tabellaRelCaratteristiche = "gtb_tip_ctech"
	.chiaveEsternaRelCaratteristiche = "rct_tipologia_id"
	.ordineRelCaratteristiche = "rct_ordine"
	.idCarRelCaratteristiche = "rct_ctech_id"
	
	.tabellaCaratteristiche = "gtb_carattech"
	.idCaratteristiche = "ct_id"
	.nomeCaratteristiche = "ct_nome_it"
	.tipoCaratteristiche = "ct_tipo"
	.tabellaRelCorCaratteristiche = "grel_art_ctech"
	.idArtRelCorCaratteristiche = "rel_art_id"
	.idCarRelCorCaratteristiche = "rel_ctech_id"
	
	.isB2B = true
	.GestioneCategorieMiste = false
	
	'abilitazione indice e contenuti
	.Index = Index
end with



function GetCatList(conn, codCat)
    dim sql
    sql = "SELECT tip_id FROM gtb_tipologie WHERE tip_codice LIKE '" & codCat & "'"
	GetCatList = GetValueList(conn, NULL, sql)
end function


'imposta filtro categorie in base al parametro indicato dalla sessione
cat_modelli.filtroCategorieBase = GetCatList(cat_modelli.conn, CODICE_CAT_MODELLI)
'cat_modelli.IsB2B = false

cat_ricambi.filtroCategorieBase = GetCatList(cat_ricambi.conn, CODICE_CAT_RICAMBI)


'response.write cat_ricambi.QueryElenco(false, "")
'response.end


%>