<!--#INCLUDE FILE="../library/Tools4Admin.asp" -->
<!--#INCLUDE FILE="../library/Categorie/ClassCategorie.asp" -->
<!--#INCLUDE FILE="../library/IndexContent/Tools_IndexContent.asp" -->
<% 
dim categorie
set categorie = New objCategorie
with categorie
	
	.tabella = "mtb_documenti_categorie"
	.prefisso = "catC"
	
    'gestione delle relazioni con le tabelle categorizzate
    CALL .AddRelazione("mtb_documenti", "doc_id", "doc_categoria_id", true)

	.abilitaFoto = true
	.abilitaLogo = false
	.abilitaDescrittori = true
	.tabellaCaratteristiche = "mtb_carattech"
	.idCaratteristiche = "ct_id"
	.tabellaRelCaratteristiche = "mrel_categ_ctech"
	.idCarRelCaratteristiche = "rcc_ctech_id"
	.chiaveEsternaRelCaratteristiche = "rcc_categoria_id"
	.ordineRelCaratteristiche = "rcc_ordine"
	.nomeCaratteristiche = "ct_nome_it"
	.tipoCaratteristiche = "ct_tipo"
	
	.tabellaRelCorCaratteristiche = "mrel_doc_ctech"
	.idArtRelCorCaratteristiche = "rdc_doc_id"
	.idCarRelCorCaratteristiche = "rdc_ctech_id"
	
	.GestioneCategorieMiste = true
	
		
	'abilitazione indice e contenuti
	.Index = Index
end with

%>