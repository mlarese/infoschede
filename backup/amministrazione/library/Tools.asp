<!--#INCLUDE FILE="Tools4Debug.asp"-->
<%
'.................................................................................................
'.................................................................................................
'.................................................................................................
'COSTANTI GENERALI
'.................................................................................................
'.................................................................................................
'definizione lingue
const LINGUA_ITALIANO	= "it"
const LINGUA_INGLESE 	= "en"
const LINGUA_SPAGNOLO 	= "es"
const LINGUA_TEDESCO 	= "de"
const LINGUA_FRANCESE 	= "fr"
const LINGUA_RUSSO		= "ru"
const LINGUA_CINESE		= "cn"
const LINGUA_PORTOGHESE	= "pt"
dim LINGUE_CODICI, LINGUE_NOMI, LINGUE_NAMES, LINGUE_URLS
LINGUE_CODICI =  Array(LINGUA_ITALIANO, LINGUA_INGLESE, LINGUA_FRANCESE, LINGUA_TEDESCO, LINGUA_SPAGNOLO, LINGUA_RUSSO, LINGUA_CINESE, LINGUA_PORTOGHESE)
LINGUE_NOMI =  Array("ITALIANA", "INGLESE", "FRANCESE", "TEDESCA", "SPAGNOLA", "RUSSA", "CINESE", "PORTOGHESE")
LINGUE_NAMES = Array("Italiano", "English", "Fran&ccedil;ais", "Deutsch", "Espa&ntilde;ol", "Pусский", "中文", "Português")
LINGUE_URLS = Array("Italiano", "English", "Francais", "Deutsch", "Espanol", "Russian", "Chinese", "Portuguese")

'definizioni indici applicazioni del framework
const NEXTPASSPORT 							= 1			'NEXT-passport [gestione utenti]
const NEXTWEB 								= 2			'NEXT-web [gestione grafica e contenuti]
const NEXTCOM 								= 3			'NEXT-com [gestione comunicazioni]
'		""									""			 NEXT-doc+ [comunicazioni & documenti]
const NEXTNEWS 								= 4			'NEXT-news [gestione news]
const NEXTLINK 								= 5			'NEXT-link [gestione link utili]
const NEXTMENU 								= 6			'NEXT-menu [gestione menu' e ricette]
const NEXTFLAT 								= 7			'NEXT-flat [gestione appartamenti turistici]
const NEXTMEMO 								= 8			'NEXT-memo [gestione pubblicazione documenti]
const NEXTBANNER							= 9			'NEXT-banner [gestione banners pubblicitari]
const NEXTCLUB								= 10		'NEXT-club [gestione associati]
const NEXTBOOKING							= 11		'NEXT-booking [gestione prenotazioni]
const NEXTGUESTBOOK							= 12		'NEXT-guestbook [gestione guestbook]
const NEXTCONTRACT							= 13		'NEXT-contract [gestione bandi ed appalti]
const NEXTFAQ								= 14		'NEXT-f.a.q. [gestione frequently asked questions]
const NEXTTEAM								= 15		'NEXT-team [gestione organigramma aziendale]
const NEXTBOOKINGPORTALE					= 16		'NEXT-booking portal [gestione portale di prenotazione]
const NEXTFLATPORTAL						= 17		'NEXT-flat portal [Gestione appartamenti turistici]
const NEXTREALESTATE						= 18		'NEXT-realestate [gestione immobili]
const NEXTB2B								= 19		'NEXT-b2b [gestione prodotti, magazzino e vendita]
const NEXTSCHOOL							= 20		'NEXT-school [gestione struttura di una scuola]
const NEXTCONGRESS							= 21		'NEXT-congress [gestione congressi e prenotazioni]
const NEXTB2B_IMPORT						= 22		'NEXT-b2b integration [Import, export ed integrazione dati]
const NEXTB2B_MAILING						= 23		'NEXT-b2b mailing [Statistiche ed elaborazioni per mailing list]
const NEXTTRAVEL							= 24		'NEXT-travel [gestione agenzia di viaggi]
const NEXTWEB4								= 25		'NEXT-web 4.0 [gestione grafica e contenuti]
const NEXTWEB5								= 26		'NEXT-web 5.0 [gestione grafica e contenuti accessibili]
const NEXTGALLERY							= 27		'NEXT-gallery [gestione gallerie di immagini]
const NEXTCONTENT							= 27		'NEXT-content [gestione contenuti]
const NEXTBOOKING2							= 28		'NEXT-booking 2.0 [gestione prenotazioni]
const NEXTINFO								= 29		'NEXT-info [gestione informazioni ed eventi]
const NEXTBOOKING3							= 30		'NEXT-booking 3.0 [gestione prenotazioni]
const NEXTTOUR                              = 31        'NEXT-tour [gestione pacchetti turistici]
const NEXTBANNER2							= 32		'NEXT-banner 2.0 [gestione banners pubblicitari]
const NEXTCOMMENT							= 33		'NEXT-comment 1.0 [Gestione commenti degli utenti]
const NEXTTRAVEL2							= 34		'NEXT-travel 2.0 [gestione viaggi, pacchetti turistici ed escursioni]
const NEXTINFO_EXPORT						= 35		'NEXT-info export [export dati informativi]
const NEXTMEMO2 							= 36		'NEXT-memo 2.0 [gestione pubblicazione documenti]
const NEXTTOUR_LOGACCESSI 					= 37		'LOG Accessi della Basilica di San Marco
const INFOSCHEDE		 					= 38		'Infoschede, applicativo di infoschede.it
const COMMESSE								= 39		'Gestione commesse Combinario
const APPMEDICI								= 40		'Gestione applicativo Medici

'definizioni indici applicazioni verticali integrate nel framework
const PANIZZI 								= 100		'
const APT_ADMIN 							= 101		'APT - Amministrazione dati dei portali
const APT_FOTO 								= 102		'APT - Archivio immagini
const APT_QUALITA 							= 103		'APT - Controllo qualita'
const APT_BUSSOLA 							= 104		'APT - Estrazione dati Bussola
const APT_LEO 								= 105		'APT - LEO, la rivista di Venezia
const APT_PRENOTAZIONIEVENTI 				= 106		'APT - Gestione prenotazioni eventi
const APT_VENICECARD 						= 107		'VeniceCard - Gestione prenotazioni
const APT_VENEZIASI 						= 108		'VeneziaSi - Gestione prenotazioni
const APT_VILLAWIDMANN 						= 109		'APT - Gestione Villa Widmann
const APT_PRENOTAZIONITURIVE 				= 110		'TU.RI.VE. - Prenotazioni alberghiere
const APT_PRENOTAZIONIUNINDUSTRIA 			= 111		'UNINDUSTRIA - Prenotazioni alberghiere
const APT_MAGAZZINOCENTRALE 				= 112		'APT - Magazzino centrale
const APT_MAGAZZINOOMAGGISTICA 				= 113		'APT - Magazzino omaggistica
const APT_MAGAZZINOSPEDIZIONI 				= 114		'APT - Magazzino spedizioni
const APT_PRESENZE 							= 115		'APT - Presenze
const APT_DISTRIBUZIONE_PRODOTTI			= 116		'APT - Gestione distribuzione prodotti
const BOOKINGS								= 117		'bookings.net [import dati]
const PROVINCIA_TURISMO						= 118		'Import dati turismo                                                            (ABBANDONATO)
const AGENZIE								= 119		'Agenzie
const TIME_TABLE							= 120		'TimeTable [gestione orari di partenza e arrivo]
const EUROPE								= 121		'Europ Assistance [gestione Viaggi Assicurati])
const CENSIMENTOIMMOBILI					= 122		'Censimento immobili
const APT_CIRCOLARI							= 123		'APT - Circolari interne
const ISEVENEZIA							= 124		'isevenezia.it Area amministrativa vecchio sito
const PAYPAL								= 125		'PAYPAL - parametri di accesso
const APT_DISPO_HOTEL						= 126		'APT - Gestione disponibilità hotel
const B2B_CONSUMABILI						= 127		'gestione consumabili e compatibili per stampanti
const APT_STATISTICHE_IAT					= 128		'APT - Gestione statistiche regione-IAT'
const BANCA_KEYCLIENT						= 129		'KeyClient - pagamento attraverso la banca
const PAYPAL_2_0							= 130		'PAYPAL - parametri di accesso - versione 2.0
const BANCA_VIRTUALPAY						= 131		'VirtualPay - pagamento attraverso la banca

'definizioni indici applicazioni verticali turismo provincia
const TURISMO_ASSOCIAZIONI  				= 203		'Assessorato al turismo [gestione associazioni di categoria]
const TURISMO_MODULI_OL 					= 204		'Assessorato al turismo [gestione moduli on-line]
const TURISMO_ALBERGHI 						= 206		'Strutture ricettive [Alberghi]
const TURISMO_CAMPEGGI 						= 207		'Strutture ricettive [Campeggi]
const TURISMO_AFFITTACAMERE 				= 209		'Strutture ricettive [Affittacamere]
const TURISMO_UA_CLASSIFICATE 				= 210		'Strutture ricettive [Unit&agrave; abitative classificate]
const TURISMO_GUIDE_TURISTICHE 				= 211		'Professioni turistiche [Guide turistiche]
const TURISMO_ACCOMPAGNATORI_TURISTICI 		= 212		'Professioni turistiche [Accompagnatori turistici]
const TURISMO_GUIDE_NATURALISTICHE 			= 216		'Professioni turistiche [Guide naturalistico-ambientali]
const TURISMO_ANIMATORI_TURISTICI 			= 215		'Professioni turistiche [Animatori turistici]
const TURISMO_RESIDENCE 					= 217		'Strutture ricettive [Residence]
const TURISMO_RICETTIVITA_SOCIALI 			= 218		'Strutture ricettive [Ricettivit&agrave; sociali]
const TURISMO_B_B 							= 219		'Strutture ricettive [Bed & breakfast]
const TURISMO_UA_NCLASSIFICATE 				= 220		'Strutture ricettive [Unit&agrave; abitative non classificate]
const TURISMO_FORESTERIE 					= 221		'Strutture ricettive [Foresterie]
const TURISMO_COUNTRY_HOUSE 				= 222		'Strutture ricettive [Country house]
const TURISMO_UA_AGENZIE 					= 223		'Strutture ricettive [Unit&agrave; abitative gestite da agenzie immobiliari]
const TURISMO_COMMON_SEARCH 				= 224		'Assessorato al turismo [ricerca comune]
const TURISMO_DIRETTORI_TECNICI 			= 225		'Professioni turistiche [Direttori tecnici]
const TURISMO_AGENZIE_VIAGGIO 				= 226		'Professioni turistiche [Agenzie di viaggio]
const TURISMO_ACCOMPAGNATORI_AGENZIE 		= 227		'Professioni turistiche [Accompagnatori turistici agenzie]
const TURISMO_STATISTICHE			 		= 228		'Statistiche
const TURISMO_PASSPORT				 		= 229		'Gestione permessi applicativi interni alla provincia                           (ABBANDONATO il 12/12/2007)
const TURISMO_CONGRESSUALE_ALBERGHI			= 230		'Turismo Congressuale [Alberghi]
const TURISMO_CONGRESSUALE_CENTRICONGRESSI	= 231		'Turismo Congressuale [Centri congressi]
const TURISMO_CONGRESSUALE_ALTRESEDI		= 232		'Turismo Congressuale [Altre sedi]
const TURISMO_CONGRESSUALE_SEDISTORICHE		= 233		'Turismo Congressuale [Sedi storiche]
const TURISMO_CONGRESSUALE_AGENZIE			= 234		'Turismo Congressuale [Agenzie con reparto congressuale]
const TURISMO_CONGRESSUALE_ORGANIZZATORI	= 235		'Turismo Congressuale [Organizzazioni professionali]
const TURISMO_CONGRESSUALE_SERVIZITRADUZIONE= 236		'Turismo Congressuale [Imprese di servizi traduzione ed interpretariato]
const TURISMO_CONGRESSUALE_SERVIZITECNICI	= 237		'Turismo Congressuale [Imprese di servizi tecnici]
const TURISMO_CONGRESSUALE_IMPRESEASSISTENZA= 238		'Turismo Congressuale [Imprese di assistenza congressuale]
const TURISMO_CONGRESSUALE_SERVIZICATERING	= 239		'Turismo Congressuale [Imprese servizi di catering]
const TURISMO_CONGRESSUALE_SERVIZITRASPORTO	= 240		'Turismo Congressuale [Imprese servizi di trasporto]
const TURISMO_CONGRESSUALE_ALLESTITORI		= 241		'Turismo Congressuale [Imprese servizi di allestimento]
const TURISMO_AGRITURISMI 					= 242		'Strutture ricettive [Agriturismi]
const TURISMO_UA		 					= 243		'Strutture ricettive [Unit&agrave; abitative]

'definizione set di caratteri validi
const EDITOR_BASE_CHARSET 		= "abcdefghijklmnopqrstuvwxyz|!%&/()=?'[]*+-,.;:_<>0123456789@# "
const EMAIL_VALID_CHARSET 		= "abcdefghijklmnopqrstuvwxyz@.-_1234567890"
const PHONE_VALID_CHARSET		= "+0123456789"
const DOMAIN_VALID_CHARSET 		= "abcdefghijklmnopqrstuvwxyz"
const LOGIN_VALID_CHARSET 		= "abcdefghijklmnopqrstuvwxyz@.-_1234567890"
const DOCUMENTS_FILES_CHARSET 	= "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
const JS_OBJECTS_NAME_CHARSET   = "abcdefghijklmnopqrstuvwxyz1234567890"
const ALPHANUMERIC_CHARSET 		= "abcdefghijklmnopqrstuvwxyz1234567890"
const NUMERIC_CHARSET 			= "01234564789.,-"
const TAGS_INVALID_CHARSET		= "|!%&/()=?'[]*+-.;:_<>@#""£${}^ç°§\•"
const FILENAME_VALID_CHARSET	= "abcdefghijklmnopqrstuvwxyz.-_1234567890"
const FOLDER_VALID_CHARSET		= "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_abcdefghijklmnopqrstuvwxyz"

'tipi di database gestiti
const DB_SQL 		= "SQL Server"
const DB_SQL_2000	= "08"
const DB_SQL_2005	= "09"
const DB_SQL_2008 	= "10"
const DB_ACCESS 	= "MS Jet"
const DB_UNKNOWN	= "UNKNOWN"

'definizione tipi di layers del next-Web
const LAYER_TEXT 	= 1
const LAYER_IMAGE 	= 2
const LAYER_FLASH 	= 3
const LAYER_OBJECT	= 4
const LAYER_S_TEXT	= 5

'definizione nomi tipi layers del next-Web
const LAYER_TEXT_NAME 	= "layers_text"
const LAYER_IMAGE_NAME 	= "layers_image"
const LAYER_FLASH_NAME 	= "layers_flash"
const LAYER_OBJECT_NAME	= "layers_object"
const LAYER_S_TEXT_NAME	= "layers_text_s"

'definizione tipi di oggetto recuperabili dal FileSystem
const FILE_SYSTEM_FILE = "File"
const FILE_SYSTEM_DIRECTORY = "Directory"

'definizione dei tipi di file e relativi nomi di directory registrate nel filesystem
const FILE_TYPE_IMAGE = "images"
const FILE_TYPE_FLASH = "flash"
const FILE_TYPE_OBJECTS = "objects"
const FILE_TYPE_TEXT = "testi"
const FILE_TYPE_XML = "xml"
const FILE_TYPE_CSS = "css"
const FILE_TYPE_PDF = "pdf"

'definizione tipi di recapito del next-Com
const VAL_NONE 			= 0
const VAL_TELEFONO 		= 1
const VAL_TEL_UFFICIO 	= 2
const VAL_CELLULARE 	= 3
const VAL_TEL_CASA 		= 4
const VAL_FAX 			= 5
const VAL_EMAIL 		= 6
const VAL_URL 			= 7

'definizione tipi per upload file (per estensione si intendono le prime 3 lettere dell'estensione: HTML = HTM)
const EXTENSION_ALLOWED = 	" ASF AVI BMP CAB CHM CSS CSV DOC DOT EML FLA GIF HLP HTM JPE JPG MDB MDE MPE MPG MSG PCX PDF PNG PPS PPT PSD PSP PST PUB RAR RTF SQL SWF TIF TXT VCF WAB WMV XLS XLT XML XSL ZIP ARJ DBF PSD TAR TGZ JPEG "
const EXTENSION_ICONS = 	" ASF AVI BMP CAB CHM CSS CSV DOC DOT EML FLA GIF HLP HTM INI JPE JPG MDB MDE MPE MPG MSG PCX PDF PNG PPS PPT PSD PSP PST PUB RAR REG RTF SQL SWF TIF TXT VCF WAB WMV XLS XLT XML XSL ZIP "
const EXTENSION_IMAGES = 	" BMP GIF JPG JPE TIF PNG JPEG "
const EXTENSION_TEXT = 		" TXT "
const EXTENSION_FLASH = 	" SWF "
const EXTENSION_HTML = 		" HTM HTML TML "
const EXTENSION_FAX = 		" PDF DOC XLS TIF JPG JPE JPEG HTML HTM TML "

'costante che ingloba una parte di codice JavaScript per impostare la barra di stato del link
const ACTIVE_STATUS = " onmouseover=""return(status=this.title);"" onmouseout=""status=''; "" "

'definizione tipi di operatore
const OP_ADMIN = "A"
const OP_USER = "U"

'costante per gestire la provenienza della navigazione dei form di inserimento e modifica degli alberi gerarchici
const FROM_ELENCO = "elenco"
const FROM_ALBERO = "albero"

'costante di filtro testuale
const FILTRO_TEXT_UGUALE 		= 1
const FILTRO_TEXT_FULLTEXT 		= 2
const FILTRO_TEXT_INIZIO 		= 3
const FILTRO_TEXT_FINE 			= 4

'costante di definizione del tipo di user agent
const BROWSER = "contUtenti"
const CRAWLER = "contCrawler"
const UNRECOGNIZED = "contAltro"

'costanti di definizione dei servizi di spedizione messaggi
const MSG_EMAIL = 1
const MSG_SMS = 2
const MSG_FAX = 3

'costanti di definizione tipi mime dei messaggi spediti
const MIME_TEXT = "text/plain"
const MIME_HTML = "text/html"

'costante per definizione handler per cssinlining
const PREMAILER_HANDLER = "nextmail.axd"

'variabile usata nella funzione WriteFileSystemPicker_Input su Tools4Admin.asp per gestire in automatico la crezione delle cartelle per l'upload dei file
dim automaticPathWriteFileSystemPicker_Input

'.................................................................................................
'.................................................................................................
'.................................................................................................
'funzioni per la gestione ed il riconoscimento del database in uso (SQL SERVER o ACCESS)
'.................................................................................................
'.................................................................................................


'.................................................................................................
'..          Funzione che restituisce il tipo di database della connessione
'..          conn		connessione da analizzare
'.................................................................................................
function DB_Type(conn)
	if instr(1, conn.properties("DBMS name"), DB_ACCESS, vbTextCompare)>0 then
		'connessione a database access
		DB_Type = DB_ACCESS
	elseif instr(1, conn.properties("DBMS name"), DB_SQL, vbTextCompare)>0 then
		'connessione a database SQL Server
		DB_Type = DB_SQL
	else
		DB_type = DB_UNKNOWN
	end if
end function


'.................................................................................................
'..		Funzione che restituisce la versione di sql server	
'..		conn		connessione da analizzare
'.................................................................................................
function DB_SQL_version(conn)
	if cIntero(left(conn.properties("DBMS Version"), 2)) = cIntero(DB_SQL_2000) then
		DB_SQL_version = DB_SQL_2000
	elseif cIntero(left(conn.properties("DBMS Version"), 2)) = cIntero(DB_SQL_2005) then
		DB_SQL_version = DB_SQL_2005
	elseif cIntero(left(conn.properties("DBMS Version"), 2)) = cIntero(DB_SQL_2008) then
		DB_SQL_version = DB_SQL_2008
	else
		DB_SQL_version = DB_UNKNOWN
	end if
end function


'.................................................................................................
'..          restituisce la codifica della funzione uCase() per il database
'..          conn		connessione da analizzare
'.................................................................................................
function SQL_Ucase(conn)
	Select case DB_Type(conn)
		case DB_Access
			SQL_Ucase = "UCASE"
		case DB_SQL
			SQL_Ucase = "UPPER"	
		case DB_UNKNOWN
			SQL_Ucase = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la codifica della funzione Now()
'..          conn		connessione da analizzare
'.................................................................................................
function SQL_Now(conn)
	Select case DB_Type(conn)
		case DB_Access
			SQL_Now = "NOW()"
		case DB_SQL
			SQL_Now = "GETDATE()"	
		case DB_UNKNOWN
			SQL_Now = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per creare la condizione booleana vera
'..          conn		connessione da analizzare
'..			 field		campo con cui creare la condizione
'.................................................................................................
function SQL_IsTrue(conn, field)
	Select case DB_Type(conn)
		case DB_Access
			SQL_IsTrue = " " & field & " "
		case DB_SQL
			SQL_IsTrue = " ISNULL(" & field & ",0) = 1 " 
		case DB_UNKNOWN
			SQL_IsTrue = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per verificare se un campo e' nullo
'..          conn		connessione da analizzare
'..			 field		campo con cui creare la condizione
'.................................................................................................
function SQL_IsNULL(conn, field)
	Select case DB_Type(conn)
		case DB_Access
			SQL_IsNULL = " (ISNULL(" & field & ")) "
		case DB_SQL
			SQL_IsNULL = " (" & field & " IS NULL) " 
		case DB_UNKNOWN
			SQL_IsNULL = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per sostituire la funzione ISNULL(field, value) di SQL server
'..          conn		connessione da analizzare
'..			 field		campo con cui creare la condizione
'..			 value		valore da sostituire al valore NULL
'.................................................................................................
function SQL_IfIsNull(conn, field, value)
	Select case DB_Type(conn)
		case DB_Access
			SQL_IfIsNull = " IIF(ISNULL(" & field & "), " & value & ", " & field & ") "
		case DB_SQL
			SQL_IfIsNull = " ISNULL(" & field & ", " & value & ") "
		case DB_UNKNOWN
			SQL_IfIsNull = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per creare la condizione booleana vera
'..          conn		connessione da analizzare
'..			 condition	condizione SQL
'..			 ifTrue		parte di SQL da eseguire se condizione vera
'..			 ifFalse	parte di SQL da eseguire se condizione falsa
'.................................................................................................
function SQL_If(conn, condition, ifTrue, ifFalse)
	Select case DB_Type(conn)
		case DB_Access
			SQL_If = " IIF(" & condition & ", " & ifTrue & ", " & ifFalse & ") "
		case DB_SQL
			SQL_If = " CASE WHEN " & condition & " THEN " & ifTrue & " ELSE " & ifFalse & " END "
		case DB_UNKNOWN
			SQL_If = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per inserire una data
'..          conn		connessione da analizzare
'..			 data		data da inserire
'.................................................................................................
function SQL_date(conn, data)
	Select case DB_Type(conn)
		case DB_Access
			SQL_date = "#" & DateIso(data) & "#"
		case DB_SQL
			SQL_date = "CONVERT(DATETIME, '" & DateIso(data) & " 00:00:00',102)"
		case DB_UNKNOWN
			SQL_date = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per convertire un campo testo in numerico
'..          conn		connessione da analizzare
'..			 campo		colonna da converitre
'.................................................................................................
function SQL_Numeric(conn, campo)
	Select case DB_Type(conn)
		case DB_Access
			SQL_Numeric = "CDbl("& campo &")"
		case DB_SQL
			SQL_Numeric = "CAST(CAST("& campo &" AS NVARCHAR) AS FLOAT)"
		case DB_UNKNOWN
			SQL_Numeric = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per convertire un campo in testo
'..          conn		connessione da analizzare
'..			 campo		colonna da converitre
'.................................................................................................
function SQL_String(conn, campo)
	Select case DB_Type(conn)
		case DB_Access
			SQL_String = "CStr("& campo &")"
		case DB_SQL
			SQL_String = "CAST("& campo &" AS NVARCHAR)"
		case DB_UNKNOWN
			SQL_String = DB_UNKNOWN
	end select
end function


function SQL_UTF8Qualifier(conn)
	Select case DB_Type(conn)
		case DB_SQL
			SQL_UTF8Qualifier = "N"
		case else
			SQL_UTF8Qualifier = ""
	end select
end function

'.................................................................................................
'..          restituisce la parte di codice sql per inserire una data e time
'..          conn		connessione da analizzare
'..			 data		data da inserire
'.................................................................................................
function SQL_DateTime(conn, DateTime)
	Select case DB_Type(conn)
		case DB_Access
			SQL_DateTime = "#" & DateIso(DateTime) & " " & TimeIta(DateTime) & ":00#"
		case DB_SQL
			SQL_DateTime = "CONVERT(DATETIME, '" & DateIso(DateTime) & " " & TimeIta(DateTime) & ":00',102)"
		case DB_UNKNOWN
			SQL_DateTime = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per il confronto tra un campo ed una data
'..          conn				connessione da analizzare
'..			 field				campo con cui effettuare la comparazione
'..			 ComparisonType:	adCompareGreaterThan per >=, adCompareLessThan per <=
'..			 data 				valore con cui confrontare il campo
'.................................................................................................
function SQL_CompareDateTime(conn, field, ComparisonType, data)
	Select case DB_Type(conn)
		case DB_Access
			if ComparisonType = adCompareGreaterThan then		
				'valore campo maggiore di data (estremo inferiore)
				SQL_CompareDateTime = field & " >= #" & DateIso(data) & " 00:00:00 #"
			else
				'valore campo minore di data (estremo superiore)
				SQL_CompareDateTime = field & " <= #" & DateIso(data) & " 23:59:59 #"
			end if
		case DB_SQL
			if ComparisonType = adCompareGreaterThan then
				'valore campo maggiore di data (estremo inferiore)
				SQL_CompareDateTime = field & " >= CONVERT(DATETIME, '" & DateIso(data) & " 00:00:00',102)"
			else
				'valore campo minore di data (estremo superiore)
				SQL_CompareDateTime = field & " <= CONVERT(DATETIME, '" & DateIso(data) & " 23:59:59',102)"
			end if
		case DB_UNKNOWN
			SQL_CompareDateTime = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per il confronto tra un campo ed una data (compresa ora)
'..          conn				connessione da analizzare
'..			 field				campo con cui effettuare la comparazione
'..			 ComparisonType:	adCompareGreaterThan per >=, adCompareLessThan per <=
'..			 data 				valore con cui confrontare il campo
'.................................................................................................
function SQL_CompareDateTimeComplete(conn, field, ComparisonType, data)
	Select case DB_Type(conn)
		case DB_Access
			if ComparisonType = adCompareGreaterThan then		
				'valore campo maggiore di data (estremo inferiore)
				SQL_CompareDateTimeComplete = field & " >= #" & Replace(DateTimeIso(data), ".", ":") & " #"
			else
				'valore campo minore di data (estremo superiore)
				SQL_CompareDateTimeComplete = field & " <= #" & Replace(DateTimeIso(data), ".", ":") & " #"
			end if
		case DB_SQL
			if ComparisonType = adCompareGreaterThan then
				'valore campo maggiore di data (estremo inferiore)
				SQL_CompareDateTimeComplete = field & " >= CONVERT(DATETIME, '" & Replace(DateTimeIso(data), ".", ":") & "',102)"
			else
				'valore campo minore di data (estremo superiore)
				SQL_CompareDateTimeComplete = field & " <= CONVERT(DATETIME, '" & Replace(DateTimeIso(data), ".", ":") & "',102)"
			end if
		case DB_UNKNOWN
			SQL_CompareDateTimeComplete = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per il between di un valore tra due date
'..          conn		connessione da analizzare
'..			 field		campo da utilizzare nel begin
'..			 datefrom	data di partenza
'..			 dateto		data di fine
'.................................................................................................
function SQL_BetweenDate(conn, field, DateFrom, DateTo)
	Select case DB_Type(conn)
		case DB_Access
			SQL_BetweenDate = " " & field & " BETWEEN #" & DateIso(DateFrom) & " 00:00:00 # AND " & _
							  " #" & DateIso(DateTo) & " 23:59:59 # "
		case DB_SQL
			SQL_BetweenDate = " " & field & " BETWEEN CONVERT(DATETIME, '" & DateIso(DateFrom) & " 00:00:00',102) AND " & _
							  " CONVERT(DATETIME, '" & DateIso(DateTo) & " 23:59:59',102) "
		case DB_UNKNOWN
			SQL_BetweenDate = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la parte di codice sql per il between di un valore tra due date comprese di orari precisi
'..          conn		connessione da analizzare
'..			 field		campo da utilizzare nel begin
'..			 DateTimeFrom	data e ora di partenza
'..			 DateTimeTo		data e ora di fine
'.................................................................................................
function SQL_BetweenDateTime(conn, field, DateTimeFrom, DateTimeTo)
	Select case DB_Type(conn)
		case DB_Access
			SQL_BetweenDateTime = " " & field & " BETWEEN #" & DateIso(DateTimeFrom) & " " & TimeIso(DateTimeFrom) & "# AND " & _
							  " #" & DateIso(DateTimeTo) & " " & TimeIso(DateTimeTo) & "# "
		case DB_SQL
			SQL_BetweenDateTime = " " & field & " BETWEEN CONVERT(DATETIME, '" & DateIso(DateTimeFrom) & " " & TimeIso(DateTimeFrom) & "',102) AND " & _
							  " CONVERT(DATETIME, '" & DateIso(DateTimeTo) & " " & TimeIso(DateTimeTo) & "',102) "
		case DB_UNKNOWN
			SQL_BetweenDateTime = DB_UNKNOWN
	end select
end function



'.................................................................................................
'..          sql per la funzione DATEDIFF presente sia in SQL sia in access. non converte le date in ingresso.
'..          conn		connessione da analizzare
'..			 part		parte della data interessata nel calcolo della differenza
'..			 datefrom	data di partenza
'..			 dateto		data di fine
'.................................................................................................
Function SQL_DateDiff(conn, part, dateFrom, dateTo)
	Select case DB_Type(conn)
		case DB_Access
			if LCase(part) = "hh" then
				part = "h"
			end if
			SQL_DateDiff = " DateDiff('"& part &"', "& dateFrom &", "& dateTo &")"
		case DB_SQL
			if LCase(part) = "h" then
				part = "hh"
			end if
			SQL_DateDiff = " DATEDIFF("& part &", "& dateFrom &", "& dateTo &") "
		case DB_UNKNOWN
			SQL_DateDiff = DB_UNKNOWN
	end select
End Function


'.................................................................................................
'.. Giacomo 22/01/2014
'.................................................................................................
Function SQL_MergeDateAndTimeFields(conn, dateField, timeField)
	SQL_MergeDateAndTimeFields = " CONVERT(DATETIME, CAST(CAST("&dateField&" AS DATE) AS NVARCHAR(10))+' '+CAST(CAST("&timeField&" AS TIME) AS NVARCHAR(12)), 102) "
End Function

		
'.................................................................................................
'..          restituisce l'operatore per la concatenazione
'..          conn		connessione da analizzare
'.................................................................................................
function SQL_concat(conn)
	Select case DB_Type(conn)
		case DB_Access
			SQL_concat = " & "
		case DB_SQL
			SQL_concat = " + "
		case DB_UNKNOWN
			SQL_concat = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce la porzione di sql con i campi concatenati da " "
'..          conn		connessione da analizzare
'..          FieldsList Elenco di campi suddivisi da ";"
'.................................................................................................
function SQL_ConcatFields(conn, FieldsList)
    dim List, i, Op
    
    SQL_ConcatFields = ""
    List = split(FieldsList, ";")
    Op = SQL_concat(conn)
    for i = lBound(List) to uBound(List)
        SQL_ConcatFields = SQL_ConcatFields + _
                           IIF(SQL_ConcatFields<>"", Op + "' '" + Op, "") + List(i)
    next
    if SQL_ConcatFields<>"" then
        SQL_ConcatFields = "(" & SQL_ConcatFields & ")"
    end if
end function


'.................................................................................................
'..         compone una stringa sql per la ricerca full-text della stringa richiesta 
'..			nell'elenco di campi selezionati
'..         StrToSearch		    :			Stringa da ricercare
'..			FieldsToSearchInto	:			Elenco dei campi da ricercare (separati da ";")
'.................................................................................................
function SQL_FullTextSearch(StrToSearch, FieldsToSearchInto)
	SQL_FullTextSearch = SQL_TextSearch(StrToSearch, FieldsToSearchInto, true)
end function


'.................................................................................................................
'..         compone una stringa sql per la ricerca testuale della stringa richiesta 
'..			nell'elenco di campi selezionati
'..         StrToSearch		    :			Stringa da ricercare
'..			FieldsToSearchInto	:			Elenco dei campi da ricercare (separati da ";")
'..			FullTextSearch:					Indica se la ricerca deve essere effettuata in modalita' full-text
'.................................................................................................................
function SQL_TextSearch(StrToSearch, FieldsToSearchInto, FullTextSearch)
	if Trim(StrToSearch)<>"" AND Trim(FieldsToSearchInto)<>"" then
		dim SearchBaseSQL, SearchSQL
		dim List, i
		
		if FullTextSearch then
			'compone parti della stringa di ricerca per cercare tutte le parole (full text)
	        List = split(StrToSearch, " ")
	        
			SearchBaseSQL = ""
			for i = lBound(List) to uBound(List)
	            if List(i)<>"" then
	    			SearchBaseSQL = SearchBaseSQL + " <FIELD_NAME> LIKE '%" + ParseSQL(List(i), adChar) + "%' AND "
	            end if
			next
			SearchBaseSQL = "(" + left(SearchBaseSQL, len(SearchBaseSQL) - 4) + ")"
		else
			'compone stringa di confronto
			SearchBaseSQL = SearchBaseSQL + " <FIELD_NAME> LIKE '" + ParseSQL(StrToSearch, adChar) + "' "
		end if
		
		'compone string adi ricerca per ogni campo
		List = split(FieldsToSearchInto, ";")
		SearchSQL = ""
		for i = lBound(List) to uBound(List)
            if List(i)<>"" then
    			SearchSQL = SearchSQL + replace(SearchBaseSQL, "<FIELD_NAME>", List(i)) + " OR "
            end if
		next
		SearchSQL = "(" + left(SearchSQL, len(SearchSQL) - 3) + ")"
	else 
		SearchSQL = ""
	end if
	SQL_TextSearch = SearchSQL
end function


'.................................................................................................
'..          restituisce l'sql per filtrare in base ad un id contenuto in un campo lista
'..          conn		connessione da analizzare
'..          field      campo che contiene la lista in cui cercare
'..          id         id da cercare nella lista.
'.................................................................................................
function SQL_IdListSearch(conn, field, id)
    SQL_IdListSearch = " ( ','" & SQL_concat(conn) & " " & field & " " & SQL_concat(conn) & " ',' LIKE '%," & id & ",%' ) "
end function


'.................................................................................................
'..          restituisce l'sql per ordinare in modo casuale
'..          conn		connessione da analizzare
'..          cmpId		nome del campo id della tabella (deve essere un campo intero)
'.................................................................................................
function SQL_OrderByRandom(conn, cmpId)
	Select case DB_Type(conn)
		case DB_Access
			SQL_OrderByRandom = " ORDER BY Rnd("& cmpId &") "
		case DB_SQL
			SQL_OrderByRandom = " ORDER BY NEWID() "
		case DB_UNKNOWN
			SQL_OrderByRandom = DB_UNKNOWN
	end select
end function


'.................................................................................................
'..          restituisce l'sql per ordinare in base ad un campo booleano
'..          conn		connessione da analizzare
'..          cmpNome	nome del campo booleano
'..          prima		true se prima il 'vero' altrimento false
'.................................................................................................
function SQL_OrderByBoolean(conn, cmpNome, prima)
	SQL_OrderByBoolean = cmpNome
	Select case DB_Type(conn)
		case DB_Access
			SQL_OrderByBoolean = SQL_OrderByBoolean + IIF(prima, "", " DESC")
		case DB_SQL
			SQL_OrderByBoolean = SQL_OrderByBoolean + IIF(prima, " DESC", "")
	end select
end function


'.................................................................................................
'..        restituisce l'sql in input moltiplicato per le lingue attive concatenate tramite concat
'..			sql:			sql da moltiplicare, la lingua va definita col tag <LINGUA>			..
'..			concat:			stringa che concatena l'sql delle varie lingue						..
'.................................................................................................
function SQL_MultiLanguage(sql, concat)
	dim lingua
    for each lingua in Application("LINGUE")
		SQL_MultiLanguage = SQL_MultiLanguage & Replace(sql, "<LINGUA>", lingua) & concat
	next
	SQL_MultiLanguage = Left(SQL_MultiLanguage, Len(SQL_MultiLanguage) - Len(concat))
end function


'.................................................................................................
'..        	restituisce l'operatore condizionale da aggiungere alla query per accodare un filtro sql
'..			sql:			sql da verificare
'..			operator:		operatore sql da aggiungere
'.................................................................................................
function SQL_AddOperator(sql, operator)
	if instr(1, sql, "WHERE", vbTextCompare)>0 then
		SQL_AddOperator = " " & operator & " "
	else
		SQL_AddOperator = " WHERE "
	end if
end function


'.................................................................................................
'..        controlla la stringa sql e la prepara per essere messa in una query sql			    ..
'..			str:			stringa da controllare e convertire									..
'..			typ:			tipo di dati da convertire (il tipo viene definito in ADOVBS.INC)	..
'.................................................................................................
function ParseSQL(byVal str, typ)
	str = str & ""
	select case typ
		case adNumeric
			if isItalian_SO() then
				ParseSQL = replace(str,",",".")
			else
				ParseSQL = str
			end if
		case adChar
			ParseSQL = replace(RTrim(str), "'", "''")
		case else
			ParseSQL = str
	end select
end function


'.................................................................................................
'converte la data in ingresso nel formato del sistema operativo.
'ATTENZIONE: la data in ingresso DEVE essere in formato italiano
'   data:       data da convertire
'.................................................................................................
function ConvertForSave_Date(byVal data)
	dim strData, dateParts
	strData = cString(data)
	if isDate(strData) AND strData<>"" then
		dateParts = split(strData, "/")
		if ubound(dateParts)<2 then
			dateParts = split(strData,".")
			if ubound(dateParts)<2 then
				dateParts = split(strData," ")
			end if
		elseif instr(1, dateParts(2), " ", vbTextCompare)>0 AND _
			   (instr(1, dateParts(2), ".", vbTextCompare)>0 OR instr(1, dateParts(2), ":", vbTextCompare)>0) then
			'dateparts 2 contiene un orario
			dim orarioparts 
			orarioparts = split(dateParts(2), " ")
			dateParts(2) = orarioparts(0)
			dateParts(0) = dateParts(0) + " " + orarioparts(1)
		end if
		if ubound(dateParts)>=2 then
			ConvertForSave_Date = cDate(dateParts(2) & " - " & dateParts(1) & " - " & dateParts(0))
		else
			ConvertForSave_Date = NULL
		end if
	elseif data="NOW" then
		ConvertForSave_Date = NOW()
	elseif data="DATE" then
		ConvertForSave_Date = DATE()
	else
		ConvertForSave_Date = NULL
	end if
end function


'.................................................................................................
'converte l'orario in ingresso nel formato del sistema operativo.
'   orario:       orario da convertire
'.................................................................................................
function ConvertForSave_Time(byval orario)
	dim strTime, timeParts
	strTime = cString(orario)
	if isDate(strTime) AND strTime<>"" then
		timeParts = split(replace(replace(strTime, ".", ":"), " ", ""), ":")
		if ubound(timeParts) > 1 then
			strTime = timeParts(0) & ":" & timeParts(1)
			if ubound(timeParts) >= 2 then
				strTime = strTime & ":" & timeParts(2)
			end if
			ConvertForSave_Time = cDate(strTime)
		else
			ConvertForSave_Time = NULL
		end if
	elseif orario="NOW" then
		ConvertForSave_Time = NOW()
	else
		ConvertForSave_Time = NULL
	end if
	
end function


'.................................................................................................
'converte il numero in ingresso nel formato del sistema operativo (serve per eliminare i problemi tra 
'punti e virgole nel passaggio tra formato italiano ed inglese
'se il valore indicato non e' valido restituisce il valore InvalidValue
'   number:                 valore da convertire in numerico
'   InvalidNumberValue:     valore restituito se il numero da convertire non e' un valore numerico valido
'.................................................................................................
function ConvertForSave_Number(number, invalidNumberValue)
	dim strNumber
	strNumber = cString(number)
	if IsNumeric(strNumber) AND strNumber<>"" then
		ConvertForSave_Number = cReal(strNumber)
	elseif instr(1, number, "NULL", vbTextCompare)>0 then
		ConvertForSave_Number = NULL
	else
		ConvertForSave_Number = invalidNumberValue
	end if
end function


'.................................................................................................
'.................................................................................................
'.................................................................................................
'FUNZIONI SUI FILES
'.................................................................................................
'.................................................................................................
'resituisce la dimensione del file richiesto
function File_Size(filePath)
	dim fso, path, fo
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	if instr(1, filepath, ":", vbTextCompare)<=0 then
		if left(filepath,1)="/" then
			path = Application("IMAGE_PATH") & Application("AZ_ID") & filePath
		else
			path = Application("IMAGE_PATH") & Application("AZ_ID") & "\" & filePath
		end if
	else
		path = filePath
	end if
	if fso.FileExists(path) then
		set fo = fso.GetFile(path)
		File_Size = fo.size
	else
		File_Size = 0
	end if
	set fso = nothing
end function



'resituisce la data di creazione del file richiesto
function File_Date(filePath)
	dim fso, path, fo
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	if instr(1, filepath, ":", vbTextCompare)<=0 then
		if left(filepath,1)="/" then
			path = Application("IMAGE_PATH") & Application("AZ_ID") & filePath
		else
			path = Application("IMAGE_PATH") & Application("AZ_ID") & "\" & filePath
		end if
	else
		path = filePath
	end if
	if fso.FileExists(path) then
		set fo = fso.GetFile(path)
		File_Date = fo.DateCreated
	else
		File_Date = ""
	end if
	set fso = nothing
end function


function File_Copy(filePathToCopy, pathTo)
	dim fso, path
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	path = Left(pathTo, inStrRev(pathTo, "\"))
	if NOT fso.FolderExists(path) then
		fso.CreateFolder(path)
	end if
	'response.write "from:"&filePathToCopy & "--->to:" & path
	fso.CopyFile filePathToCopy, pathTo, true
	set fso = nothing
end function


'.................................................................................................
'	restituisce il corretto evento onclick sul file a seconda se e' una immagine o un altro file
'	FileUrl:		url completo del file da aprire sul click
'.................................................................................................
function File_OpenInNewWindow( FileUrl )
	
	if instr(1, EXTENSION_IMAGES, File_Extension( FileUrl ), vbTextCompare)>0 then
		'immagine
		File_OpenInNewWindow = "OpenSmartImage('" + JsEncode(FileUrl, "'") + "')"
	else
		'altro file
		File_OpenInNewWindow = "OpenWindow('" + JsEncode(FileUrl, "'") + "', '', '')"
	end if
	
end function


'.................................................................................................
'..		formatta il numero passato in dimensione in BYTE, KBYTE O MBYTE a seconda della dimensione
'..		size:		dimensione del file
'.................................................................................................
function  File_Dimension( size )
	if size < 1024 then
   		File_Dimension = FormatPrice(size, 2, true) & " Byte"
  	elseif size < 1048576 then
   		File_Dimension = FormatPrice((size\1024), 2, true) & " KB"
  	else
   		File_Dimension = FormatPrice((size / 1048576),2, true) & " MB"
  	end if
end function


'.................................................................................................
'..		estrae il nome del file
'..		name:		nome completo del file
'.................................................................................................
function File_Name(name)
	if instr(1, name, "/", vbTextCompare)>0 then
		File_Name = right(name, len(name) - instrrev(name, "/", vbTrue, vbTextCompare))
	else
		File_Name = name
	end if
end function


'.................................................................................................
'..		estrae l'estensione del file dal nome (estrae le prime 3 lettere dell'estensione es: HMTL = HTM)
'..		name:		nome completo del file
'.................................................................................................
function File_Extension( name )
	if instr(1, name, ".", vbTextCompare)>0 then
		File_Extension = left(right(name, len(name) - instrrev(name, ".", vbTrue, vbTextCompare)),3)
	else
		File_Extension = ""
	end if
end function


'.................................................................................................
'..		estrae l'estensione del file dal path (estrae le prime 3 lettere dell'estensione es: HMTL = HTM)
'..		name:		path con nome del file
'.................................................................................................
function FilePath_Extension( filePath )
	dim file_name
	if cString(filePath) <> "" then
		filePath = replace(filePath,"/","\")
		file_name = right(filePath, len(filePath)-inStrRev(filePath,"\"))
		FilePath_Extension = File_Extension(file_name)
	else
		FilePath_Extension = ""
	end if
end function


'.................................................................................................
'..		restituisce il nome del file icona per il tipo di file riconoscito
'..		Extension:		estensione di cui si richiede l'icona
'.................................................................................................
function File_Icon( Extension )
	if instr(1, EXTENSION_ICONS, " " & Extension & " ", vbTextCompare)>0 then
		'sceglie tra le icone
		File_Icon = "FileIcon_" & Extension & ".gif"
	else
		'tipo di file sconosciuto
		File_Icon = "FileIcon_UnknowType.gif"
	end if
end function


'.................................................................................................
'..		restituisce il nome del file per il tipo di file riconosciuto
'..		Extension:		estensione di cui si richiede l'icona
'.................................................................................................
function File_Type( Extension )
	 'sceglie tra i tipi di file
	select case Ucase(Extension)
		case "BMP"
			File_Type = "Immagine Bitmap"
		case "DOC"
			File_Type = "Documento MS WORD"
		case "DOT"
			File_Type = "Modello MS WORD"
		case "EML"
			File_Type = "Messaggio Email"
		case "MSG"
			File_Type = "Messaggio Email"
		case "GIF"
			File_Type = "Immagine GIF"
		case "HTM"
			File_Type = "Documento HTML"
		case "JPG"
			File_Type = "Immagine JPEG"
		case "JPE"
			File_Type = "Immagine JPEG"
		case "MDB"
			File_Type = "Database MS ACCESS"
		case "MDE"
			File_Type = "Applicazione MS ACCESS"
		case "PDF"
			File_Type = "Documento PDF"
		case "PPS"
			File_Type = "Presentazione MS PowerPoint"
		case "PPT"
			File_Type = "Presentazione MS PowerPoint"
		case "PUB"
			File_Type = "Documento MS Publisher"
		case "SWF"
			File_Type = "Flash Movie"
		case "FLA"
			File_Type = "Flash Sorgente"
		case "TXT"
			File_Type = "File di testo"
		case "RTF"
			File_Type = "Documento di testo formattato"
		case "XLS"
			File_Type = "Foglio MS EXCEL"
		case "XML"
			File_Type = "Documento XML"
		case "XLT"
			File_Type = "Trasformazione XLT"
		case "ZIP"
			File_Type = "Archivio ZIP"
		case "RAR"
			File_Type = "Archivio RAR"
		case "TIF"
			File_Type = "Scansione / Fax"
		case "VCF"
			File_Type = "File vCard"
		case "WAB"
			File_Type = "File della rubrica"
		case else
			'tipo di file sconosciuto
			File_Type = "File " & Ucase(Extension)
	end select
end function


'.................................................................................................
'.................................................................................................
'.................................................................................................
'FUNZIONI COMUNI
'.................................................................................................
'.................................................................................................

'.................................................................................................
'..                     Creazione drop-down da query sql                                        ..
'..                     sql		    = Query di lettura della lista                              ..
'..                     field_id    = Nome campo ID della tabella                               ..
'..                     Field_value = Nome campo che verra' visualizzato                        ..
'..                     ctrl_name   = Nome controllo                                            ..
'..                     selected    = Valore selezionato                                        ..
'..                     obbligatorio= campo obbligatorio                                        ..
'..                     style    	= style CSS			                                        ..
'..                     lingua    	= lingua attuale	                                       	..
'.................................................................................................
Sub dropDown(conn, sql, field_id, field_value, ctrl_name, selected, obbligatorio, style, lingua)
	dim rs_down
	Set rs_down = Server.CreateObject("ADODB.RecordSet")
'response.write sql
	rs_down.open sql, conn, adOpenStatic, adLockReadOnly

	CALL DropDownRecordset(rs_down, field_id, field_value, ctrl_name, selected, obbligatorio, style, lingua)
	
	rs_down.close
	set rs_down = nothing
end sub


'.................................................................................................
'..                     Creazione drop-down da query sql che estrae i valori di field_id e      ..
'..                     field_value (i primi due campi della select)                            ..
'..                     sql		    = Query di lettura della lista                              ..
'..                     ctrl_name   = Nome controllo                                            ..
'..                     selected    = Valore selezionato                                        ..
'..                     obbligatorio= campo obbligatorio                                        ..
'..                     style    	= style CSS			                                        ..
'..                     lingua    	= lingua attuale	                                       	..
'.................................................................................................
Sub dropDownSql(conn, sql, ctrl_name, selected, obbligatorio, style, lingua)
	dim rs_down
	Set rs_down = Server.CreateObject("ADODB.RecordSet")	
	rs_down.open sql, conn, adOpenStatic, adLockOptimistic

	CALL DropDownRecordset(rs_down, rs_down.Fields(0).Name, rs_down.Fields(1).Name, ctrl_name, selected, obbligatorio, style, lingua)
	
	rs_down.close
	set rs_down = nothing
end sub


'.................................................................................................
'..                     Creazione drop-down da query sql                                        ..
'..                     sql		    = Query di lettura della lista                              ..
'..                     field_id    = Nome campo ID della tabella                               ..
'..                     Field_value = Nome campo che verra' visualizzato                        ..
'..                     ctrl_name   = Nome controllo                                            ..
'..                     selected    = Valore selezionato                                        ..
'..                     obbligatorio= campo obbligatorio                                        ..
'..                     style    	= style CSS			                                        ..
'..                     BaseMessage 		= Etichetta di base "primo record - scegli" se non obbligatorio
'..                     NoRecordsMessage	= Messaggio renderizzato se non ci sono record
'................................................................................................
Sub DropDownAdvanced(conn, sql, field_id, field_value, ctrl_name, selected, obbligatorio, style, BaseMessage, NoRecordsMessage)
	dim rs_down
	Set rs_down = Server.CreateObject("ADODB.RecordSet")
	rs_down.open sql, conn, adOpenStatic, adLockReadOnly
	
	CALL DropDownRecordsetAdvanced(rs_down, field_id, field_value, ctrl_name, selected, obbligatorio, style, BaseMessage, NoRecordsMessage)
	
	rs_down.close
	set rs_down = nothing
end sub

'.............................................................................................................
'..                     Creazione drop-down da recordset                                        			..
'..                     rs_down    	= Oggetto recordset creato ed aperto contenente i dati del dropdown 	..
'..                     Field_value = Nome campo che verra' visualizzato                        			..
'..                     ctrl_name   = Nome controllo                                            			..
'..                     selected    = Valore selezionato                                        			..
'..                     obbligatorio= campo obbligatorio                                        			..
'..                     style    	= style CSS			                                        			..
'..                     lingua    	= lingua attuale	                                        			..
'.............................................................................................................
Sub DropDownRecordset(rs_down, field_id, field_value, ctrl_name, byVal selected, obbligatorio, style, lingua)
	if rs_down.absoluteposition > 1 then
		rs_down.movefirst
	end if
	CALL DropDownRecordsetAdvanced(rs_down, field_id, field_value, ctrl_name, _
								   selected, obbligatorio, style, _
								   ChooseValueByAllLanguages(lingua, "scegli...", "choose...", "w&auml;hlen...", "choisi...", "elija...", "Выбирать", "选择", "escolher"), _
								   ChooseValueByAllLanguages(lingua, "nessun elemento", "unavailable", "keine vorhanden", "aucun disponible", "ningunos disponibles", "ничего", "无", "nada"))
end Sub


'.............................................................................................................
'..                     Creazione drop-down da recordset
'..                     rs_down    			= Oggetto recordset creato ed aperto contenente i dati del dropdown
'..                     Field_value 		= Nome campo che verra' visualizzato
'..                     ctrl_name   		= Nome controllo
'..                     selected    		= Valore selezionato
'..                     obbligatorio		= campo obbligatorio
'..                     style    			= style CSS
'..                     BaseMessage 		= Etichetta di base "primo record - scegli" se non obbligatorio
'..                     NoRecordsMessage	= Messaggio renderizzato se non ci sono record
'.............................................................................................................
Sub DropDownRecordsetAdvanced(rs_down, field_id, field_value, ctrl_name, byVal selected, obbligatorio, style, BaseMessage, NoRecordsMessage)
	dim isSelected , SelectedNumeric, value

	selected = cString(selected) 
	
	SelectedNumeric = isNumeric(selected)
	if rs_down.recordcount>0 then
		rs_down.movefirst
	end if%>
	<select name="<%= ctrl_name %>" <%= style %> <%= IIF(instr(1, style, "id=""", vbTextCompare), "", " id=""" + ctrl_name + """") %>>
		<% if not obbligatorio then %>
			<option value="" <% if selected="" then %> selected <% end if %>>
				<%= BaseMessage %>
			</option>
		<% Else %>
			<% if rs_down.eof then %>
			<option value="" <% if selected="" then %> selected <% end if %>>
				<%= NoRecordsMessage %>
				</option>
			<% End If %>
		<% end if 
		while not rs_down.eof
            value = rs_down(field_id)
			if SelectedNumeric then
				isSelected = (selected = cString(value))
			elseif selected <> "" then
				isSelected = (instr(1, selected, value, vbTextCompare)>0)
			else
				isSelected = false
			end if%>
			<option value="<%= rs_down(field_id) %>" <% if isSelected then %> selected <% end if %>><%= replace(server.HTMLEncode(CString(CBLE(rs_down,field_value,Session("LINGUA")))), "  ", "&nbsp; ") %></option>
			<% rs_down.moveNext
		wend%>
	</select>
<%end Sub



'.................................................................................................
'..                     Creazione drop-down da dictionary                        				..
'..                     elements	= oggetto dictionary contenente gli elementi del select		..
'..                     ctrl_name   = Nome controllo                                            ..
'..                     selected    = Valore selezionato                                        ..
'..                     obbligatorio= campo obbligatorio                                        ..
'..                     style    	= style CSS			                                        ..
'.................................................................................................
Sub DropDownDictionary(elements, ctrl_name, selected, obbligatorio, style, lingua)
    dim isSelected , SelectedNumeric, value
	selected = cString(selected) 
	SelectedNumeric = isNumeric(selected)
    
	dim key %>
	<select name="<%= ctrl_name %>" <%= style %> <%= IIF(instr(1, style, "id=""", vbTextCompare), "", " id=""" + ctrl_name + """") %>>
		<% if not obbligatorio then %>
			<option value="" <% if selected="" then %> selected <% end if %>>
				<%= ChooseValueByAllLanguages(lingua, "scegli...", "choose...", "w&auml;hlen...", "choisi...", "elija...", "Выбирать", "选择", "escolher") %>
			</option>
		<% Else
            if elements.count=0 then %>
    			<option value="" <% if selected="" then %> selected <% end if %>>
	    			<%= ChooseValueByAllLanguages(lingua, "nessun elemento", "unavailable", "keine vorhanden", "aucun disponible", "ningunos disponibles", "ничего", "无", "nada") %>
				</option>
			<% End If
        end if 
		for each key in elements.keys
            value = key
			if SelectedNumeric then
				isSelected = (selected = cString(value))
			elseif selected <> "" then
				isSelected = (instr(1, selected, value, vbTextCompare)>0)
			else
				isSelected = false
			end if%>
			<option value="<%= Server.HTMLEncode(key) %>" title="<%= elements(key) %>" <% if isSelected then %> selected <% end if %>><%= replace(server.HTMLEncode(CString(elements(key))), "  ", "&nbsp; ") %></option>
		<%next%>
	</select>
<% end sub


'.................................................................................................
'..                     Creazione drop-down da array.
'..                     elements	= oggetto dictionary contenente gli elementi del select		..
'..                     ctrl_name   = Nome controllo                                            ..
'..                     selected    = Valore selezionato                                        ..
'..                     obbligatorio= campo obbligatorio                                        ..
'..                     style    	= style CSS			                                        ..
'.................................................................................................
Sub DropDownArray(elements, ctrl_name, selected, obbligatorio, style, lingua)
    dim isSelected , SelectedNumeric, value
	selected = cString(selected) 
	SelectedNumeric = isNumeric(selected)
    
	dim i%>
	<select name="<%= ctrl_name %>" <%= style %> <%= IIF(instr(1, style, "id=""", vbTextCompare), "", " id=""" + ctrl_name + """") %>>
		<% if not obbligatorio then %>
			<option value="" <% if selected="" then %> selected <% end if %>>
				<%= ChooseValueByAllLanguages(lingua, "scegli...", "choose...", "w&auml;hlen...", "choisi...", "elija...", "Выбирать", "选择", "escolher") %>
			</option>
		<% Else
            if ubound(elements)=0 then %>
    			<option value="" <% if selected="" then %> selected <% end if %>>
	    			<%= ChooseValueByAllLanguages(lingua, "nessun elemento", "unavailable", "keine vorhanden", "aucun disponible", "ningunos disponibles", "ничего", "无", "nada") %>
				</option>
			<% End If
        end if
		for i = lbound(elements) to ubound(elements)
            value = elements(i)
			if SelectedNumeric then
				isSelected = (selected = cString(value))
			elseif selected <> "" then
				isSelected = (instr(1, selected, value, vbTextCompare)>0)
			else
				isSelected = false
			end if%>
			<option value="<%= elements(i) %>" <% if isSelected then %> selected <% end if %>><%= server.HTMLEncode(CString(elements(i))) %></option>
		<%next%>
	</select>
<%end sub


'.................................................................................................
'..                     Creazione drop-down di interi dato l'intervallo
'..                     Inizio		= inizio dell'intervallo
'..                     Fine		= fine dell'intervallo
'..                     selected    = Valore selezionato                                        ..
'..                     obbligatorio= campo obbligatorio                                        ..
'..                     style    	= style CSS			                                        ..
'.................................................................................................
Sub DropDownInterval(Inizio, Fine, ctrl_name, selected, obbligatorio, style, lingua)
	dim i
	dim elements()
	ReDim elements(Abs(Fine - Inizio))
	if Fine < Inizio then
		for i = Inizio to Fine step - 1
			elements( Abs(inizio - i)) = i
		next
	else
		for i = Inizio to Fine
			elements(i - inizio) = i
		next
	end if
	CALL DropDownArray(elements, ctrl_name, selected, obbligatorio, style, lingua)
end sub



'.................................................................................................
'..          converte il valore n in intero, se n non e' un numero valido ritorna NULL
'..				n:		valore da convertire
'.................................................................................................
function cInteger(ByVal n)
	if instr(1, cString(n), "-", vbTextCompare) <1 then
		n = "0" & Trim(n)
	end if
	if isNumeric(n) then
		cInteger = cLng(n)
	else
		cInteger = NULL
	end if
end function


'.................................................................................................
'..          converte il valore n in intero, se n non e' un numero valido ritorna 0
'..				n:		valore da convertire
'.................................................................................................
function cIntero(ByVal n)
	if instr(1, cString(n), "-", vbTextCompare) <1 then
		n = "0" & Trim(n)
	end if
	if isNumeric(n) then
		cIntero = CLng(n)
	else
		cIntero = 0
	end if
end function


'.................................................................................................
'..          converte il valore s in una stringa: se s=NULL restituisce stringa vuota
'..				s:		valore da convertire
'.................................................................................................
function cString(ByVal s)
	cString = (s & "")
end function


'.................................................................................................
'..          converte il valore b in boolean
'..				b:			valore da convertire
'..				default:	valore di default nel caso in cui non riesco a convertire
'.................................................................................................
function CBoolean(byVal b, default)
	b = Trim(UCase(CString(b)))
	if b = "TRUE" OR b = "VERO" OR b = "1" OR b = "ON" then
		CBoolean = true
	elseif b = "FALSE" OR b = "FALSO" OR b = "0" OR b = "OFF" then
		CBoolean = false
	else
		CBoolean = default
	end if
end function


'.................................................................................................
'..         formatta la stringa aggiungendo quanti caratteri mancano al completamento
'..			str			stringa da formattare
'..			character	carattere da aggiungere per completare la lunghezza
'..			lenght		lunghezza fissa
'.................................................................................................
function FixLenght(byVal str, character, lenght)
	str = cString(str)
	if len(str)<lenght then
		FixLenght = string(lenght - len(str), character) & str
	else
		FixLenght = str
	end if
end function


'.................................................................................................
'..         converte il valore n in reale, se n non e' un numero valido ritorna NULL
'..			value:		valore da convertire
'.................................................................................................
function cReal(byVal value)
	dim str
	value = cString(value)
	if instr(1, value, "-", vbTextCompare) <1 then
		value = "0" & value
	end if

	if isItalian_SO() then
		str = replace(value,".",",")
	else
		str = replace(value,",",".")
	end if
    
	if isNumeric(str) then
		cReal = cDbl(str)
	else
		cReal = NULL
	end if
end function


'.................................................................................................
'..         converte il valore n in reale, se n non e' un numero valido o e' vuoto ritorna NULL
'..			value:		valore da convertire
'.................................................................................................
function cRealNull(byVal value)
	if CString(value) = "" then
		CRealNull = null
	else
		CRealNull = CReal(value)
	end if
end function


'.................................................................................................
'			funzione che esegue restituisce il valore ifTrue se la codizione risulta vera, altrimenti
'			restituisce il valore ifFalse
'.................................................................................................
function IIF(condition, ifTrue, ifFalse)
	
	if condition then
		IIF = ifTrue
	else
		IIF = ifFalse
	end if
    
end function


'.................................................................................................
'			funzione che verifica se il parametro "obj" &egrave; un oggetto valido e non e':
'           nothing, non e' NULL e non e' empty
'.................................................................................................
function IsObjectCreated(obj)
    if IsEmpty(obj) then
        IsObjectCreated = false
    elseif IsNull(obj) then
        IsObjectCreated = false
    elseif not IsObject(obj) then
        IsObjectCreated = false
    elseif instr(1, TypeName(obj), "nothing", vbTextCompare) then
        IsObjectCreated = false
    else
        IsObjectCreated = true
    end if
end function


'.................................................................................................
'			restituisce una stringa javascript valida che viene delimitata da delimiter
'			text:		testo da convertire
'			delimiter	delimitatore di stringa usato nello script esterno ['|"]
'.................................................................................................
function JSEncode(text, delimiter)
	dim JSText
	JSText = cString(text)
	JSText = replace(JSText, "\", "\\")
	JSText = replace(JSText, delimiter, "\" & delimiter)
	JSText = Server.HTMLEncode(JSText)
	JSText = replace(JSText, vbCrLf, "<br>")
	JSText = replace(JSText, vbLf, "<br>")
	JSText = replace(JSText, vbCr, "<br>")
	JSText = replace(JSText, "&lt;br&gt;", "<br>")
	JSEncode = JSText
end function


'.................................................................................................
'			restituisce una stringa javascript valida per la gestione del replace nei textarea di 
'           selezione multipla di elementi (vedere writecontactpicker)
'			text:		testo da convertire
'.................................................................................................
function JSReplacerEncode(text)
    JSReplacerEncode = cString(Trim(text))
    JSReplacerEncode = replace(JSReplacerEncode, "^", "")
    JSReplacerEncode = replace(JSReplacerEncode, "'", "")
    JSReplacerEncode = replace(JSReplacerEncode, """", "")
    JSReplacerEncode = replace(JSReplacerEncode, "(", " ")
    JSReplacerEncode = replace(JSReplacerEncode, ")", " ")
    JSReplacerEncode = replace(JSReplacerEncode, "\", "")
    JSReplacerEncode = replace(JSReplacerEncode, "/", "")
    JSReplacerEncode = replace(JSReplacerEncode, "{", " ")
    JSReplacerEncode = replace(JSReplacerEncode, "}", " ")
	JSReplacerEncode = replace(JSReplacerEncode, vbCrLf, "")
	JSReplacerEncode = replace(JSReplacerEncode, vbLf, "")
	JSReplacerEncode = replace(JSReplacerEncode, vbCr, "")
end function


'.................................................................................................
'			restituisce una stringa formattata per la visualizzazione di testo
'			text:		testo da convertire
'.................................................................................................
function TextEncode(byVal text)
	text = cString(text)
	text = Server.HtmlEncode(text)
	TextEncode = TextHtmlEncode(text)
end function


'.................................................................................................
'			restituisce una stringa formattata per la visualizzazione di testo
'			text:		testo da convertire
'.................................................................................................
function TextHtmlEncode(byval text)
	text = cString(text)
	text = replace(text, """", "&quot;")	'replace delle doppie virgolette
	text = replace(text, "  ", " &nbsp;")	'replace degli spazi doppi con uno spazio ed un 
											'nbsp per mantenere il risultato sul testo
	text = replace(text, vbTab, "&nbsp; &nbsp; &nbsp;")	'sostituzione dei tab con 5 spazi non eliminabili.
	text = replace(text, vbCrLf, "<br>")	'replace degli a capo con BR
    text = replace(text, vbCr, "<br>")	'replace degli a capo con BR
    text = replace(text, vbLf, "<br>")	'replace degli a capo con BR
	TextHtmlEncode = text
end function


'.................................................................................................
'..          restituisce true se il sistema operativo che si sta usando e' in italiano           ..
'.................................................................................................
function isItalian_SO()
	dim str
	str = cString(FormatNumber(4/3,2))		'numero fittizio
	isItalian_SO = instr(1,str, ",",vbTextCompare)>0
end function


'.................................................................................................
'..                     Conversione data in formato italiano                                    ..
'..                     data	    = data da convertire			                            ..
'.................................................................................................
function dateITA(byval data)
	if isDate(data) and (data&"")<>"" then
		if isItalian_SO() then
			dateITA = cDate(data)
			'tolgo l'eventuale ora
			if len(dateITA) > 10 then
				dateITA = left(dateITA, 10)
			end if
		else
			dateITA = FixLenght(Day(data), "0", 2) & "/" & FixLenght(Month(data), "0", 2) & "/" & FixLenght(Year(data), "0", 4)
		end if
	end if
end function


'.................................................................................................
'..                     conversione date in inglese                                             ..
'..                     d	    = data da convertire			                            ..
'.................................................................................................
Function dateENG(byVal d)
	if isDate(d) then
		d = cDate(d)
		if Day(d) < 13 then 
			'Corregge l'errore commesso dal S.O.se il formato e' italiano
			dateENG = FixLenght(Day(d), "0", 2) & "/" & FixLenght(Month(d), "0", 2) & "/" & FixLenght(Year(d), "0", 4)
		else  
			dateENG = d
		end if
	else
		dateENG = ""
	end if
end function


'.................................................................................................
'.. 					Giacomo 01/03/2011
'..                     conversione date in inglese                                             ..
'..                     data    = data da convertire			                            ..
'.................................................................................................
Function dateEN(byVal data)
	if isDate(data) then
		dateEN = FixLenght(Month(data), "0", 2) & "/" & FixLenght(Day(data), "0", 2) & "/" & FixLenght(Year(data), "0", 4)
	else
		dateEN = ""
	end if
end function


'.................................................................................................
'..                     conversione date in base alla lingua
'..                     d	    = data da convertire			                            ..
'.................................................................................................
Function FormatDate(d, lingua)
	SELECT CASE lingua
		CASE LINGUA_INGLESE
			FormatDate = DateENG(d)
		CASE ELSE
			FormatDate = DateITA(d)
	END SELECT
end function


'.................................................................................................
'..                     conversione time in italiano                                            ..
'..                     dt	    = data con ore da convertire			                        ..
'.................................................................................................
function TimeIta(dt)
	if isDate(dt) then
		TimeIta = FixLenght(Hour(dt), "0", 2) & ":" & FixLenght(minute(dt), "0", 2)
	else
		TimeIta = ""
	end if
end function


'.................................................................................................
'..                     conversione date-time in italiano                                            ..
'..                     d	    = data da convertire			                            ..
'.................................................................................................
function DateTimeIta(d)
	if isDate(d) then
		DateTimeIta = FixLenght(Day(d), "0", 2) & "/" & FixLenght(Month(d), "0", 2) & "/" & FixLenght(Year(d), "0", 4) & " " & TimeIta(d)
	else
		DateTimeIta = ""
	end if
end function



'.................................................................................................
'..                     conversione date-time in inglese                                            ..
'..                     d	    = data da convertire			                            ..
'.................................................................................................
function DateTimeIng(d)
	if isDate(d) then
		DateTimeIng = FixLenght(Month(d), "0", 2) & "/" & FixLenght(Day(d), "0", 2) & "/" & FixLenght(Year(d), "0", 4) & " " & TimeIta(d)
	else
		DateTimeIng = ""
	end if
end function



'.......................................................................................................
'..                     conversione date-time in italiano o inglese a seconda del parametro lingua    ..
'..                     d	    = data da convertire			      	                              ..
'.......................................................................................................
function DateTimeLingua(d, lingua)
	if isDate(d) then
		if lingua <> "" then
			select case lingua
				case "it"
					DateTimeLingua = DateTimeIta(d)
				case "en"
					DateTimeLingua = DateTimeIng(d)
				case else
					DateTimeLingua = DateTimeIta(d)
			end select
		else
			DateTimeLingua = DateTimeIta(d)
		end if
	else
		DateTimeLingua = ""
	end if
end function



'.................................................................................................
'..                     Conversione data in formato iso                                         ..
'..                     data	    = data da convertire			                            ..
'.................................................................................................
function DateISO(data)
	if isDate(data) and (data&"")<>"" then
		DateISO = FixLenght(Year(data), "0", 4) & "-" & FixLenght(Month(data), "0", 2) & "-" & FixLenght(Day(data), "0", 2)
	else
		DateISO = ""
	end if
end function


'.................................................................................................
'..                     Conversione data in formato iso                                         ..
'..                     data	    = data da convertire			                            ..
'.................................................................................................
function DateTimeISO(data)
	DateTimeISO = DateISO(data) &" "& FixLenght(Hour(data), "0", 2) &"."& FixLenght(Minute(data), "0", 2) &"."& FixLenght(Second(data), "0", 2)
end function


'.................................................................................................
'..                     Conversione data in formato iso per nome di file                                        ..
'..                     data	    = data da convertire			                            ..
'.................................................................................................
function DateTimeISOFile(data)
	DateTimeISOFile = replace(DateISO(data) & "--" & FixLenght(Hour(data), "0", 2) & "-" & FixLenght(Minute(data), "0", 2) & "-" & FixLenght(Second(data), "0", 2),"-", "")
end function


'.................................................................................................
'..                     Conversione tempo in formato iso                                         ..
'..                     data	    = data della quale convertire il tempo			                            ..
'.................................................................................................
function TimeISO(data)
	TimeISO = Hour(data) &":"& Minute(data) &":"& Second(data)
end function


'.................................................................................................
'..                     Conversione data in formato esteso                                      ..
'..                     data	    = data da convertire			                            ..
'..                     lingua	    = lingua in cui devono essere restituiti il mese ed il giorno
'.................................................................................................
function DataEstesa(data, lingua)
	DataEstesa = NomeGiorno(data, lingua) & " " & day(data) & " " & Nomemese(Month(data), lingua) & " " & Year(data)
end function


'.................................................................................................
'..                     Restituisce una data dall'input in UTC, ignora lo scostamento da GW     ..
'..                     data	    = data da convertire			                            ..
'.................................................................................................
function DataFromUTC(byVal data)
	if CString(data) = "" then
		DataFromUTC = ""
	else
		DataFromUTC = Replace(data, "T", " ")
		DataFromUTC = Left(DataFromUTC, 19)
	end if
end function



'.................................................................................................
'..                     Restituisce il nome del giorno                                          ..
'..                     input	    = data di cui prendere il giorno  o indice visual basic del giorno                          ..
'..                     lingua	    = lingua in cui deve essere restituito il nome del giorno
'.................................................................................................
function NomeGiorno(input, lingua)
	dim vbDay
	if isDate(input) then
		'se viene passata una data calcola il giorno della settimana di quella data
		vbDay = WeekDay(input, VbSunday)
	elseif isNumeric(input) then
		'se viene passato un nunmero lo converte e lo usa come indice vb del giorno
		vbDay = cInteger(input)
	end if
	Select case lingua
		case LINGUA_INGLESE
			Select case vbDay
				case VbSunday		NomeGiorno = "sunday"
				case VbMonday		NomeGiorno = "monday"
				case vbTuesday		NomeGiorno = "tuesday"
				case vbWednesday	NomeGiorno = "wednesday"
				case vbThursday		NomeGiorno = "thursday"
				case vbFriday		NomeGiorno = "friday"
				case vbSaturday		NomeGiorno = "saturday"
			end select
		case LINGUA_FRANCESE
			Select case vbDay
				case VbSunday		NomeGiorno = "dimanche"
				case VbMonday		NomeGiorno = "lundi"
				case vbTuesday		NomeGiorno = "mardi"
				case vbWednesday	NomeGiorno = "mercredi"
				case vbThursday		NomeGiorno = "jeudi"
				case vbFriday		NomeGiorno = "vendredi"
				case vbSaturday		NomeGiorno = "samedi"
			end select 
		case LINGUA_TEDESCO
			Select case vbDay
				case VbSunday		NomeGiorno = "sonntag"
				case VbMonday		NomeGiorno = "montag"
				case vbTuesday		NomeGiorno = "dienstag"
				case vbWednesday	NomeGiorno = "mittwoch"
				case vbThursday		NomeGiorno = "donnerstag"
				case vbFriday		NomeGiorno = "freitag"
				case vbSaturday		NomeGiorno = "sonnabend"
			end select 
		case LINGUA_SPAGNOLO
			Select case vbDay
				case VbSunday		NomeGiorno = "domingo"
				case VbMonday		NomeGiorno = "lunes"
				case vbTuesday		NomeGiorno = "martes"
				case vbWednesday	NomeGiorno = "mi&eacute;rcoles"
				case vbThursday		NomeGiorno = "jueves"
				case vbFriday		NomeGiorno = "viernes"
				case vbSaturday		NomeGiorno = "s&aacute;bado"
			end select
		case LINGUA_RUSSO
			Select case vbDay
				case VbSunday		NomeGiorno = "воскресенье"
				case VbMonday		NomeGiorno = "понедельник"
				case vbTuesday		NomeGiorno = "вторник"
				case vbWednesday	NomeGiorno = "cреда"
				case vbThursday		NomeGiorno = "четверг"
				case vbFriday		NomeGiorno = "пятница"
				case vbSaturday		NomeGiorno = "Суббота"
			end select
		case LINGUA_CINESE
			Select case vbDay
				case VbSunday		NomeGiorno = "星期日"
				case VbMonday		NomeGiorno = "星期一"
				case vbTuesday		NomeGiorno = "星期二"
				case vbWednesday	NomeGiorno = "星期三"
				case vbThursday		NomeGiorno = "星期四"
				case vbFriday		NomeGiorno = "星期五"
				case vbSaturday		NomeGiorno = "星期六"
			end select
		case LINGUA_PORTOGHESE
			Select case vbDay
				case VbSunday		NomeGiorno = "domingo"
				case VbMonday		NomeGiorno = "segunda-feira"
				case vbTuesday		NomeGiorno = "terça-feira"
				case vbWednesday	NomeGiorno = "Quarta-feira"
				case vbThursday		NomeGiorno = "Quinta-feira"
				case vbFriday		NomeGiorno = "Sexta-feira"
				case vbSaturday		NomeGiorno = "Sábado"
			end select
		case else
			'LINGUA_ITALIANO
			Select case vbDay
				case VbSunday 		NomeGiorno = "domenica"
				case VbMonday 		NomeGiorno = "lunedi"
				case vbTuesday		NomeGiorno = "martedi"
				case vbWednesday	NomeGiorno = "mercoledi"
				case vbThursday		NomeGiorno = "giovedi"
				case vbFriday		NomeGiorno = "venerdi"
				case vbSaturday		NomeGiorno = "sabato"
			end select 
		
	end select
end function


'.................................................................................................
'..                     Restituisce il nome del mese                                          ..
'..                     month	    = indice del mese
'..                     lingua	    = lingua in cui deve essere restituito il nome del mese
'.................................................................................................
function NomeMese(mese, lingua)
	Select case lingua
		case LINGUA_INGLESE
			Select Case mese
				Case 1		NomeMese = "january"
				Case 2		NomeMese = "february"
				Case 3		NomeMese = "march"
				Case 4		NomeMese = "april"
				Case 5		NomeMese = "may"
				Case 6		NomeMese = "june"
				Case 7		NomeMese = "july"
				Case 8		NomeMese = "august"
				Case 9		NomeMese = "september"
				Case 10		NomeMese = "october"
				Case 11		NomeMese = "november"
				Case 12		NomeMese = "december"
			end select
		case LINGUA_FRANCESE
			Select Case mese
				Case 1		NomeMese = "janvier"
				Case 2		NomeMese = "f&eacute;vrier"
				Case 3		NomeMese = "mars"
				Case 4		NomeMese = "avril"
				Case 5		NomeMese = "mai"
				Case 6		NomeMese = "juin"
				Case 7		NomeMese = "juillet"
				Case 8		NomeMese = "ao&ucirc;t"
				Case 9		NomeMese = "septembre"
				Case 10		NomeMese = "octobre"
				Case 11		NomeMese = "novembre"
				Case 12		NomeMese = "d&eacute;cembre"
			end select
		case LINGUA_TEDESCO
			Select Case mese
				Case 1		NomeMese = "januar"
				Case 2		NomeMese = "februar"
				Case 3		NomeMese = "m&auml;rz"
				Case 4		NomeMese = "april"
				Case 5		NomeMese = "mai"
				Case 6		NomeMese = "juni"
				Case 7		NomeMese = "juli"
				Case 8		NomeMese = "august"
				Case 9		NomeMese = "september"
				Case 10		NomeMese = "oktober"
				Case 11		NomeMese = "november"
				Case 12		NomeMese = "dezember"
			end select
		case LINGUA_SPAGNOLO
			Select Case mese
				Case 1		NomeMese = "enero"
				Case 2		NomeMese = "febrero"
				Case 3		NomeMese = "marzo"
				Case 4		NomeMese = "abril"
				Case 5		NomeMese = "mayo"
				Case 6		NomeMese = "junio"
				Case 7		NomeMese = "julio"
				Case 8		NomeMese = "agosto"
				Case 9		NomeMese = "septiembre"
				Case 10		NomeMese = "octubre"
				Case 11		NomeMese = "noviembre"
				Case 12		NomeMese = "diciembre"
			end select
		case LINGUA_RUSSO
			Select Case mese
				Case 1		NomeMese = "Январь"
				Case 2		NomeMese = "Февраль"
				Case 3		NomeMese = "Март"
				Case 4		NomeMese = "Апрель"
				Case 5		NomeMese = "май"
				Case 6		NomeMese = "июнь"
				Case 7		NomeMese = "Июль"
				Case 8		NomeMese = "август"
				Case 9		NomeMese = "сентябрь"
				Case 10		NomeMese = "октябрь"
				Case 11		NomeMese = "Ноябрь"
				Case 12		NomeMese = "декабрь"
			end select
		case LINGUA_CINESE
			Select Case mese
				Case 1		NomeMese = "1月"
				Case 2		NomeMese = "2月"
				Case 3		NomeMese = "3月"
				Case 4		NomeMese = "4月"
				Case 5		NomeMese = "5月"
				Case 6		NomeMese = "6月"
				Case 7		NomeMese = "7月"
				Case 8		NomeMese = "8月"
				Case 9		NomeMese = "9月"
				Case 10		NomeMese = "10月"
				Case 11		NomeMese = "11月"
				Case 12		NomeMese = "12月"
			end select
		case LINGUA_PORTOGHESE
			Select Case mese
				Case 1		NomeMese = "janeiro"
				Case 2		NomeMese = "fevereiro"
				Case 3		NomeMese = "março"
				Case 4		NomeMese = "abril"
				Case 5		NomeMese = "maio"
				Case 6		NomeMese = "junho"
				Case 7		NomeMese = "julho"
				Case 8		NomeMese = "agosto"
				Case 9		NomeMese = "setembro"
				Case 10		NomeMese = "outubro"
				Case 11		NomeMese = "novembro"
				Case 12		NomeMese = "dezembro"
			end select
		case else
			'LINGUA_ITALIANO
			Select Case mese
				case 1		NomeMese = "gennaio"
				case 2		NomeMese = "febbraio"
				case 3		NomeMese = "marzo"
				case 4		NomeMese = "aprile"
				case 5		NomeMese = "maggio"
				case 6		NomeMese = "giugno"
				case 7		NomeMese = "luglio"
				case 8		NomeMese = "agosto"
				case 9		NomeMese = "settembre"
				case 10		NomeMese = "ottobre"
				case 11		NomeMese = "novembre"
				case 12		NomeMese = "dicembre"
			end select
	end select
end function



'.................................................................................................
'restituisce l'ultimo giorno del mese indicato, nell'anno indicato
'.................................................................................................
Function Date_LastDay(intMonth,intYear)

	Dim intDay

	Select Case intMonth
		Case 1, 3, 5, 7, 8, 10, 12
			intDay = 31
		Case 4, 6, 9, 11
			intDay = 30
		Case 2
			If intYear mod 4 = 0 Then
				If intYear mod 100 = 0 AND intYear mod 400 <> 0 Then
					intDay = 28
				Else
					intDay = 29
				End If
			Else
				intDay = 28
			End If
	End Select

	Date_LastDay = intDay

End Function


'.................................................................................................
'..	restituisce la differenza percentuale tra due valori
'calcola la percentuale dal valore 1 al valore 2
'...............................................................................................
function GetPercentualeVariazione(value1, value2)
	value1 = cReal(value1)
	value2 = cReal(value2)
	if value1 = value2 then
		GetPercentualeVariazione = 0
	elseif value1 = 0 then
		GetPercentualeVariazione = 100
	elseif value2 = 0 then
		GetPercentualeVariazione = 0-100
	else
		GetPercentualeVariazione = (value1/value2)*100
	end if
end function


'.................................................................................................
'..                     Restituisce un elenco con i valori della query                          ..
'..                     conn:		connessione al database aperta
'..						rs:			oggetto recordset chiuso (se NULL lo crea internamente)
'..						sql:		query per prelevare i valori
'.................................................................................................
Function GetValueList(conn, rs, sql)
	dim rsCreated, connCreated
	if not IsObjectCreated(conn) then
		connCreated = true
		set conn = server.createobject("adodb.connection")
		conn.open Application("DATA_ConnectionString"),"",""
	else
		connCreated = false
	end if
    if not IsObjectCreated(rs) then
		rsCreated = true
		set rs = server.createobject("adodb.recordset")
	else
		rsCreated = false
	end if
	'response.write "<br><br>" &  sql & "<br><br>"

 	rs.open sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	GetValueList = ValueList(rs, 0)
	
	rs.close
	if rsCreated then
		set rs = nothing
	end if
	if connCreated then
		conn.close
		set conn = nothing
	end if
end function


Function ValueList(rs, field)
	dim list
	while not rs.eof
		list = list & rs(field)
		rs.movenext
		if not rs.eof then
			list = list & ", "
		end if
	wend
	ValueList = list
end function


'.................................................................................................
'funzione che controlla se l'id e' presente nella lista
'   list        lista di id nella quale cercare
'   id          id da ricercare nella lista
'.................................................................................................
function InIdList(list, id)
    if instr(1, "," & replace(list, " ", "") & ",", "," & id & ",", vbTextCompare)>0 then
        InIdList = true
    else
        InIdList = false
    end if
end function


'.................................................................................................
'..                     Controlla la correttezza dell'email
'..                     Email:		email da controllare
'.................................................................................................
FUNCTION IsEmail( byVal Email )
	dim at_position, dot_position
	Email = cString(Email)
	'Controlla se email troppo corta
	IsEmail = Len(Email)>7
		if not IsEmail then Exit function
	
	at_position = instr(1,Email,"@",vbTextCompare)
	dot_position = InStrRev(Email, ".", -1,vbTextCompare)
	
	'controlla se e' presente @
	IsEmail = at_position > 0
		if not IsEmail then Exit function
	
	'controlla se e' presente una sola @
	IsEmail = instr(at_position+1, Email,"@" ,vbTextCompare)=0
		if not IsEmail then Exit function
	
	'controlla se e' presente almeno un . dopo @
	IsEmail = instr(at_position, Email,"." ,vbTextCompare)>0
		if not IsEmail then Exit function
		
	'controlla se i caratteri sono corretti
	IsEmail = CheckChar(Email, EMAIL_VALID_CHARSET)
		if not IsEmail then Exit function
		
	'controlla se i caratteri dopo l'ultimo punto sono corretti: solo lettere
	IsEmail = CheckChar(right(email, len(email)-dot_position), DOMAIN_VALID_CHARSET)
		if not IsEmail then Exit function
	
END FUNCTION


'.................................................................................................
'..                     Controlla la correttezza del numero di telefono
'..                     numero:		numero di tele
'.................................................................................................
function IsPhoneNumber(byVal numero )
	numero = cString(numero)
	'ripulisce da caratteri spuri.
	numero = RemoveInvalidChar(numero, PHONE_VALID_CHARSET)
	
	IsPhoneNumber = len(numero) > 2
		if not IsPhoneNumber then Exit function
	
end function


'.................................................................................................
'..                     formatta il numero di cellulare (ITALIANO)
'..                     numero:		numero di cellulare
'.................................................................................................
function FormatMobilePhone(byVal numero )
	numero = RemoveInvalidChar(numero, PHONE_VALID_CHARSET)
	if len(numero) < 11 then
		numero = "+39" + numero
	elseif left(4, numero) = "0039" then
		numero = "+" + right(numero, len(numero)-2)
	elseif instr(1, numero, "+", vbTextCompare)<1 then
		numero = "+" + numero
	end if
	
	FormatMobilePhone = numero
end function


'.................................................................................................
'..          converte il valore c in un carattere base asci es: converte &egrave; con e
'..			 rimuove eventuali caratteri aggu
'..				c:		valore da convertire
'.................................................................................................
function CharToAsii(c, DefaultChar)
	
	Select case Server.HtmlEncode(c)
		case "&#227;", "&#229;", "&#230;", "&#228;", "&#226;", "&#224;", "&atilde;", "&aring;", "&aelig;", "&auml;", "&acirc;", "&agrave;"
			CharToAsii = "a"
		case "&#193;", "&#195;", "&#197;", "&#198;", "&#196;", "&#194;", "&#192;", "&Aacute;", "&Atilde;", "&Aring;", "&AElig;", "&Auml;", "&Acirc;", "&Agrave;"
			CharToAsii = "A"
		case "&#233;", "&#235;", "&#234;", "&#232;", "&eacute;", "&euml;", "&eth;", "&ecirc;", "&egrave;"
			CharToAsii = "e"
		case "&#201;", "&#203;", "&#202;", "&#200;", "&#8364;", "&Eacute;", "&Euml;", "&Ecirc;", "&Egrave;", "&euro;"
			CharToAsii = "E"
		case "&#237;", "&#239;", "&#238;", "&#236;", "&iacute;", "&iuml;", "&icirc;", "&igrave;"
			CharToAsii = "i"
		case "&#205;", "&#207;", "&#206;", "&#204;", "&Iacute;", "&Iuml;", "&Icirc;", "&Igrave;"
			CharToAsii = "I"
		case "&#243;", "&#245;", "&#248;", "&#246;", "&#244;", "&#242;", "&#240;", "&oacute;", "&otilde;", "&oslash;", "&ouml;", "&ocirc;", "&ograve;"
			CharToAsii = "o"
		case "&#211;", "&#213;", "&#216;", "&#214;", "&#212;", "&#210;", "&Oacute;", "&Otilde;", "&Oslash;", "&Ouml;", "&Ocirc;", "&Ograve;"
			CharToAsii = "O"
		case "&#249;", "&#251;", "&#252;", "&#250;", "&ugrave;", "&ucirc;", "&uuml;", "&uacute;"
			CharToAsii = "u"
		case "&#217;", "&#219;", "&#220;", "&#218;", "&Ugrave;", "&Ucirc;", "&Uuml;", "&Uacute;"
			CharToAsii = "U"
		case "&#253;", "&#255;", "&yacute;", "&yuml;"
			CharToAsii = "y"
		case "&#221;", "&Yacute;"
			CharToAsii = "Y"
		case "&#231;", "&#162;", "&ccedil;", "&cent;"
			CharToAsii = "c"
		case "&#199;", "&Ccedil;"
			CharToAsii = "C"
		case "&#241;", "&ntilde;"
			CharToAsii = "n"
		case "&#209;", "&Ntilde;"
			CharToAsii = "N"
		case else
			c = DefaultChar
	end select
	
end function


'.................................................................................................
'		funzione che converte la stringa in ingresso da alfatebto cirrilico e caratteri utf8 estesi ad alfabeto latino base
'.................................................................................................
function ExtendedUTFToBaseLatin(byVal str)
	dim i, latin , extended
	'dichiara array per sostituzione caratteri cirillici
	latin =    array("ae", "ss", "oe", "ue", "c", "a", "a", "b", "b", "v", "v", "g", "g", "d", "d", "je", "je", "jo", "jo", "zh", "zh", "z", "z", "i", "i", "j", "j", "k", "k", "l", "l", "m", "m", "n", "n", "o", "o", "p", "p", "r", "r", "s", "s", "t", "t", "u", "u", "f", "f", "h", "h", "ts", "ts", "ch", "ch", "sh", "sh", "shch", "shch", "",  "",  "y", "y", "",  "",  "e", "e", "ju", "ju", "ja", "ja")
	extended = array("ä",  "ß",  "ö",  "ü",  "ç", "А", "а", "Б", "б", "В", "в", "Г", "г", "Д", "д", "Е",  "е",  "Ё",  "ё",  "Ж",  "ж",  "З", "з", "И", "и", "Й", "й", "К", "к", "Л", "л", "М", "м", "Н", "н", "О", "о", "П", "п", "Р", "р", "С", "с", "Т", "т", "У", "у", "Ф", "ф", "Х", "х", "Ц",  "ц",  "Ч",  "ч",  "Ш",  "ш",  "Щ",    "щ",    "Ъ", "ъ", "Ы", "ы", "Ь", "ь", "Э", "э", "Ю",  "ю",  "Я",  "я")
	
	'esegue sostituzione
	for i = lbound(latin) to ubound(latin)	
		str = replace(str, extended(i), latin(i))
	next
	
	ExtendedUTFToBaseLatin = str

end function


'.................................................................................................
'		funzione che rimuove il carattere indicato all'inizio della stringa
'		strToTrim:			stringa da cui rimuovere il carattere
'		charToRemove:		carattere da rimuovere
'.................................................................................................
function TrimStart(byVal strToTrim, strToRemove)
	strToTrim = cString(strToTrim)
	strToRemove = cString(strToRemove)
	if strToTrim <> "" then
		while left(strToTrim, len(strToRemove)) = strToRemove
			strToTrim = right(strToTrim, len(strToTrim) - len(strToRemove))
		wend
	end if
	TrimStart = strToTrim
end function


'.................................................................................................
'		funzione che rimuove il carattere indicato alla fine della stringa
'		strToTrim:			stringa da cui rimuovere il carattere
'		charToRemove:		carattere da rimuovere
'.................................................................................................
function TrimEnd(byVal strToTrim, strToRemove)
	strToTrim = cString(strToTrim)
	strToRemove = cString(strToRemove)
	if strToTrim <> "" then
		while right(strToTrim, len(strToRemove)) = strToRemove
			strToTrim = left(strToTrim, len(strToTrim) - len(strToRemove))
		wend
	end if
	TrimEnd = strToTrim
end function


'.................................................................................................
'..                     Controlla che i caratteri della stringa str siano tutti contenuti in valid_charset
'..                     str:			stringa di caratteri da controllare
'..						valid_charset:	insieme di caratteri
'.................................................................................................
function CheckChar(byVal str, valid_charset)
	CheckChar = (str = RemoveInvalidChar(str, valid_charset))
end function


'....................................................................................................
'	funzione che ripulisce la stringa dai caratteri non compresi nel charset indicato
'	ritorna la stringa ripulita.
'....................................................................................................
function RemoveInvalidChar(byVal str, valid_charset)
	RemoveInvalidChar = RemoveByCharset(str, valid_charset, false)
end function
'....................................................................................................
'	
'....................................................................................................
function RemoveByCharset(str, charset, remove_charset)
	dim c, i, CheckedStr
	CheckedStr = ""	
	if cString(str)<>"" then
		for i=1 to len(str)
			c = Mid(str, i, 1)
			if remove_charset then
				'rimuove caratteri presenti nel charset
				if not instr(1, charset, c, vbTextCompare)>0 then
					CheckedStr = CheckedStr & c
				end if
			else
				'rimuove caratteri NON presenti nel charset
				if instr(1, charset, c, vbTextCompare)>0 then
					CheckedStr = CheckedStr & c
				end if
			end if
		next
	end if
	RemoveByCharset = CheckedStr
end function


'.................................................................................................
'..				Funzione che conta le occorrenze della stringa "splitter" nella stringa "str"
'..				str			stringa in cui cercare il carattere
'..				splitter	stringa da cercare
'.................................................................................................
function Count(byVal str, splitter)
	if str <> "" AND splitter<>"" then
        if instr(1, str, splitter, vbTextCompare)>0 then
    		dim a
	    	a = Split(str, splitter)
		    Count = uBound(a)
        else
            Count = 0
        end if
	else
		Count = 0
	end if
end Function


'.................................................................................................
'..				Compone input per immissione date con finestra di selezione
'..				form: 				nome del form contenente l'input
'..				Input:				nome dell'input
'..				value:				valore associato
'..				PathStili:			percorso relativo del file di stili: specificare solo se il file non e' in library
'..				PathCalendarOffset : displacement path per directory library
'..				ShowReset:			indica se mostrare il pulsante di reset o no
'..				AllowPast:			indica se la finestra permette di scegliere date precedenti ad oggi
'..				lingua:				lingua in cui deve essere generato il calendario
'.................................................................................................
sub WriteDataPicker_Input(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua)
	CALL WriteDataPicker_Input_Ex(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua,"")
end sub



'.................................................................................................
'..				Compone input per immissione date con finestra di selezione e coordinamento
'..				valore input2 collegato (dal - al)
'..				form: 				nome del form contenente l'input
'..				Input:				nome dell'input
'..				value:				valore associato
'..				PathStili:			percorso relativo del file di stili: specificare solo se il file non e' in library
'..				PathCalendarOffset : displacement path per directory library
'..				ShowReset:			indica se mostrare il pulsante di reset o no
'..				AllowPast:			indica se la finestra permette di scegliere date precedenti ad oggi
'..				lingua:				lingua in cui deve essere generato il calendario
'..				Input2:				nome dell'imput da mantenere collegato.
'.................................................................................................
sub WriteDataPicker_Input_Ex(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua, Input2)
	CALL WriteDataPicker_Input_Manuale(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua, Input2, false, "")
end sub


'.................................................................................................
'..				Compone input per immissione date con finestra di selezione e coordinamento
'..				valore input2 collegato (dal - al)
'..				form: 				nome del form contenente l'input
'..				Input:				nome dell'input
'..				value:				valore associato
'..				PathStili:			percorso relativo del file di stili: specificare solo se il file non e' in library
'..				PathCalendarOffset : displacement path per directory library
'..				ShowReset:			indica se mostrare il pulsante di reset o no
'..				AllowPast:			indica se la finestra permette di scegliere date precedenti ad oggi
'..				lingua:				lingua in cui deve essere generato il calendario
'..				Input2:				nome dell'imput da mantenere collegato
'..				AllowInsert:		permette anche l'inserimento manuale della data
'..				InputStyle:			stili aggiuntivi per l'input che mostra la data
'.................................................................................................
sub WriteDataPicker_Input_Manuale(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua, Input2, AllowInsert, InputStyle)
	CALL WriteDataPicker_Input_Manuale2(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua, Input2, AllowInsert, InputStyle, "")
end sub

'.................................................................................................
'..				Compone input per immissione date con finestra di selezione e coordinamento
'..				valore input2 collegato (dal - al)
'..				form: 				nome del form contenente l'input
'..				Input:				nome dell'input
'..				value:				valore associato
'..				PathStili:			percorso relativo del file di stili: specificare solo se il file non e' in library
'..				PathCalendarOffset : displacement path per directory library
'..				ShowReset:			indica se mostrare il pulsante di reset o no
'..				AllowPast:			indica se la finestra permette di scegliere date precedenti ad oggi
'..				lingua:				lingua in cui deve essere generato il calendario
'..				Input2:				nome dell'imput da mantenere collegato
'..				AllowInsert:		permette anche l'inserimento manuale della data
'..				InputStyle:			stili aggiuntivi per l'input che mostra la data
'..				NameFunctionAfterClick:	nome della funzione richiamata dopo aver scelto il giorno nel calendario
'.................................................................................................
sub WriteDataPicker_Input_Manuale2(Form, Input, Value, PathStili, PathCalendarOffset, ShowReset, AllowPast, Lingua, Input2, AllowInsert, InputStyle, NameFunctionAfterClick)
	dim label%>
	<script language="JavaScript" type="text/javascript">
		function <%= Form %>_<%= Input %>_onClick(){
			var top, left, width, height, position_properties, except;
			<% if AllowPast then %>
				var past = "&AllowPast=1"
			<% else %>
				var past = ""
			<% end if %>
			width=240;
			height=155;
			
			try {
				//calcola coordinata Y
				top = (event.screenY - event.offsetY) + 20;
				if ((top + height)>(screen.height-90))
					top = (screen.height - height - 90);
	
				//calcola coordinata X
				left = (event.screenX - (width/2));
				if (left <20)
					left = 20;
				else if ((left + width)>(screen.width-20))
					left = (screen.width - width - 20);
					
				position_properties = "left=" + left + ",top=" + top + ",";
			}
			catch(except){
				position_properties = "";
			}
			
			<%if PathStili = "" then
				PathStili = "stili.css"
			end if%>
			
			if (!document.<%= Form %>.<%= input %>.disabled){
				window.open("<%= PathCalendarOffset %>amministrazione/library/PickerDate.asp?lingua=<%= Lingua %>&inputvalue=" + document.<%= Form %>.<%= input %>.value + "&input=<%= Input %>&form=<%= Form %>&stili=<%= Server.UrlEncode(PathStili) %>&campoAggiorna=<%= Input2 %>&nameFunctionAfterClick=<%= NameFunctionAfterClick %>" + past, "<%= Form & "_" & Input %>" ,
							position_properties + "width=" + width + ",height=" + height + ",scrollbars=no,statusbar=no,menubars=no,resizable=yes");
			}			
			
		}
		
		function <%= Form %>_<%= Input %>_onReset(){
			if (!document.<%= Form %>.<%= input %>.disabled){
				document.<%= Form %>.<%= input %>.value = ""
			}
		}
		
	</script>
		<% Select case lingua
			case LINGUA_INGLESE
				label = "CHOOSE"
			case LINGUA_FRANCESE
				label = "CHOISIR"
			case LINGUA_TEDESCO
				label = "W&Auml;HLEN"
			case LINGUA_SPAGNOLO
				label = "ELIJA"
			case else
				label = "SCEGLI"
		end select %>
	<table border="0" cellspacing="0" cellpadding="0" class="PickerComponent">
		<tr>
			<td><input type="Text" <%if not cBoolean(AllowInsert,false) then%>onclick="<%= Form %>_<%= Input %>_onClick()" READONLY <%end if%> name="<%= Input %>" id="<%= Input %>" value="<%=Trim(cString(Value))%>" size="12" maxlength="10" style="text-align:center; letter-spacing:1px; <%=InputStyle%>" class="PickerDateInput"></td>
			<td>
				<a href="javascript:void(0);" class="button_input" id="<%= Form %>_link_scegli_<%= Input %>" name="<%= Form %>_link_scegli_<%= Input %>" onclick="<%= Form %>_<%= Input %>_onClick()" title="<%= label %>" <%= ACTIVE_STATUS %>>
					<%= label %>
				</a>
			</td>
			<% if ShowReset then %>
			<td>	
				<a href="javascript:void(0);" class="button_input" id="<%= Form %>_link_reset_<%= Input %>" name="<%= Form %>_link_reset_<%= Input %>" onclick="<%= Form %>_<%= Input %>_onReset()" title="RESET" <%= ACTIVE_STATUS %>>
					RESET
				</a>
			</td>
			<% end if %>
		</tr>
	</table>
<%end sub


'.................................................................................................
'..		Restituisce il drop down con la scelta di un orario (formato hh:mm)
'..		inputName			nome del drop down
'..		intervallo			intervallo, in minuti, tra due orari selezionabili
'..		selectedValue		valore selezionato, con formato hh:mm (se null allora sarà il limiteInferiore)
'..		limiteInferiore
'..		limiteSuperiore
'.................................................................................................
sub WriteDropDownOrario(inputName,intervallo,selectedValue,limiteInferiore,limiteSuperiore)
	CALL WriteDropDownOrarioCompleto(inputName,intervallo,selectedValue,limiteInferiore,limiteSuperiore, true, false)
end sub


'.................................................................................................
'..		Restituisce il drop down con la scelta di un orario (formato hh:mm)
'..		inputName			nome del drop down
'..		intervallo			intervallo, in minuti, tra due orari selezionabili
'..		selectedValue		valore selezionato, con formato hh:mm (se null allora sarà il limiteInferiore)
'..		limiteInferiore
'..		limiteSuperiore
'..		obbligatorio		se true deve per forza essere selezionato un valore
'..		valueInMinuti		se true il valore del campo è espresso in minuti
'.................................................................................................
sub WriteDropDownOrarioCompleto(inputName, intervallo, selectedValue, limiteInferiore, limiteSuperiore, obbligatorio, valueInMinuti)
	dim ora, minuti, controllo, orario, ora_da, ora_a, min_da, min_a
	if cString(limiteInferiore) = "" then limiteInferiore = "00:00"
	if cString(limiteSuperiore) = "" then limiteSuperiore = "23:59"
	ora_da = Left(limiteInferiore,Instr(limiteInferiore,":")-1)
	min_da = Right(limiteInferiore,2)
	ora_a = cIntero(Left(limiteSuperiore,Instr(limiteSuperiore,":")-1))
	min_a = cIntero(Right(limiteSuperiore,2))
	
	'if cString(selectedValue) = "" then selectedValue = "00:00"
	minuti = -intervallo
	
	%>
	<select name="<%=inputName%>">
		<% if not obbligatorio then %>
			<option value="">Scegli...</option>
		<% end if %>
		<%
		ora = DATE() & " "&ora_da&":"&min_da&":00"
		do while true
			minuti = minuti + intervallo
			orario = TimeIta(cString(Hour(ora))&":"&cString(Minute(ora)))
			%>
			<option value="<%=IIF(valueInMinuti, minuti, orario)%>" <%= IIF(orario = selectedValue OR minuti = selectedValue, "selected", "")%>><%=orario%></option>
			<% ora = DateAdd("n", intervallo, ora) 
			if Hour(ora) = ora_a then
				controllo = Minute(ora)
				while controllo < min_a
					orario = TimeIta(cString(Hour(ora))&"."&cString(Minute(ora)))
					%>
					<option value="<%=IIF(valueInMinuti, minuti, orario)%>" <%= IIF(orario = selectedValue OR minuti = selectedValue, "selected", "")%>><%=orario%></option>
					<% ora = DateAdd("n", intervallo, ora) 
					controllo = controllo + intervallo
				wend
				exit do
			end if
		loop %>
	</select>
	<%
end sub


'.................................................................................................
'.. 			taglia la lunghezza di una stringa a partire dal primo spazio o virgola o punto 
'..				presente prima della lunghezza minima prefissata (ricerca inversa)
'..				s:		stringa da troncare
'..				l:		lunghezza massima della stringa risultante
'..				text:	testo da accodare alla stringa troncata (es: "...")
'.................................................................................................
function Sintesi(byVal s, byVal l, end_text)
	dim primo_spazio
	s = cString(s)
	l = cInteger(l)
	if len(s)>l then
		primo_spazio = instr(l, s, " ", vbTextCompare)
		if primo_spazio <= 1 then
			primo_spazio = instr(l,s,vbCr,vbTextCompare)
		end if
		if primo_spazio <= 1 then
			primo_spazio = instr(l,s,vbLf,vbTextCompare)
		end if
		if primo_spazio <= 1 then
			primo_spazio = instr(l,s,"<br>",vbTextCompare)
		end if
		if primo_spazio <= 1 then
			primo_spazio = instr(l,s,",",vbTextCompare)
		end if
		if primo_spazio <= 1 then
			primo_spazio = instr(l,s,";",vbTextCompare)
		end if
		if primo_spazio <= 1 then
			primo_spazio = instr(l,s,".",vbTextCompare)
		end if
		if primo_spazio <= 1 then
			primo_spazio = l
		end if
		Sintesi = left(s,primo_spazio-1) & end_text
	else
		Sintesi = s
	end if
end function

'.................................................................................................
'.. 			taglia la lunghezza di una stringa a partire dal primo spazio 
'..				presente mettendo inizio e fine del messaggio
'..				s:		stringa da troncare
'..				maxlen:		lunghezza massima della stringa risultante
'..				text:	testo da sostituire alla stringa troncata (es: "...")
'.................................................................................................
function collapse(byVal s, byVal maxlen, byVal text)
	dim primo_spazio,messaggio,coda
	if text="" then text=" ~ " end if
	if len(s) <= maxlen then
		messaggio = s ' non serve fare nulla
	else
		primo_spazio = instr(1, s, " ", vbTextCompare)
		if primo_spazio <= 1 or primo_spazio>maxlen then
			primo_spazio = CInt(maxlen/5) '
		end if
		messaggio = left(s,primo_spazio-1) & text
		coda = right(s, maxlen-len(messaggio))
		messaggio = messaggio & coda
	end if
	collapse = messaggio
end function

'.................................................................................................
'..          Formatta il valore secondo le impostazioni date dai parametri
'..			value:				valore da formattare
'..			dec:				numero di posizioni da visualizzare dopo la virgola
'..			con_sep_migliaia:	se vero inserisce anche i separatori delle migliaia nella stringa risultante
'.................................................................................................
function FormatPrice(byVal value, byVal dec, con_sep_migliaia)
	dim result
	if isNumeric(value) then
		if con_sep_migliaia then
			result = FormatNumber(round(value, dec), dec)
		else
			result = FormatNumber(round(value, dec), dec, , vbFalse, vbFalse)
		end if
		
		if not isItalian_SO() then
			'se il sistema non e' in italiano inverte punti e virgole
			result = replace(result, ",", ";")	'; usato come carattere di scambio
			result = replace(result, ".", ",")
			result = replace(result, ";", ".")
		end if
	else
		result = ""
	end if
	FormatPrice = result
end function


'.................................................................................................
'..         Arrotonda il valore alle 2 cifre
'..			value:				valore da arrotondare
'.................................................................................................
function ArrotondaEuro(byVal value)
	ArrotondaEuro = round(value, 2)
end function


'.................................................................................................
'..			Restituisce il nome della lingua richiesta
'..			lingua			codice della lingua di cui recuperare il nome
'.................................................................................................
function GetNomeLingua(byVal lingua)
	
	GetNomeLingua = GetNomeLinguaFrom(lingua, LINGUE_NOMI)
	
end function


'.................................................................................................
'..			Restituisce il nome della lingua indicata
'..			lingua			codice della lingua di cui recuperare il nome
'.................................................................................................
function GetNomeLinguaFrom(lingua, LINGUE)
	dim i
	for i=lbound(LINGUE_CODICI) to uBound(LINGUE_CODICI)
		if LINGUE_CODICI(i) = lingua then
			GetNomeLinguaFrom = LINGUE(i)
			Exit function
		end if
	next
end function


'.................................................................................................
'..			Restituisce "checked" se il primo parametro = TRUE
'..			cond: 		condizione da controllare
'.................................................................................................
Function Chk(cond)
	if CString(cond) = "" then
		chk = ""
	elseif cond then
		chk = " checked "
	else
		chk = ""
	end if
End Function


'.................................................................................................
'..			Restituisce disabled se la condizione risulta vera
'..			condition 		condizione da controllare
'.................................................................................................
function Disable(condition)
	if condition then
		Disable = " disabled "
	else
		Disable = ""
	end if
end function


'.................................................................................................
'..			Imposta la proprieta class del tag e lo disabilita se la condizione e true
'..			condition 		condizione da controllare
'..			class 			valore della proprieta class se il tag non e disabilitato
'.................................................................................................
function DisableClass(condition, classNormal)
	if condition then
		DisableClass = " class="""& classNormal & IIF(classNormal <> "", "_", "") &"disabled"" disabled "
	else
		DisableClass = " class="""& classNormal &""" "
	end if
end function


'**************************************************************************************
'funzione che seleziona il valore da restituire sulla base della lingua corrente (letta da sessione)
'lingua:		lingua da scegliere
'valueIT:		valore in lingua italiana
'valueEN:		valore in lingua inglese
'valueDE:		valore in lingua tedesca
'valueFR:		valore in lingua francese
'valueES:		valore in lingua spagnola
'**************************************************************************************
function ChooseValueByLanguage(lingua, valueIT, valueEN, valueDE, valueFR, valueES)
	ChooseValueByLanguage = ChooseValueByAllLanguages(lingua, valueIT, valueEN, valueDE, valueFR, valueES, "", "", "")
end function


'**************************************************************************************
'funzione che seleziona il valore da restituire sulla base della lingua richiesta
'lingua:		lingua da scegliere
'valueIT:		valore in lingua italiana
'valueEN:		valore in lingua inglese
'valueDE:		valore in lingua tedesca
'valueFR:		valore in lingua francese
'valueES:		valore in lingua spagnola
'valueRU:		valore in lingua russa
'valueCH:		valore in lingua cinese
'**************************************************************************************
function ChooseValueByAllLanguages(lingua, valueIT, valueEN, valueDE, valueFR, valueES, valueRU, valueCN, valuePT)
	if cString(lingua)="" then lingua = LINGUA_ITALIANO end if
	
	Select case lingua
		case LINGUA_SPAGNOLO
			ChooseValueByAllLanguages = valueES
		case LINGUA_FRANCESE
			ChooseValueByAllLanguages = valueFR
		case LINGUA_TEDESCO
			ChooseValueByAllLanguages = valueDE
		case LINGUA_RUSSO
			ChooseValueByAllLanguages = valueRU
		case LINGUA_CINESE
			ChooseValueByAllLanguages = valueCN
		case LINGUA_PORTOGHESE
			ChooseValueByAllLanguages = valuePT	
		case LINGUA_INGLESE
			ChooseValueByAllLanguages = valueEN
		case else
			ChooseValueByAllLanguages = valueIT
	end select
end function


'**************************************************************************************
'versione concisa e performante di ChooseByLanguage (occhio ai prerequisiti)
'dizionario:	oggetto che contiene i valori
'nome:			i nomi devono essere tutti nome_lingua
'lingua:		lingua da selezionare
'SE SI PASSA UN RECORDSET COME DIZIONARIO OCCHIO ALL'EOF
'**************************************************************************************
Function CBLL(ByRef dizionario, ByVal nome, ByVal lingua)
	
	select case LCase(lingua)
		case LINGUA_SPAGNOLO
			CBLL = dizionario(nome &"_es")
		case LINGUA_FRANCESE
			CBLL = dizionario(nome &"_fr")
		case LINGUA_TEDESCO
			CBLL = dizionario(nome &"_de")
		case LINGUA_INGLESE
			CBLL = dizionario(nome &"_en")
		case LINGUA_RUSSO
			CBLL = dizionario(nome &"_ru")	
		case LINGUA_CINESE
			CBLL = dizionario(nome &"_cn")
		case LINGUA_PORTOGHESE
			CBLL = dizionario(nome &"_pt")
		case else
			CBLL = dizionario(nome &"_it")
	end select
	
	if CString(CBLL) = "" AND _
	   LCase(lingua) <> LINGUA_ITALIANO AND LCase(lingua) <> LINGUA_INGLESE then
		CBLL = dizionario(nome &"_en")
	end if
	if CString(CBLL) = "" AND LCase(lingua) <> LINGUA_ITALIANO then
		CBLL = dizionario(nome &"_it")
	end if
End Function




'**************************************************************************************
'Se il campo 'nome' è un campo lingua chiama CBLL
'dizionario:	oggetto che contiene i valori
'nome:			i nomi devono essere tutti nome_lingua
'lingua:		lingua da selezionare
'SE SI PASSA UN RECORDSET COME DIZIONARIO OCCHIO ALL'EOF
'**************************************************************************************
Function CBLE(ByRef dizionario, ByVal nome, ByVal lingua)
	dim abbr,lingue
	lingue = Join(Application("LINGUE"), ", ")
	abbr = Right(Trim(nome),2)
	if InStr(1,lingue,abbr,1) > 0 then
		CBLE = CBLL(dizionario, Left(Trim(nome),(Len(nome)-3)), lingua)
	else
		CBLE = dizionario(nome)
	end if
End Function



'.................................................................................................
'       funzione che moltiplica il nome del campo declinandolo per tutte le lingue
'       field:      campo da moltiplicare
'.................................................................................................
Function FieldLanguageList(FieldList)
    dim lingua, List, i
    List = split(FieldList, ";")
    FieldLanguageList = ""
    for i = lbound(List) to ubound(List)
        for each lingua in Application("LINGUE")
            FieldLanguageList = FieldLanguageList + _
                                IIF(FieldLanguageList<>"", ";", "") + _
                                List(i) + lingua
        next
    next
end function 


'...............................................................................................................
'	funzione che verifica se il campo esiste nel set di risultati del recordset
'	rs			recordset da verificare
'	FieldName	campo da ricercare
'...............................................................................................................
function FieldExists(rs, FieldName)
	dim Field
	FieldExists = false
	for each Field in rs.Fields
		if instr(1, uCase(Field.name), uCase(FieldName), vbTextCompare)>0 AND _
		   len(Trim(Field.name)) = len(Trim(FieldName)) then
			FieldExists = true
			exit for
		end if
	next
end function


'.................................................................................................
'..		restituisce l'indirizzo del contatto
'..		rs:		recordset aperto su un record dell'indirizzario contenente il contatto interessato
'.................................................................................................
function ContactAddress(rs)
	ContactAddress = rs("indirizzoElencoIndirizzi")
	if cstring(rs("localitaElencoIndirizzi"))<>"" then
		ContactAddress = ContactAddress & IIF(cString(ContactAddress)<>"", " - ", "") & rs("localitaElencoIndirizzi")
	end if
	if (CString(rs("capElencoIndirizzi")) <> "" OR CString(rs("cittaElencoIndirizzi")) <> "" OR _
	   CString(rs("statoProvElencoIndirizzi")) <> "" OR cString(rs("countryElencoIndirizzi")) <> "") _
	   AND CString(ContactAddress) <> "" then
	   ContactAddress = ContactAddress + " - "
	end if
	if CString(rs("capElencoIndirizzi")) <> "" then
		ContactAddress = ContactAddress &" "& rs("capElencoIndirizzi")
	end if
	if CString(rs("cittaElencoIndirizzi")) <> "" then
		ContactAddress = ContactAddress &" "& rs("cittaElencoIndirizzi")
	end if
	if CString(rs("statoProvElencoIndirizzi") & rs("countryElencoIndirizzi")) <> "" then
		ContactAddress = ContactAddress + " ("
		if CString(rs("statoProvElencoIndirizzi")) <> "" then
			ContactAddress = ContactAddress & rs("statoProvElencoIndirizzi")
		end if
		if CString(rs("statoProvElencoIndirizzi")) <> "" AND CString(rs("countryElencoIndirizzi")) <> "" then
			ContactAddress = ContactAddress + " - "
		end if
		if CString(rs("countryElencoIndirizzi")) <> "" then
			ContactAddress = ContactAddress &" "& UCase(rs("countryElencoIndirizzi"))
		end if
		ContactAddress = ContactAddress + ")"
	end if
end function


'.................................................................................................
'..		restituisce il nominativo esteso di tutti i dati del contatto o della societa'
'..		rs:		recordset aperto su un record dell'indirizzario contenente il contatto interessato
'.................................................................................................
function ContactFullName(rs)
	dim result
	if rs("isSocieta") then
		result = rs("NomeOrganizzazioneElencoIndirizzi")
		if cString(rs("CognomeElencoIndirizzi") & rs("NomeElencoIndirizzi")) <>"" then
			result = result & " - " & rs("CognomeElencoIndirizzi") & " " & rs("NomeElencoIndirizzi")
		end if
	else
		result = rs("CognomeElencoIndirizzi") & " " & rs("NomeElencoIndirizzi")
		if cString(rs("NomeOrganizzazioneElencoIndirizzi"))<>"" then
			result = result & " - " & rs("NomeOrganizzazioneElencoIndirizzi")
		end if
	end if
	ContactFullName = Result
end function



'.................................................................................................
'..		restituisce il nominativo del contatto o della societa'
'..		rs:		recordset aperto su un record dell'indirizzario contenente il contatto interessato
'.................................................................................................
function ContactName(rs)
	if rs("isSocieta") then
		ContactName = rs("NomeOrganizzazioneElencoIndirizzi")
	else
		ContactName = rs("CognomeElencoIndirizzi") & " " & rs("NomeElencoIndirizzi")
	end if
end function


'.................................................................................................
'..			genera casualmente una stringa di lunghezza lenght composta dai caratteri presenti il charset
'..			charset:		stringa contenente i carattere con i quali comporre la stringa
'..			lenght:			lunghezza della stringa casuale
'.................................................................................................
function GetRandomString(CharSet, lenght)
	dim i, RndIndex
	'inizializza ciclo random
	randomize
	for i = 1 to lenght
		'calcola indice random del carattere
		RndIndex = (rnd()*100000) Mod len(CharSet)
		RndIndex = IIF(RndIndex = 0, len(CharSet), RndIndex)
		'estrae carattere indicato dall'indice random
		GetRandomString = GetRandomString & Mid(CharSet, RndIndex, 1)
	next
end function


'.................................................................................................
'..			inserisce nel campo CodiceInserimento di tb_indirizzario una striga di n caratteri casuali
'..			idCnt:			id del contatto
'.................................................................................................
sub SetCodiceInserimento(conn, idCnt)
	dim codiceInserimento, sql
	
	codiceInserimento = GetRandomString(DOCUMENTS_FILES_CHARSET, 10)
	
	sql = "UPDATE tb_Indirizzario SET codiceInserimento='"&codiceInserimento&"' WHERE IDElencoIndirizzi="&idCnt
	if DB_Type(conn) = DB_SQL then
		sql = sql & " AND ISNULL(codiceInserimento,'')=''"
	else
		sql = sql & " AND ISNULL(codiceInserimento) OR codiceInserimento = ''"
	end if
	CALL conn.execute(sql, , adExecuteNoRecords)
	
	set sql = nothing
	set codiceInserimento = nothing
end sub	
		

'.................................................................................................
'..			Converte un pixel in un em (partendo da un size del carattere di 16).
'..			px:				numero in pixel da convertire
'.................................................................................................
Function PxToEm(px)
	pxToEm = Replace((px / 16), ",", ".")
End Function


'.................................................................................................
'..			Funzione che ritorna autonomamente il percorso relativo per arrivare alla library.
'..			presume che il sito sia installato nella directory application("SERVER_NAME")
'..			e che da li parta la directory amministrazione
'.................................................................................................
function GetLibraryPath()
	GetLibraryPath = GetAmministrazionePath() & "library/"
end function



'.................................................................................................
'..			Funzione che ritorna autonomamente il percorso assoluto dell'area amministrativa
'..			presume che il sito sia installato nella directory application("SERVER_NAME")
'..			e che da li parta la directory amministrazione
'.................................................................................................
function GetAmministrazionePath()
	dim CurrentUrl, CurrentUrlParts
    dim ServerUrl, ServerUrlParts
	dim Levels, BaseUrl
	CurrentUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	CurrentUrlParts = split(CurrentUrl, "/")
	if instr(1, request.ServerVariables("HTTPS"), "on", vbTextCompare) then
        ServerUrl = Application("SECURE_SERVER_NAME")
    else
        ServerUrl = Application("SERVER_NAME")
    end if
    if right(ServerUrl, 1) <> "/" then
        ServerUrl = ServerUrl + "/"
    end if
	ServerUrlParts = split(ServerUrl, "/")

	Levels = uBound(CurrentUrlParts) - ubound(ServerUrlParts)
	if levels > 0 then
		dim i
		BaseUrl = ""
		for i=1 to Levels
			BaseUrl = BaseUrl + "../"
		next
		GetAmministrazionePath = BaseUrl + "amministrazione/"
	else
		GetAmministrazionePath = "amministrazione/"
	end if
end function

'.................................................................................................
'..			Funzione che ritorna autonomamente il percorso assoluto per amministrazione 2
'..			presume che il sito sia installato nella directory application("SERVER_NAME")
'..			e che da li parta la directory amministrazione
'.................................................................................................
function GetAmministrazione2Path()
	dim amministrazionePath
	amministrazionePath = GetAmministrazionePath()
	amministrazionePath = left(amministrazionePath, len(amministrazionePath)-1) + "2/"
	
	GetAmministrazione2Path = amministrazionePath 
end function


'.................................................................................................
'..			Funzione che restituisce il path relativo alla directory "amministrazione" corrente
'.................................................................................................
Function GetAmministrazioneRelativePath()
	dim CurrentUrl, a, i
	CurrentUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	a = InStr(1, currentUrl, "/amministrazione", vbTextCompare)
	if a = 0 then
		GetAmministrazioneRelativePath = ""
	else
		currentUrl = Right(currentUrl, Len(currentUrl) - a - Len("/amministrazione"))
		for i = 1 to UBound(Split(currentUrl, "/"))
			GetAmministrazioneRelativePath = GetAmministrazioneRelativePath + "../"
		next
		GetAmministrazioneRelativePath = server.MapPath(GetAmministrazioneRelativePath)
	end if
End Function


'**************************************************************************************
'funzione che restituisce la pagina corrispondente alla pagina-sito nella lingua richiesta
'conn:			eventuale connessione aperta su NEXT-WEB
'rs				recordset eventualmente creato e chiuso
'PaginaSito		id della pagina sito da decodificare
'Lingua			lingua della quale si vuole la pagina
'**************************************************************************************
function GetPageByLanguage(conn, rs, PaginaSito, Lingua)
	dim CreatedConn, sql
	if CInteger(paginaSito) > 0 then
        if not IsObjectCreated(conn) then
			set conn = server.CreateObject("ADODB.connection")
			conn.open Application("L_conn_ConnectionString")
			CreatedConn = true
		else
			CreatedConn = false
		end if
		sql = "SELECT id_pagDyn_"& Lingua &" FROM tb_pagineSito WHERE id_pagineSito="& cIntero(paginaSito)
		GetPageByLanguage = GetValueList(conn, rs, sql)

		if CreatedConn then	
			conn.close
			set conn = nothing
		end if
	end if
end function


'.................................................................................................
'				funzione che restituisce il nome della pagina corrente
'.................................................................................................
function GetPageName()
	GetPageName = Right(Request.ServerVariables("SCRIPT_NAME"), (Len(Request.ServerVariables("SCRIPT_NAME")) - instrRev(Request.ServerVariables("SCRIPT_NAME"), "/")))
end function



'.................................................................................................
'..			Funzione che ritorna autonomamente l'url completo corrente
'.................................................................................................
function GetCurrentUrl()
	dim CurrentUrl
	CurrentUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
		GetCurrentUrl = "https://" & CurrentUrl
	else
		GetCurrentUrl = "http://" & CurrentUrl
	end if
end function


'.................................................................................................
'..			Funzione che ritorna autonomamente l'url della directory corrente
'.................................................................................................
function GetCurrentBaseUrl()
	dim CurrentUrl
	CurrentUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	CurrentUrl = left(CurrentUrl, (instrrev(CurrentUrl, "/", -1, vbTextCompare)-1))
	if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
		GetCurrentBaseUrl = "https://" & CurrentUrl
	else
		GetCurrentBaseUrl = "http://" & CurrentUrl
	end if
end function



'.................................................................................................
'..			Funzione che ritorna autonomamente l'url completo corrente, compreso di tutti i paramentri
'.................................................................................................
function GetCurrentFullUrl()
	dim CurrentFullUrl
	CurrentFullUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
	if Request.ServerVariables("QUERY_STRING")<>"" then
		CurrentFullUrl = CurrentFullUrl + "?" & Request.ServerVariables("QUERY_STRING")
	end if
	
	if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
		GetCurrentFullUrl = "https://" & CurrentFullUrl
	else
		GetCurrentFullUrl = "http://" & CurrentFullUrl
	end if
end function



'.................................................................................................
'restituisce la versione del next-web attualmente in uso
'la recupera direttamente dalla directory di lavoro del nextweb nel passport
'.................................................................................................
function GetNextWebCurrentVersion(conn, rs)
	dim ConnCreated, sql, value
	'controlla e crea connessione
	
	if instr(1,ucase(Application("L_conn_ConnectionString")),"DBLAYERS4.MDB",vbTextCompare) >0 then
		GetNextWebCurrentVersion = 4
		exit function
	end if
	if instr(1,ucase(Application("L_conn_ConnectionString")),"DBLAYERS.MDB",vbTextCompare) >0 then
		GetNextWebCurrentVersion = 3
		exit function
	end if
	if not IsObjectCreated(conn) then
		set conn = server.createobject("adodb.connection")
		conn.open Application("DATA_ConnectionString")
	else
		ConnCreated = false
	end if
	
	sql = "SELECT TOP 1 sito_dir FROM tb_siti WHERE sito_dir LIKE '%nextweb%' ORDER BY id_sito DESC"
	value = cString(GetValueList(conn, rs, sql))
	value = RemoveInvalidChar(value, "0123456789")
	if cIntero(value) = 0 then
		GetNextWebCurrentVersion = 3
	else
		GetNextWebCurrentVersion = value
	end if
	
	if ConnCreated then
		conn.close
		set conn = nothing
	end if
end function


'.................................................................................................
'restituisce il nome del file su cui gira l'area pubblica
'.................................................................................................
function GetNextWebPageFileName(NextWebVersion)
    if cIntero(NextWebVersion)=0 then
        NextWebVersion =  GetNextWebCurrentVersion(NULL, NULL)
    end if
    Select case NextWebVersion
        case 5
            GetNextWebPageFileName = "default.aspx"
        case else
            GetNextWebPageFileName = "dynalay.asp"
    end select
end function


'.................................................................................................
'restituisce il nome della directory in cui e' presente il nextweb
'.................................................................................................
function GetNextWebDirectory(NextWebVersion)
    if cIntero(NextWebVersion)=0 then
        NextWebVersion =  GetNextWebCurrentVersion(NULL, NULL)
    end if
	if NextWebVersion >= 4 then
		GetNextWebDirectory = "nextWeb" & NextWebVersion
	else
       	GetNextWebDirectory = "nextWeb"
	end if
end function


'.................................................................................................
'restituisce l aparte di sql con il nome della pagina completo
'.................................................................................................
function SQL_PaginaNome(conn)
    SQL_PaginaNome = _
        " (" + SQL_IfIsNull(conn, "nome_ps_it", "''") + ") " + SQL_Concat(conn) + _
        " (" + SQL_If(conn, " (NOT " + SQL_IsNull(conn, "nome_ps_interno") + " AND nome_ps_interno<>'' )", "' ('" + SQL_concat(conn) + " nome_ps_interno" + SQL_concat(conn) + "')'", "''") & ") " + SQL_Concat(conn) + _
		" (' - (lingua: '" + SQL_Concat(conn) + " lingua_nome_it " + SQL_Concat(conn) + " ')')" + SQL_Concat(conn) + _
		" (" + SQL_If(conn, SQL_IsTrue(conn, "template.semplificata"), "' [ usa template per email ] '", "''") + ")" + SQL_Concat(conn) + _
		" (" + SQL_If(conn, SQL_IsTrue(conn, "w.sito_mobile"), "' [ versione mobile ] '", "''") + ")"
end function


'.................................................................................................
'restituisce l aparte di sql con il nome della pagina sito completo
'.................................................................................................
function SQL_PaginaSitoNome(conn, field)
    SQL_PaginaSitoNome = _
        " (" + SQL_IfIsNull(conn, field, "''") + ") " + SQL_Concat(conn) + _
        " (" + SQL_If(conn, " (NOT " + SQL_IsNull(conn, "nome_ps_interno") + " AND nome_ps_interno<>'' )", "' ('" + SQL_concat(conn) + " nome_ps_interno" + SQL_concat(conn) + "')'", "''") & ") "
		'" (" + SQL_If(conn, " (SELECT COUNT(*) FROM tb_pages LEFT JOIN tb_pages tb_templates " + _
		'			 	  			  " ON tb_pages.id_template=tb_templates.id_page " + _
		'				  			  " WHERE " + SQL_IsTrue(conn, "tb_templates.semplificata") + _
		'							  " AND ( tb_pages.id_page = tb_pagineSito.id_pagDyn_IT OR tb_pages.id_page = tb_pagineSito.id_pagStage_IT OR " + _
		'							  	    " tb_pages.id_page = tb_pagineSito.id_pagDyn_EN OR tb_pages.id_page = tb_pagineSito.id_pagStage_EN OR " + _
		'									" tb_pages.id_page = tb_pagineSito.id_pagDyn_FR OR tb_pages.id_page = tb_pagineSito.id_pagStage_FR OR " + _
		'									" tb_pages.id_page = tb_pagineSito.id_pagDyn_DE OR tb_pages.id_page = tb_pagineSito.id_pagStage_DE OR " + _
		'									" tb_pages.id_page = tb_pagineSito.id_pagDyn_ES OR tb_pages.id_page = tb_pagineSito.id_pagStage_ES )) = 0 ", + _
		'			  "''", _
		'			  "' [ usa template per email ] ' ") + ")"
end function


'.................................................................................................
'restituisce l'SQL per elencare le pagine
'.................................................................................................
Function SQL_Pagine(nextWeb_Conn, NextWebVersion, web_id, campoId, campoNome, where, pubblica)
	dim nextWeb_rs, sql, value, nextWeb_Count, nextWeb_Ordine, i, paginePrefisso
	set nextWeb_rs = Server.CreateObject("ADODB.Recordset")
	if pubblica then
		paginePrefisso = "Dyn"
	else
		paginePrefisso = "Stage"
	end if
	
	if NextWebVersion <= 4 then
		
		'composizione query per recupero pagine: scorre siti
		sql = "SELECT * FROM tb_webs"
		if cIntero(web_id)>0 then
			sql = sql + " WHERE id_webs=" & cIntero(web_id)
		end if
		nextWeb_rs.open sql, nextWeb_Conn, adOpenStatic, adLockOptimistic, adCmdText
		nextWeb_Count = 0
		sql = ""
    	
		while not nextWeb_rs.eof
			'scorre lingue attive
			for i=lbound(Application("LINGUE")) to uBound(Application("LINGUE"))
				if Application("LINGUE")(i)<> LINGUA_ITALIANO then
					value = nextWeb_rs("lingua_" & Application("LINGUE")(i))
				else
					value = true
				end if
				if value then
					nextWeb_Ordine = "'" + ParseSQL(nextWeb_rs("nome_webs"), adChar) + " - ' + " + SQL_IfIsNull(nextWeb_Conn, "nome_ps_it", "''") + " + ' - (lingua " + lcase(LINGUE_NOMI(i)) + ")'"
					sql = sql + _
						  " ( " + _
						  " 		SELECT (id_pag" + paginePrefisso + "_" + Application("LINGUE")(i) + ") AS " + campoId + ", " + _
						  " 		(" & IIF(cIntero(web_id)>0, "", "'" + ParseSQL(nextWeb_rs("nome_webs"), adChar) + " - ' + ") + "nome_ps_" + Application("LINGUE")(i) + " + ' - (lingua " + lcase(LINGUE_NOMI(i)) + ")') AS " + campoNome + ", " + _
						  " 		(" + nextWeb_Ordine + ") AS PAGINA_ORDINE " + _
						  " 		FROM tb_pagineSito WHERE id_web=" & nextWeb_rs("id_webs") & " AND NOT (" & SQL_IsNull(nextWeb_Conn, "nome_ps_" + Application("LINGUE")(i)) & ") AND nome_ps_" & Application("LINGUE")(i) & "<>'' " & _
						  " ) UNION "
					nextWeb_Count = nextWeb_Count + 1
				end if
			next
			nextWeb_rs.movenext
		wend
		
		if nextWeb_Count=1 then
			'rimuove union e parentesi
			sql = left(sql, len(sql) - 8)
			sql = right(sql, len(sql) - 2) + _
				  " ORDER BY " + nextWeb_Ordine
		else
			'rimuove solo ultima union
			sql = left(sql, len(sql) - 6) + _
				  " ORDER BY PAGINA_ORDINE "
		end if
		nextWeb_rs.close
		
	else
		
		sql = " SELECT id_pagineSito, lingua_nome_it, p.id_page AS " + campoId + ", ("
		if cIntero(web_id) = 0 then
			sql = sql + "nome_webs " + SQL_concat(nextWeb_Conn) + "' - '" + SQL_concat(nextWeb_Conn)
		end if
		sql = sql + SQL_PaginaNome(nextWeb_Conn) + ") AS " + campoNome + _
			  " FROM (((tb_webs w" + _
			  " INNER JOIN tb_pagineSito ps ON w.id_webs = ps.id_web)" + _
			  " INNER JOIN tb_pages p ON ps.id_pagineSito = p.id_PaginaSito)" + _
			  " INNER JOIN tb_cnt_lingue l ON p.lingua = l.lingua_codice)" + _
			  " LEFT JOIN tb_pages template ON p.id_template = template.id_page" + _
			  " WHERE (1=0"
		for each i in Application("LINGUE")
			sql = sql + " OR ps.id_pag" + paginePrefisso + "_" + i + " = p.id_page AND NOT " & SQL_IsNull(nextWeb_Conn, "nome_ps_" + i) & " AND nome_ps_" & i & "<>'' "
		next
		sql = sql + ")"
		
		if cIntero(web_id)>0 then
			sql = sql + " AND w.id_webs=" & web_id
		end if
		
		sql = sql + where
		
		sql = sql + " ORDER BY nome_webs, nome_ps_it, id_pagineSito, p.id_page"
		
	end if
	
	SQL_Pagine = sql
End Function


'.................................................................................................
'estrae la porzione di url relativa a dominio e/o directory nella quale gira la dynalay.
'   es di input:                                                    -->     es di output:
'   http://www.turismovenezia.it/dynalay.asp?PAGINA=405&...         -->     http://www.turismovenezia.it/
'   http://www.venetoinside.com/default.aspx?PAGINA=555&...         -->     http://www.venetoinside.com/
'   http://nmobile/turismovenezia/dynalay.asp?PAGINA=405&...        -->     http://nmobile/turismovenzia/
'   http://www.next-aim.net/hotelflora/dynalay.asp?PAGINA=405&...   -->     http://www.next-aim.net/hotelflora/
' FUNZIONA SOLO PER URL DELLE PAGINE VISIBILI (nextweb: versione 3, 4 e 5)
'.................................................................................................
function ExtractPageBaseUrl(URL)
    dim tmpUrl
    tmpUrl = left(URL, instr(1, URL, "?", vbTextCompare))
    tmpUrl = left(URL, instrrev(URL, "/", -1, vbTextCompare))
    ExtractPageBaseUrl = tmpUrl
end function


'.................................................................................................
'..            Restituisce l'URL completo (http//...) della pagina dato la paginaSito richiesta
'..            per l'apertura nella lingua indicata
'.................................................................................................
Function GetPageSiteUrl(connWeb, paginaSito, LINGUA)
	GetPageSiteUrl = GetPageURL(connWeb, GetPageByLanguage(connWeb, NULL, PaginaSito, LINGUA))
end function



Function GetUrl()
	dim url
	url = "http://" & Request.ServerVariables("SERVER_NAME")
	GetUrl = url
End Function


'.................................................................................................
'..				Restituisce l'URL completo del sito richiesto
'.................................................................................................
Function GetSiteUrl(connWeb, WebId, NextWebVersion)
	dim created, sql, rs
    created = not IsObjectCreated(connWeb)
	
	if cIntero(WebId)=0 then 
       	'applicazione locale o sito non individuabile
        GetSiteUrl = "http://" & Application("SERVER_NAME")
  	else
		if created then
			set connWeb = server.createobject("adodb.connection")
			Select case NextWebVersion
                case 5
                    connWeb.open Application("DATA_ConnectionString")
                case else
                    connWeb.open Application("L_conn_ConnectionString")
            end select
		end if
		set rs = server.createObject("adodb.recordset")
        
		if cIntero(NextWebVersion) = 0 then
			NextWebVersion = cInteger(GetNextWebCurrentVersion(NULL, rs))
		end if
			
		sql = "SELECT * FROM tb_webs WHERE id_webs=" & cIntero(WebId)
		rs.open sql, connWeb, adOpenStatic, adLockOptimistic, adCmdText
		if not rs.eof then
			Select case NextWebVersion
                case 5
                    GetSiteUrl = rs("Url_base")
                case else
                    GetSiteUrl = "http://" & rs("nome_webs")
            end select
		else
			GetSiteUrl = ""
		end if
		rs.close
		
		set rs = nothing
		if created then
			connWeb.Close
			set connWeb = nothing
		end if
	end if
end function


' Giacomo 10/10/2012
function GetSiteUrlImages()
	dim url
	url = GetSiteUrl(null, null, null) & "/upload/" & Application("AZ_ID") & "/images/"
	GetSiteUrlImages = url
end function


'.................................................................................................
'..				Restituisce l'URL completo del sito richiesto a partire da un idPaginaSito
'.................................................................................................
Function GetSiteBaseUrl(connWeb, idPaginaSito)
	dim sql
	sql = "SELECT id_web FROM tb_pagineSito WHERE id_pagineSito = " & idPaginaSito
	GetSiteBaseUrl = GetSiteUrl(connWeb, cIntero(GetValueList(connWeb,NULL,sql)), 0)
end function


'.................................................................................................
'..				Restituisce l'idPaginaSito dato l'id_Page
'.................................................................................................
Function GetPaginaSitoIdByPaginaId(connWeb, idPagina)
	dim sql
	sql = "SELECT id_PaginaSito FROM tb_pages WHERE id_page = " & idPagina
	GetPaginaSitoIdByPaginaId = cIntero(GetValueList(connWeb,NULL,sql))
end function


'.................................................................................................
'..  Restituisce l'url completo della pagina richiesta con relativo querystring per caricamento
'    nella lingua corrispondente alla pagina (se vuota la imposta direttamente ad ITALIANO)
'.................................................................................................
Function GetPageURL(connWeb, PAGINA)
	dim created, sql, rs, NextWebVersion
    created = not IsObjectCreated(connWeb)
	if created then
		set connWeb = server.createobject("adodb.connection")
		connWeb.open Application("L_conn_ConnectionString")
	end if
	set rs = server.createObject("adodb.recordset")
	
    sql = " SELECT tb_webs.*, tb_pages.lingua " + _
          " FROM tb_pages LEFT JOIN tb_webs ON tb_pages.id_webs = tb_webs.id_webs " + _
          " WHERE id_page=" & cIntero(PAGINA)
    rs.open sql, connWeb, adOpenStatic, adLockOptimistic, adCmdText
    if not rs.eof then
        NextWebVersion = cInteger(GetNextWebCurrentVersion(NULL, NULL))
		
		GetPageUrl = GetSiteUrl(connWeb, rs("id_webs"), NextWebVersion)
		
		GetPageUrl = GetPageUrl & "/" & GetNextWebPageFileName(NextWebVersion) & "?PAGINA=" & PAGINA
        if cString(rs("lingua"))<>"" then
            GetPageUrl = GetPageUrl & "&LINGUA=" & rs("lingua")
        else
            GetPageUrl = GetPageUrl & "&LINGUA=" & LINGUA_ITALIANO
        end if
        'corregge eventuali doppioni
        GetPageUrl = replace(GetPageUrl, "//", "/")
        GetPageUrl = replace(GetPageUrl, ":/", "://")
        
    end if 
    
    rs.close
	set rs = nothing
	if created then
		connWeb.Close
		set connWeb = nothing
	end if
End Function


'.................................................................................................
  ' Description:
  '   Sorts a dictionary by either key or item
  ' Parameters:
  '   objDict - the dictionary to sort
  '   intSort - the field to sort (1=key, 2=item)
  ' Returns:
  '   A dictionary sorted by intSort
'.................................................................................................
Function SortDictionary(objDict, intSort)

    ' declare constants
    Const dictKey  = 1
    Const dictItem = 2

    ' declare our variables
    Dim strDict()
    Dim objKey
    Dim strKey,strItem
    Dim X,Y,Z

    ' get the dictionary count
    Z = objDict.Count

    ' we need more than one item to warrant sorting
    If Z > 1 Then
        ' create an array to store dictionary information
        ReDim strDict(Z,2)
        X = 0
        ' populate the string array
        For Each objKey In objDict
            strDict(X,dictKey)  = CStr(objKey)
            strDict(X,dictItem) = CStr(objDict(objKey))
            X = X + 1
        Next
        
        ' perform a a shell sort of the string array
        For X = 0 To (Z - 2)
            For Y = X To (Z - 1)
                If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
                    strKey  = strDict(X,dictKey)
                    strItem = strDict(X,dictItem)
                    strDict(X,dictKey)  = strDict(Y,dictKey)
                    strDict(X,dictItem) = strDict(Y,dictItem)
                    strDict(Y,dictKey)  = strKey
                    strDict(Y,dictItem) = strItem
                End If
            Next
        Next
        
        ' erase the contents of the dictionary object
        objDict.RemoveAll
        
        ' repopulate the dictionary with the sorted information
        For X = 0 To (Z - 1)
            objDict.Add strDict(X,dictKey), strDict(X,dictItem)
        Next
    
    End If
    
    
End Function


'sceglie il campo dal recordset se il form non postato
Function CBR(rs, nomeCampoRecordset, prefisso)
	CBR = CBRR(rs, nomeCampoRecordset, prefisso & nomeCampoRecordset)
End Function


'sceglie il campo dal recordset se il form non postato
Function CBRR(rs, nomeCampoRecordset, nomeCampoRequest)
	if Request.ServerVariables("REQUEST_METHOD") = "POST" OR _
	   IsNull(rs) then
		if LCase(Left(nomeCampoRequest, 4)) = "chk_" then
			CBRR = (request(nomeCampoRequest) <> "")
		else
			CBRR = request(nomeCampoRequest)
		end if
	else
	'response.write nomeCampoRecordset
		CBRR = rs(nomeCampoRecordset)
	end if
End Function


'sceglie il primo parametro se il form non postato
Function CBRV(val, valForm)
	if Request.ServerVariables("REQUEST_METHOD") = "POST" then
		CBRV = valForm
	else
		CBRV = val
	end if
End Function


'Restituisce la class css con i tre valori "first", "alternate", "last" a seconda dell'indice in ingresso.
'nel caso siano presenti più colonne setta anche le relative classi per gli estremi di ciascuna colonna e riga.
'index:				indice in base 1 dell'elemento corrente
'indexLast:			indice in base 1 dell'ultimo elemento
'columnsNumber:		eventuale numero di colonne
Function GetCssClass(ByVal index, indexLast, columnsNumber)
	'siccome la procedura e con indice in base 0 tolgo uno
	index = index - 1
	indexLast = indexLast - 1
	
	if index = 0 then
        GetCssClass = " first"
	end if
    if index = indexLast then
        GetCssClass = GetCssClass &" last"
	end if

    if columnsNumber < 2 AND (index + 1) MOD 2 = 0 then
        GetCssClass = GetCssClass &" alternate"
    elseif columnsNumber > 1 then												'pi� colonne
        'classe per colonna
        if index MOD columnsNumber = 0 then
            GetCssClass = GetCssClass &" firstcolumn"
		end if
        if (index + 1) MOD columnsNumber = 0 then
            GetCssClass = GetCssClass &" lastcolumn"
		end if
        if (index MOD columnsNumber) MOD 2 = 1 AND (index + 1) MOD columnsNumber <> 1 then
            GetCssClass = GetCssClass &" alternatecolumn"
		end if

        'classe per riga
        if index < columnsNumber then
            GetCssClass = GetCssClass &" firstrow"
		end if
        if index >= (indexLast \ columnsNumber * columnsNumber) then			'x/y * x != x
            GetCssClass = GetCssClass &" lastrow"
		end if
        if (index \ columnsNumber) MOD 2 = 1 then
            GetCssClass = GetCssClass &" alternaterow"
		end if
    end if
	
	GetCssClass = LTrim(GetCssClass)
End Function


'**************************************************************************************************************************************
'funzione che esegue i controlli di base per la tracciabilita' della richiesta nei contatori
'**************************************************************************************************************************************
function ActionLoggable(conn)
	'verifica filtri di base per validita' log delle visite
	if instr(1, Request.ServerVariables("SCRIPT_NAME"), "nextWeb", vbTextCompare)>0 then
		'dynalay del next-web
		ActionLoggable = false
	else
		'dynalay del sito
		if request.querystring("HTML_FOR_EMAIL") <> "" then
			'pagina richiesta via email
			ActionLoggable = false
		else
			'pagina richiesta da user agent
			if Request.ServerVariables("LOCAL_ADDR") = Request.ServerVariables("REMOTE_ADDR") then
				'pagina richiesta dallo stesso IP (stesso server)
				ActionLoggable = false
			else
				'pagina richiesta da altro IP
				if instr(1, Request.ServerVariables("SCRIPT_NAME"), "default.asp", vbTextCompare)>0 then
					'filtri per default.asp
					if instr(1, Request.ServerVariables("REQUEST_METHOD"), "GET", vbTextCompare)<1 then
						'tipo di richiesta della pagina non valido
						ActionLoggable = false
					else
						'tipo di richiesta valido.
						ActionLoggable = true
					end if
				elseif instr(1, Request.ServerVariables("SCRIPT_NAME"), "dynalay.asp", vbTextCompare)>0 then
					'filtri per dynalay.asp
					if Request.ServerVariables("QUERY_STRING")="" OR _
					   request.querystring("PAGINA")="" then
						'parametri della dynalay non validi
						ActionLoggable = false
					else
						'parametri della dynalay validi
						ActionLoggable = true
					end if
				else
					'script sconosciuto
					ActionLoggable = true
				end if
			end if
		end if
	end if
	
	if ActionLoggable then
		'effettua verifiche nei filtri interni e generali.
		dim connCreated, rs, sql
		connCreated = not IsObjectCreated(conn)
		if connCreated then
			set conn = server.createobject("adodb.connection")
			conn.open Application("DATA_ConnectionString")
		end if
		set rs = Server.CreateObject("ADODB.Recordset")
	
		'controllo filtri
		sql = "SELECT * FROM tb_contents_log_filtri"
		rs.open sql, conn, adOpenStatic, adLockOptimistic, adCmdText
		ActionLoggable = true
		while not rs.eof AND ActionLoggable
			select case rs("fil_tipo")
				case FILTRO_TEXT_FULLTEXT
					ActionLoggable = InStr(1, Request.ServerVariables(rs("fil_parametro")), rs("fil_valore"), vbTextCompare) = 0
				case FILTRO_TEXT_INIZIO
					ActionLoggable = UCase(Left(Request.ServerVariables(rs("fil_parametro")), Len(rs("fil_valore")))) <> UCase(rs("fil_valore"))
				case FILTRO_TEXT_FINE
					ActionLoggable = UCase(Right(Request.ServerVariables(rs("fil_parametro")), Len(rs("fil_valore")))) <> UCase(rs("fil_valore"))
				case else
					ActionLoggable = UCase(Request.ServerVariables(rs("fil_parametro"))) <> UCase(rs("fil_valore"))
			end select
			rs.movenext
		wend
		rs.close
	
		set rs = nothing
		if connCreated then
			conn.Close
			set conn = nothing
		end if
	end if
	
end function


'**************************************************************************************************************************************
'funzione che restituisce il tipo di user agent riconosciuto dal sistema:
'		contUtenti:		contatore dei BROWSER degli utenti
'		contCrawler:	contatore dei CRAWLER e motori di ricerca
'		contALTRO:		contatore degli user-agent NON RICONOSCIUTI
'**************************************************************************************************************************************
function GetUserAgentType()
	Select case Session("USER_AGENT_TYPE") 
		case CRAWLER
			GetUserAgentType = CRAWLER
		case BROWSER
			GetUserAgentType = BROWSER
		case UNRECOGNIZED
			GetUserAgentType = UNRECOGNIZED
		case else
			'user agent non ancora riconosciuto: esegue riconoscimento
			if instr(1, Request.ServerVariables("SCRIPT_NAME"), "default.asp", vbTextCompare)>0 AND _
			   Request.ServerVariables("QUERY_STRING")<>"" AND _
			   request.querystring("PS")="" AND request.querystring("LINGUA")="" AND request.querystring("PAGINA")="" then
				'parametri della default.asp non validi: in questo modo usati solo dai crawler
				GetUserAgentType = CRAWLER
			else
				'parametri default.asp corretti			
				if instr(1, request.ServerVariables("HTTP_USER_AGENT"), "bot", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "craw", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "spider", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "wget", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "url", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "test", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "check", vbTextCompare)>0 OR _
				   instr(1, request.ServerVariables("HTTP_USER_AGENT"), "grub", vbTextCompare)>0 then
					'presente indicazione crawler o boot
				   	GetUserAgentType = CRAWLER
				else
					'indicazione non valida
					if Trim(request.ServerVariables("HTTP_USER_AGENT"))="" then
						'user agent non valido
						GetUserAgentType = UNRECOGNIZED
					else
						'user agent valido
						if instr(1, request.ServerVariables("server_protocol"), "http/1.1", vbTextCompare)>0 then
							'Protocollo http 1.1 valido
							GetUserAgentType = BROWSER
						else
							'altro protocollo (da verificare)
							if instr(1, request.ServerVariables("HTTP_CONNECTION"), "keep-alive", vbTextCompare)>0 AND _
							   request.ServerVariables("HTTP_ACCEPT")<>"" then
							   	'connessione http "alive" e tipi di media accettati specificati
								GetUserAgentType = BROWSER
							else
								'connessione http close o non dichiarata e media non dichiarati
								if instr(1, request.ServerVariables("HTTP_CONNECTION"), "keep-alive", vbTextCompare)>0 OR _
								   request.ServerVariables("http_accept_language")<>"" then
								   'connessione http "alive" e dichiarazione delle lingue dei contenuti accettati
								   GetUserAgentType = BROWSER
								else
									GetUserAgentType = CRAWLER
								end if
							end if
						end if
					end if
				end if
			end if
			Session("USER_AGENT_TYPE") = GetUserAgentType
	end select
end function


'**************************************************************************************************************************************
'Spedisce una email a support
'**************************************************************************************************************************************
Sub SendEmailSupport(text)
	CALL SendEmailSupportEX("Errore: "& Request.ServerVariables("SERVER_NAME"), text)
end sub

Sub SendEmailSupportEX(subject, text)
	CALL SendEmailSupportEXAttach(subject, text, "")
End Sub

Sub SendEmailSupportEXAttach(subject, text, attachmentPath)
	CALL SendEmailTO("sviluppo@combinario.com", "sviluppo@combinario.com", "", subject, text, attachmentPath)
end sub

Sub SendEmailTO(sender, dest, cc, subject, text, attachmentPath)
	dim configuration, message
	
	'configurazione
	if IsObject(Application("Class_mailer_Configuration")) then
		Set Configuration = Application("class_mailer_configuration")
	else
		Set Configuration = Server.CreateObject("CDO.Configuration")
		'configurazione di base messaggio
		with Configuration.Fields
			.Item(cdoSMTPServer) = Request.ServerVariables("SERVER_NAME")
			.Item(cdoNNTPAuthenticate) = cdoAnonymous
			.Item(cdoSendUsingMethod) = cdoSendUsingPort
			.Item(cdoURLGetLatestVersion) = true
			.update
		end with
	end if
	
	Set Message = Server.CreateObject("CDO.Message")
	if IsLocal() then
		message.to = "sviluppo@combinario.com"
		message.from = "sviluppo@combinario.com"
		text  = text + vbCrlf + vbCrlf + vbCrlf + _
				 "***************************************************************" + vbcrlf + _
				 "DATI ORIGINALI DEL MESSAGGIO:" + vbcrlf + _
				 "***************************************************************" + vbcrlf + _
				 "from: " & sender & vbcrlf + _
				 "to:" & dest & vbcrlf + _
				 "cc:" & cc & vbcrlf + _
				 vbCrlf + "***************************************************************" + vbcrlf
	else
		message.from = sender
		if dest <>"" then
			message.to = dest
		else
			message.to = sender
		end if
		if cc<>"" then
			message.cc = cc
		end if
	end if
	
	message.subject = subject
	message.TextBody = text
	if attachmentPath <> "" then
		message.AddAttachment attachmentPath
	end if
	
	set Message.Configuration = Configuration
	Message.Send
	
	set message = nothing
	set configuration = nothing
End Sub


'**************************************************************************************************************************************
'	imposta la sessione della lingua
'**************************************************************************************************************************************
Sub SetSessionLingua(newLingua)
	select case lcase(Trim(newLingua))
		case LINGUA_ITALIANO
			Session("LINGUA") = LINGUA_ITALIANO
		case LINGUA_INGLESE
			Session("LINGUA") = LINGUA_INGLESE
		case LINGUA_SPAGNOLO
			Session("LINGUA") = LINGUA_SPAGNOLO
		case LINGUA_TEDESCO
			Session("LINGUA") = LINGUA_TEDESCO
		case LINGUA_FRANCESE
			Session("LINGUA") = LINGUA_FRANCESE
		case else
			' non modifica la lingua.
	end select
end sub


'rimuove il carattare dall'inizio e dalla fine della stringa
Function TrimChar(byVal str, c)
	TrimChar = str
	if Left(TrimChar, 1) = c then
		TrimChar = Right(TrimChar, Len(TrimChar) - 1)
	end if
	if Right(TrimChar, 1) = c then
		TrimChar = Left(TrimChar, Len(TrimChar) - 1)
	end if
End Function


'restituisce il valore tipizzato (vedi descrittori)
Function DesValore(tipo, valore, memo)
	select case tipo
		'testo normale, link ad un file, link ad una risorsa esterna, colore valido, anagrafiche selezionate dal NEXT-com (potrebbe essere un array ma in vbscript ci sono pi� funzioni per le stringhe)
		case adVarChar, adBinary, adUserDefined, adPropVariant, adIUnknown, adVarBinary, adChar
			DesValore = CString(valore)
		'numero, valuta, pagine selezionate dal NEXT-web, collegamento all'indice, rubrica, amministratore
		case adNumeric,	adCurrency, adGUID, adChapter, adIDispatch, adSingle
			DesValore = CReal(valore)
		case adBoolean												'valore true/false
			DesValore = CBoolean(valore, false)
		case adDate													'valore data
			if IsDate(valore) then
				DesValore = CDate(valore)
			else
				DesValore = ""
			end if
		case adLongVarChar											'testo lungo
			DesValore = CString(memo)
		case adDouble												'valori numerici min/max
			DesValore = Array(CReal(valore), CReal(memo))
	end select
End Function

'GetParameter
Function GetModuleParam(conn, code)
	GetModuleParam = GetModuleParamExtended(conn, code, LINGUA_ITALIANO, true)
End Function


'restituisce il valore tipizzato del parametro del modulo del NextPassport
Function GetModuleParamExtended(conn, code, lingua, checkSession)
	dim rs, sql, valoreParam
	set rs = server.createobject("adodb.recordset")
	
	if Session(code)="" OR not checkSession or true then
		sql = " SELECT * FROM tb_siti_descrittori d" & _
			  " INNER JOIN rel_siti_descrittori r ON d.sid_id = r.rsd_descrittore_id" & _
			  " WHERE sid_codice LIKE '"& code &"'"
		rs.open sql, conn, adOpenForwardOnly, adLockReadOnly
		if not rs.eof then
			valoreParam = DesValore(rs("sid_tipo"), CBLL(rs, "rsd_valore", lingua), CBLL(rs, "rsd_memo", lingua))
		end if
		rs.close
		if valoreParam = "" then 'cerco tra i vecchi parametri
			sql = "SELECT * FROM tb_siti_parametri WHERE par_key LIKE '"& code &"'"
			rs.open sql, conn, adOpenForwardOnly, adLockReadOnly
			if not rs.eof then
				valoreParam = rs("par_value")
			end if
			rs.close
		end if
		GetModuleParamExtended = valoreParam
	else
		GetModuleParamExtended = Session(code)
	end if
	
	set rs = nothing
End Function


'restituisce true se gli SMS sono abilitati
Function SMSAbilitati(conn)
	SMSAbilitati = GetModuleParam(conn, "SMS_ABILITATI")
End Function


'restituisce true se gli SMS sono abilitati
Function FaxAbilitati(conn)
	FaxAbilitati = GetModuleParam(conn, "FAX_ABILITATI")
End Function


'restituisce true se la struttura di tagging dell'indice e' abilitata
Function TagAbilitati(conn)
	TagAbilitati = GetModuleParamExtended(conn, "TAGS_ABILITED", LINGUA_ITALIANO, false)
End Function


'restituisce lo stato di attivazione dell'https per la pagina attuale.
Function IsHttpsActive()
	IsHttpsActive = ( instr(1, request.ServerVariables("HTTPS"), "on", vbTextCompare) > 0 )
end function


'esegue la richiesta http all'url indicato.
'	url:	url da recuperare
Function ExecuteHttpUrl(url)
	ExecuteHttpUrl = ExecuteHttpUrlEx(url, false)
end function


'esegue la richiesta http all'url indicato.
'	url:	url da recuperare
'	checkstatus:	se il recupero fallisce rimanda stringa vuota.
Function ExecuteHttpUrlEx(url, checkstatus)
	dim xmlhttp
	
	'crea richiesta
	set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	response.write vbCrLF & "<!--" & url & "-->" & vbCrLF
	
	xmlhttp.setTimeouts 200000, 200000, 200000, 200000
	CALL xmlhttp.open("GET", url, false)
	
	'invia richiesta
	CALL xmlhttp.send()
	
	'recupera valore
	if checkstatus then
		If xmlhttp.Status >= 400 And xmlhttp.Status <= 599 Then
			'errore generazione
			ExecuteHttpUrlEx = ""
		else
			ExecuteHttpUrlEx = xmlhttp.responseText
		end if
	else
		ExecuteHttpUrlEx = xmlhttp.responseText
	end if
	
	set xmlhttp = nothing
	
end function


'esegue la richiesta http all'url indicato restituendo lo stream con il contenuto
'url:	url da richiedere
Function ExecuteHttpUrlGetStream(url)
	
	'genera richiesta
	response.write vbCrLF & "<!--" & url & "-->" & vbCrLF
	dim objMessage, sBody, utf8Stream
	Set objMessage = Server.CreateObject("CDO.Message")
	objMessage.CreateMHTMLBody URL, CdoSuppressAll
	'recupera stream con html
	set sBody = objMessage.HTMLBodyPart.GetDecodedContentStream
	
	'crea nuovo stream per convertire in utf-8
	set utf8Stream  = Server.CreateObject("ADODB.Stream")
	utf8Stream.Open
	utf8Stream.Type = adTypeText
	utf8Stream.charset = "UTF-8"
	utf8Stream.writeText sBody.ReadText(adReadAll)
	'utf8Stream.savetofile(Application("IMAGE_PATH") & "\" & DateTimeIso(Now()) & ".html")
	
	set ExecuteHttpUrlGetStream = utf8Stream
	
end function


'Giacomo 18/04/2013
Function ExecuteUrlSaveFile(url, path)
	dim xml, oStream
	Set xml = CreateObject("Microsoft.XMLHTTP")
	xml.Open "GET", URL, False
	xml.Send

	set oStream = createobject("Adodb.Stream")
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
	Const adSaveCreateNotExist = 1

	oStream.type = adTypeBinary
	oStream.open
	oStream.write xml.responseBody

	' Do not overwrite an existing file
	'oStream.savetofile path, adSaveCreateNotExist

	' Use this form to overwrite a file if it already exists
	 oStream.savetofile path, adSaveCreateOverWrite

	oStream.close

	set oStream = nothing
	Set xml = Nothing
end function



'funzione che restituisce true in caso l'installazione sia locale.
function IsLocal()
	if instr(1, request.serverVariables("HTTP_HOST"), ".local", vbTextCompare)>0 then
		IsLocal = true
	else
		IsLocal = false
	end if
end function

'funzione che verifica se la richiesta è locale (dalla nostra rete, o dal server stesso)
function IsFromLocal()
	if IsLocal() OR _
	   instr(request("REMOTE_ADDR"), "127.0.0.1") OR _
	   instr(request("REMOTE_ADDR"), "192.168.20.") OR _
	   request("REMOTE_ADDR") = request("LOCAL_ADDR") OR _
	   instr(request("REMOTE_ADDR"), "192.168.20.") OR _
	   instr(request("REMOTE_ADDR"), "79.60.233.19") then
		IsFromLocal = true
	else
		IsFromLocal = false
	end if
end function

'funzione che restituisce true in caso l'installazione sia locale o l'utente sia next-aim.
function IsNextAim()
	IsNextAim = IsCombinario()
end function


'funzione che restituisce true in caso l'installazione sia locale o l'utente sia combinario o next-aim
function IsCombinario()
	if IsLocal() OR instr(1, Session("LOGIN_4_LOG"), "NEXTAIM", vbTextCompare)>0 OR instr(1, Session("LOGIN_4_LOG"), "COMBINARIO", vbTextCompare)>0then
		IsCombinario = true
	else
		IsCombinario = false
	end if
end function


'funzione che restituisce true in caso si sia nell'area amministrativa
function IsAmministrazione()
	if instr(1, GetCurrentUrl(), "/amministrazione", vbTextCompare)>0 then
		IsAmministrazione = true
	else
		IsAmministrazione = false
	end if
end function




'scrive dentro log_framework
function WriteLogAdmin(conn, tab_name, rec_id, cod, desc)
	CALL WriteLogAdminHttp(conn, tab_name, rec_id, cod, desc, true)
end function


'scrive dentro log_framework
function WriteLogAdminHttp(conn, tab_name, rec_id, cod, desc, httpLog)

	dim httpRaw
	
	if httpLog then
		httpRaw = GetRawHttp()
	else
		httpRaw = ""
	end if
	
	CALL WriteLogAdminHttpRaw(conn, tab_name, rec_id, cod, desc, httpRaw)
	
end function
	
'scrive nel log di sistema
function WriteLogAdminHttpRaw(conn, tab_name, rec_id, cod, desc, httpRaw)
	dim sql, connCreated
	
	if not IsObjectCreated(conn) then
		connCreated = true
		set conn = server.createobject("adodb.connection")
		conn.open Application("DATA_ConnectionString"),"",""
	else
		connCreated = false
	end if
	
	if cString(Session("ID_SITO"))<>"" then
		sql= _
		"INSERT INTO log_framework "+_
		"(log_table_nome,log_record_id,log_codice,log_descrizione,log_data,log_Admin_id,"+_
		"log_http_request,log_application_id) VALUES "+_
		"('" & ParseSql(tab_name, adChar) &"',"&rec_id&",'"&ParseSql(Left(cod, 50), adChar)&"','"&ParseSql(Left(desc, 255),adChar)&"',"&SQL_Now(conn)&","&cIntero(Session("ID_ADMIN"))&",'"+_
		ParseSql(httpRaw, adChar) & "',"&cString(Session("ID_SITO"))&");"
	else
		sql= _
		"INSERT INTO log_framework "+_
		"(log_table_nome,log_record_id,log_codice,log_descrizione,log_data,log_Admin_id,"+_
		"log_http_request) VALUES "+_
		"('"&ParseSQL(tab_name,adChar)&"',"&rec_id&",'"&ParseSQL(Left(cod, 50),adChar)&"','"&ParseSql(Left(desc, 255),adChar)&"',"&SQL_Now(conn)&","&cIntero(Session("ID_ADMIN"))&",'"+_
		ParseSql(httpRaw, adChar) & "');"
	end if	
	conn.execute(sql)	
	
	if connCreated then
		conn.close
		set conn = nothing
	end if
end function

'restituisce l'http raw della richiesta corrente
function GetRawHttp()
	dim variabile
	for each variabile in request.serverVariables
		GetRawHttp = GetRawHttp & variabile &"="""& replace(request.serverVariables(variabile), "|", " ") &"""|"
	next
end function



'**************************************************************************************************************************************
'funzione scrive il pulsante che porta al form per esportare le anagrafiche in una rubrica, a partire da qualsiasi tipo di elenco
'		sql:			 	query utilizzata per ricavare l'id delle anagrafiche da esportare nella rubrica
'		campo_id:		 	campo id presente nella query passata, deve essere il campo IDElencoIndirizzi oppure una chiava esterna che punta ad esso
'		alternativeLabel: 	NON OBBIGATORIO
'		alternativePath: 	NON OBBIGATORIO
'**************************************************************************************************************************************
function ExportContattiInRubrica(sql, campo_id, alternativeLabel, alternativePath)
	dim path, label
	Session("sql_export_in_rubrica") = sql
	Session("campo_id_export_in_rubrica") = campo_id
	
	
	label = "EXPORT CONTATTI IN RUBRICA"
	if cString(alternativeLabel)<>"" then
		label = Ucase(alternativeLabel)
	end if
	
	
	if cString(alternativePath)<>"" then
		path = alternativePath & "ExportContattiInRubrica.asp"
	else
		path = "../library/ExportContattiInRubrica.asp"
	end if
		
	%>
	<a style="width:100%; text-align:center; line-height:12px;" class="button"
		title="Apre la palette di export delle anagrafiche in una rubrica" 
		onclick="OpenAutoPositionedScrollWindow('<%=path%>', 'export', 240, 142, true);" href="javascript:void(0);">
		<%=label%>
	</a>
	<%
end function								
										
										

										
'funzione che aggiorna la tabella sentinella specificata con la data e l'ora attuali
function UpdateSentinelTable(sentinel)
	
	dim conn,sql
	set conn = server.createobject("adodb.connection")
	conn.open Application("DATA_ConnectionString"),"",""	
	
	sql="UPDATE " & sentinel & " SET sent_time=" 
	
	Select case DB_Type(conn)		
		case DB_SQL		
			sql=sql & "GETDATE();"
		case DB_Access
			sql=sql & "NOW();"	
	end select	
		
	conn.execute sql	
end function								
																				
				



'*******************************************************************************
'restituisce true se il sistema di inline css è attivo.
'*******************************************************************************
function IsPreMailerRenderingActive()
	if cString(Application("PREMAILER_PROXY_SERVER"))<>"" then
		Application("UsePremailerRendering") = true
	else
		if cstring(Application("UsePremailerRendering"))="" OR not cBoolean(Application("UsePremailerRendering"), false) then
			'stato di attivazione non presente in cache, va a verificarlo su web.config
			if GetNextWebCurrentVersion(null, null) > 4 then
				'verifica se l'handler è attivo		
				if instr(1, ExecuteHttpUrl("http://" + Application("SERVER_NAME") + "/" + PREMAILER_HANDLER), "NEXTMAILPARSER ACTIVE", vbTextCompare)>0 then
					Application("UsePremailerRendering") = true
				else
					Application("UsePremailerRendering") = false
				end if
			else
				'non è next-web con parte pubblica in .net, non ha il prerendering
				Application("UsePremailerRendering") = false
			end if
		end if
	end if	
	IsPreMailerRenderingActive = Application("UsePremailerRendering")
end function



'*******************************************************************************
'prepara l'url per essere passato al parser css per email.
'*******************************************************************************
function EncodeCssInlinedUrl(url)
	if cString(Application("PREMAILER_PROXY_SERVER"))<>"" then
		EncodeCssInlinedUrl = "http://" + Application("PREMAILER_PROXY_SERVER") + "/" + PREMAILER_HANDLER + "?url=" + Server.UrlEncode(url)
	elseif IsPreMailerRenderingActive() then
		EncodeCssInlinedUrl = "http://" + Application("SERVER_NAME") + "/" + PREMAILER_HANDLER + "?url=" + Server.UrlEncode(url)
	else
		EncodeCssInlinedUrl = replace(url, "https", "http")
	end if
end function


'*******************************************************************************
'aggiunge le parti dell'url per la generazione delle email.
'*******************************************************************************
function EncodeUrlForEmail(url)
	dim Generated
	Generated = NOW()
	url = url + IIF(instr(1, url, "?", vbTextCompare)>0, "&", "?") + _
		  "HTML_FOR_EMAIL=1&GENERATED_DATETIME=" & FixLenght(Year(Generated), "0", 4) & _
													FixLenght(Month(Generated), "0", 2) & _
													FixLenght(Day(Generated), "0", 2) & _
													FixLenght(Hour(Generated), "0", 2) & _
													FixLenght(Minute(Generated), "0", 2) & _
													FixLenght(Second(Generated), "0", 2) & _
													"_" & Fix(Timer)
	EncodeUrlForEmail = url
end function		


'*******************************************************************************
'Rende i link relativi
'*******************************************************************************
Function MakeRelativeLink(HTMLstring)
	MakeRelativeLink = Replace(HTMLstring, "http://" + Application("SERVER_NAME"), "")
end function

'*******************************************************************************
'Rende i link assoluti
'*******************************************************************************
Function MakeAbsoluteLink(HTMLstring)
	dim regEx, result, matches, match, new_match
	result = HTMLstring
	Set regEx = New RegExp
	if inStr(result, "href=") > 0 then
		With regEx
			.Pattern = "href=""[^""]+"""
			.IgnoreCase = True
			.Global = True
		End With
		set matches = regEx.Execute(result)
		for each match in matches
			if inStr(match, "href=""http") = 0 AND inStr(match, "href=""mailto") = 0 AND inStr(match, "href=""callto") = 0 _
													AND inStr(match, "href=""ftp") = 0 AND inStr(match, "href=""news") = 0 then
				new_match = "href=""http://" & Application("SERVER_NAME") & Replace(match, "href=""", "")
				result = Replace(result, match, new_match)
			end if
		next
	end if
	
	if inStr(result, "src=") > 0 then
		With regEx
			.Pattern = "src=""[^""]+"""
			.IgnoreCase = True
			.Global = True
		End With
		set matches = regEx.Execute(result)
		for each match in matches
			if inStr(match, "src=""http") = 0 AND inStr(match, "src=""mailto") = 0 AND inStr(match, "src=""callto") = 0 _
												AND inStr(match, "href=""ftp") = 0 AND inStr(match, "href=""news") = 0 then
				new_match = "src=""http://" & Application("SERVER_NAME") & Replace(match, "src=""", "")
				result = Replace(result, match, new_match)
			end if
		next
	end if
	
	Set regEx = nothing
	MakeAbsoluteLink = result
end function


'*******************************************************************************
'Rimuove i tag HTML dalla stringa
'*******************************************************************************
'Giacomo
Function RemoveHtmlTags(HTMLstring, replaceString)
	dim result
	if inStr(HTMLstring, "<") > 0 then
		dim regEx
		Set regEx = New RegExp
		With regEx
			.Pattern = "<[^>]+>"
			.IgnoreCase = True
			.Global = True
		End With
		
		result = regEx.Replace(HTMLstring, replaceString)
		
		if instrRev(result, "<", -1, vbTextCompare)>0 then
			result = left(result, instrRev(result, "<", -1, vbTextCompare)-1)
		end if
		
		Set regEx = nothing
	else
		result = HTMLstring
	end if
	'result = HTMLDecode(result)
	RemoveHtmlTags = result
End Function


'Giacomo 05/11/2012
Function HTMLDecode(sText)
    Dim regEx
    Dim matches
    Dim match
    sText = Replace(sText, "&quot;", Chr(34))
    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
    sText = Replace(sText, "&nbsp;", Chr(32))
	
	sText = Replace(sText, "&lsquo;", Chr(145))
	sText = Replace(sText, "&rsquo;", Chr(146))
	'ISO 8859-1 Characters
	sText = Replace(sText, "&Agrave;", 	Chr(192)) '	capital a, grave accent
	sText = Replace(sText, "&Aacute;", 	Chr(193)) '	capital a, acute accent
	sText = Replace(sText, "&Acirc;", 	Chr(194)) '	capital a, circumflex accent
	sText = Replace(sText, "&Atilde;", 	Chr(195)) '	capital a, tilde
	sText = Replace(sText, "&Auml;", 	Chr(196)) '	capital a, umlaut mark
	sText = Replace(sText, "&Aring;", 	Chr(197)) '	capital a, ring
	sText = Replace(sText, "&AElig;", 	Chr(198)) '	capital ae
	sText = Replace(sText, "&Ccedil;", 	Chr(199)) '	capital c, cedilla
	sText = Replace(sText, "&Egrave;", 	Chr(200)) '	capital e, grave accent
	sText = Replace(sText, "&Eacute;", 	Chr(201)) '	capital e, acute accent
	sText = Replace(sText, "&Ecirc;", 	Chr(202)) '	capital e, circumflex accent
	sText = Replace(sText, "&Euml;", 	Chr(203)) '	capital e, umlaut mark
	sText = Replace(sText, "&Igrave;", 	Chr(204)) '	capital i, grave accent
	sText = Replace(sText, "&Iacute;", 	Chr(205)) '	capital i, acute accent
	sText = Replace(sText, "&Icirc;", 	Chr(206)) '	capital i, circumflex accent
	sText = Replace(sText, "&Iuml;", 	Chr(207)) '	capital i, umlaut mark
	sText = Replace(sText, "&ETH;", 	Chr(208)) '	capital eth, Icelandic
	sText = Replace(sText, "&Ntilde;", 	Chr(209)) '	capital n, tilde
	sText = Replace(sText, "&Ograve;", 	Chr(210)) '	capital o, grave accent
	sText = Replace(sText, "&Oacute;", 	Chr(211)) '	capital o, acute accent
	sText = Replace(sText, "&Ocirc;", 	Chr(212)) '	capital o, circumflex accent
	sText = Replace(sText, "&Otilde;", 	Chr(213)) '	capital o, tilde
	sText = Replace(sText, "&Ouml;", 	Chr(214)) '	capital o, umlaut mark
	sText = Replace(sText, "&Oslash;", 	Chr(216)) '	capital o, slash
	sText = Replace(sText, "&Ugrave;", 	Chr(217)) '	capital u, grave accent
	sText = Replace(sText, "&Uacute;", 	Chr(218)) '	capital u, acute accent
	sText = Replace(sText, "&Ucirc;", 	Chr(219)) '	capital u, circumflex accent
	sText = Replace(sText, "&Uuml;", 	Chr(220)) '	capital u, umlaut mark
	sText = Replace(sText, "&Yacute;", 	Chr(221)) '	capital y, acute accent
	sText = Replace(sText, "&THORN;", 	Chr(222)) '	capital THORN, Icelandic
	sText = Replace(sText, "&szlig;", 	Chr(223)) '	small sharp s, German
	sText = Replace(sText, "&agrave;", 	Chr(224)) '	small a, grave accent
	sText = Replace(sText, "&aacute;", 	Chr(225)) '	small a, acute accent
	sText = Replace(sText, "&acirc;", 	Chr(226)) '	small a, circumflex accent
	sText = Replace(sText, "&atilde;", 	Chr(227)) '	small a, tilde
	sText = Replace(sText, "&auml;", 	Chr(228)) '	small a, umlaut mark
	sText = Replace(sText, "&aring;", 	Chr(229)) '	small a, ring
	sText = Replace(sText, "&aelig;", 	Chr(230)) '	small ae
	sText = Replace(sText, "&ccedil;", 	Chr(231)) '	small c, cedilla
	sText = Replace(sText, "&egrave;", 	Chr(232)) '	small e, grave accent
	sText = Replace(sText, "&eacute;", 	Chr(233)) '	small e, acute accent
	sText = Replace(sText, "&ecirc;", 	Chr(234)) '	small e, circumflex accent
	sText = Replace(sText, "&euml;", 	Chr(235)) '	small e, umlaut mark
	sText = Replace(sText, "&igrave;", 	Chr(236)) '	small i, grave accent
	sText = Replace(sText, "&iacute;", 	Chr(237)) '	small i, acute accent
	sText = Replace(sText, "&icirc;", 	Chr(238)) '	small i, circumflex accent
	sText = Replace(sText, "&iuml;", 	Chr(239)) '	small i, umlaut mark
	sText = Replace(sText, "&eth;", 	Chr(240)) '	small eth, Icelandic
	sText = Replace(sText, "&ntilde;", 	Chr(241)) '	small n, tilde
	sText = Replace(sText, "&ograve;", 	Chr(242)) '	small o, grave accent
	sText = Replace(sText, "&oacute;", 	Chr(243)) '	small o, acute accent
	sText = Replace(sText, "&ocirc;", 	Chr(244)) '	small o, circumflex accent
	sText = Replace(sText, "&otilde;", 	Chr(245)) '	small o, tilde
	sText = Replace(sText, "&ouml;", 	Chr(246)) '	small o, umlaut mark
	sText = Replace(sText, "&oslash;", 	Chr(248)) '	small o, slash
	sText = Replace(sText, "&ugrave;", 	Chr(249)) '	small u, grave accent
	sText = Replace(sText, "&uacute;", 	Chr(250)) '	small u, acute accent
	sText = Replace(sText, "&ucirc;", 	Chr(251)) '	small u, circumflex accent
	sText = Replace(sText, "&uuml;", 	Chr(252)) '	small u, umlaut mark
	sText = Replace(sText, "&yacute;", 	Chr(253)) '	small y, acute accent
	sText = Replace(sText, "&thorn;", 	Chr(254)) '	small thorn, Icelandic
	sText = Replace(sText, "&yuml;", 	Chr(255)) '	small y, umlaut mark



    Set regEx= New RegExp

    With regEx
     .Pattern = "&#(\d+);" 'Match html unicode escapes
     .Global = True
    End With

    Set matches = regEx.Execute(sText)

    'Iterate over matches
    For Each match in matches
        'For each unicode match, replace the whole match, with the ChrW of the digits.

        sText = Replace(sText, match.Value, ChrW(match.SubMatches(0)))
    Next

    HTMLDecode = sText
End Function


Function ReplaceRexEx(StringToExtract, MatchPattern, ReplacementText)
	Dim regEx, CurrentMatch, CurrentMatches
	Set regEx = New RegExp
	regEx.Pattern = MatchPattern
	regEx.IgnoreCase = True
	regEx.Global = True
	regEx.MultiLine = True
	StringToExtract = regEx.Replace(StringToExtract, ReplacementText)
	Set regEx = Nothing
	ReplaceRexEx = StringToExtract
End Function



'*****************************************************************************************************
dim AdminErrorManager

set AdminErrorManager = new AmministrazioneErrorManager

'*****************************************************************************************************

function LastErrorDump()
	dim errorMessage
	errorMessage = "ERRORE AREA AMMINISTRATIVA" & vbCrLF & _
				   "SERVER_NAME:" & request.serverVariables("SERVER_NAME") & vbCrLF & _
				   "SCRIPT_NAME:" & request.serverVariables("SCRIPT_NAME") & vbCrLF & _
				   "REQUEST_METHOD:" & request.serverVariables("REQUEST_METHOD") & vbCrLF & _
				   "QUERY_STRING:" & request.serverVariables("QUERY_STRING") & vbCrLF
				   
	errorMessage = errorMessage & vbCrLF & _
				   "err.number=" & err.number & vbCrLF & _
				   "err.source=" & err.source & vbCrLF & _
				   "err.description=" & err.description & vbCrLF
				   
	dim errore
	set errore = server.GetLastError
	errorMessage = errorMessage & vbCrLF & _
				   "server.GetLastError.number=" & vbCrLF & _
				   "server.GetLastError.ASPCode=" & errore.ASPCode & vbCrLF & _
				   "server.GetLastError.Category=" & errore.Category & vbCrLF & _
				   "server.GetLastError.ASPDescription=" & errore.ASPDescription & vbCrLF & _
				   "server.GetLastError.Description=" & errore.Description & vbCrLF & _
				   "server.GetLastError.File=" & errore.File & vbCrLF & _
				   "server.GetLastError.Line=" & errore.Line & vbCrLF & _
				   "server.GetLastError.Column=" & errore.Column & vbCrLF & _
				   "server.GetLastError.Source=" & errore.Source & vbCrLF
	LastErrorDump = errorMessage
end function


'*****************************************************************************************************
class AmministrazioneErrorManager
	
	Private Sub Class_Terminate()
		if cIntero(err.number) <> 0 then
			%>
			<!--
			VERIFICA ERRORI AMMINISTRAZIONE ATTIVA
			<%
			
			on error resume next
			dim errorMessage
			errorMessage = LastErrorDump()
			
			response.write errorMessage
			
			if not isLocal() AND cIntero(err.number) <> 0 then
				'spedisce email con eventuale errore
				CALL SendEmailSupportEX("ERRORE area amministrativa - " & request.serverVariables("SERVER_NAME"), errorMessage)
				
				'registra nel log errore avvenuto
				dim logConn 
				set logConn = server.createobject("adodb.connection")
				logConn.open Application("DATA_ConnectionString"),"",""	
				
				CALL WriteLogAdminHttpRaw(logConn, "log_framework", 0, "ERRORE_AMMINISTRAZIONE", errorMessage, errorMessage & vbCrlf & vbcrLf & GetRawHttp())
				
				logConn.close
				set logConn = nothing
				
			end if %>
			Esecuzione completata correttamente.
			-->
		<% end if 
	End Sub
	
end class


%>
