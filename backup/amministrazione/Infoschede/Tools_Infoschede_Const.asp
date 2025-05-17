
<%
const CODICE_CAT_MODELLI = "MODELLI"
const CODICE_CAT_RICAMBI = "RICAMBI"

'costanti per i permessi per utenti area amministrativa
const POS_PERMESSO_ADMIN_INFOSCHEDE = 1
const POS_PERMESSO_CENTRO_ASSISTENZA = 2
const POS_PERMESSO_OFFICINA = 3

const MODALITA_EASY = "passa direttamente alla richiesta"
const MODALITA_NON_EASY = "richiedi conferma"
'Colori modalità guasti
const MODALITA_EASY_COLOR = "#aff7af"
const MODALITA_NON_EASY_COLOR = "#f8f890"


'ID dei profili clienti
const TRASPORTATORI = 4
const COSTRUTTORI = 1
const CLIENTI_PRIVATI = 2
const CLIENTI_PROFESSIONALI = 6
const RIVENDITORI = 3
const SUPERVISORI_NEGOZI = 5


'PERMESSI utenti area riservata
const PERMESSO_AR_CENTRO_ASSISTENZA = "USER_CENTRO_ASSISTENZA"
const PERMESSO_AR_TRASPORTATORI = ""
const PERMESSO_AR_COSTRUTTORI = "USER_COSTRUTTORE"
const PERMESSO_AR_CLIENTI_PRIVATI = "USER_CLIENTE_PRIVATO"
const PERMESSO_AR_CLIENTI_PROFESSIONALI = "USER_CLIENTE_PROFESSIONALE"
const PERMESSO_AR_RIVENDITORI = "USER_RIVENDITORE"
const PERMESSO_AR_SUPERVISORI_NEGOZI = "USER_SUPERVISORE_NEGOZI"


'Label profili clienti
set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("DATA_ConnectionString"),"",""
sql = "SELECT pro_nome_it FROM gtb_profili WHERE pro_id = "
dim LABEL_TRASPORTATORI, LABEL_COSTRUTTORI, LABEL_CLIENTI_PRIVATI, LABEL_CLIENTI_PROFESSIONALI, LABEL_RIVENDITORI, LABEL_SUPERVISORI_NEGOZI
LABEL_TRASPORTATORI = cString(GetValueList(conn, NULL, sql & TRASPORTATORI))
LABEL_COSTRUTTORI = cString(GetValueList(conn, NULL, sql & COSTRUTTORI))
LABEL_CLIENTI_PRIVATI = cString(GetValueList(conn, NULL, sql & CLIENTI_PRIVATI))
LABEL_CLIENTI_PROFESSIONALI = cString(GetValueList(conn, NULL, sql & CLIENTI_PROFESSIONALI))
LABEL_RIVENDITORI = cString(GetValueList(conn, NULL, sql & RIVENDITORI))
LABEL_SUPERVISORI_NEGOZI = cString(GetValueList(conn, NULL, sql & SUPERVISORI_NEGOZI))
sql = ""


'Colori profili anagrafiche clienti
const COLOR_CLIENTI_PRIVATI = "#f5d0f3"
const COLOR_CLIENTI_PROFESSIONALI = "#fcfabc"
const COLOR_RIVENDITORI = "#fcc8bc"
const COLOR_SUPERVISORI_NEGOZI = "#d7f5d0"


'categorie spedizioni
const DDT_CAT_ID = 1			'ddt
const LETTERE_CAT_ID = 2		'lettere d'accompagnamento
const RITIRI_CAT_ID = 3			'richieste di ritiro


'modello di default
const MODELLO_DEFAULT = 78584

const ID_PAGINA_AVVISO_NUOVA_SCHEDA = 93




%>