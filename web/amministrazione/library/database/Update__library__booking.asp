<%
'...........................................................................................
'........................................................................................... 
'libreria di funzioni che contiene tutti gli aggiornamenti per il NEXT-booking
'...........................................................................................
'...........................................................................................


'*******************************************************************************************
'AGGIORNAMENTO BOOKING 1
'...........................................................................................
'aggiungo campi per descrivere disponibilita
'...........................................................................................
function Aggiornamento__BOOKING__1(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING__1 = "ALTER TABLE btb_disponibilita ADD "+ _
										"dis_min_stay INT NULL, "+ _
										"dis_promozione BIT NULL, "+ _
										"dis_bloccata BIT NULL;"
		case DB_SQL
			Aggiornamento__BOOKING__1 = "ALTER TABLE btb_disponibilita ADD "+ _
										"dis_min_stay INT NULL, "+ _
										"dis_promozione BIT NULL, "+ _
										"dis_bloccata BIT NULL;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING 2
'...........................................................................................
'setto i campi precedentemente aggiunti
'...........................................................................................
function Aggiornamento__BOOKING__2(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING__2 = "UPDATE btb_disponibilita SET "+ _
										"dis_min_stay = 0, "+ _
										"dis_promozione = 0, "+ _
										"dis_bloccata = 0;"
		case DB_SQL
			Aggiornamento__BOOKING__2 = "UPDATE btb_disponibilita SET "+ _
										"dis_min_stay = 0, "+ _
										"dis_promozione = 0, "+ _
										"dis_bloccata = 0;"
	end select
end function
'*******************************************************************************************


'*******************************************************************************************
'AGGIORNAMENTO BOOKING 3
'...........................................................................................
'aggiungo una chiave random alla prenotazione per accesso plugin visualizzazione dati
'ad esempio per l'oggetto che scrive l'email
'...........................................................................................
function Aggiornamento__BOOKING__3(conn)
	Select case DB_Type(conn)
		case DB_Access
			Aggiornamento__BOOKING__3 = "ALTER TABLE btb_prenotazioni ADD "+ _
										"pre_chiave TEXT(10) WITH COMPRESSION NULL;"
		case DB_SQL
			Aggiornamento__BOOKING__3 = "ALTER TABLE btb_prenotazioni ADD "+ _
										"pre_chiave nvarchar(10) NULL;"
	end select
end function
'*******************************************************************************************

%>