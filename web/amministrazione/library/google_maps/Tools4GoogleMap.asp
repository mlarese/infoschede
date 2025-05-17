<% 
'.................................................................................................
'.................................................................................................
'.................................................................................................
'COSTANTI
'.................................................................................................
'.................................................................................................
const ZOOM_LEVEL_PREVIEW = 10
const ZOOM_LEVEL_SELECTION = 10
const ZOOM_LEVEL_NAVIGATION = 10


'.................................................................................................
'.................................................................................................
'.................................................................................................
'FUNZIONI
'.................................................................................................
'.................................................................................................

'restituisce il codice di google maps relativo all'indirizzo corrente
Function GetGoogleMapsCode(conn, rs)
	dim CurrentUrl, sql, ConnCreated
	
	'controlla e crea connessione
	if not IsObjectCreated(conn) then
		set conn = server.createobject("adodb.connection")
		conn.open Application("DATA_ConnectionString")
	else
		ConnCreated = false
	end if
	
	set conn = server.createobject("adodb.connection")
	conn.open Application("DATA_ConnectionString")
	
	sql = "SELECT TOP 1 dir_google_maps_key FROM tb_webs_directories WHERE dir_url LIKE '"& ParseSql(Request.ServerVariables("HTTP_HOST"), adChar) &"'"
	GetGoogleMapsCode = GetValueList(conn, rs, sql)
	
	if ConnCreated then
		conn.close
		set conn = nothing
	end if
End Function


'procedura che scrive la parte di javascript per la localizzazione del punto sulla mappa dato l'indirizzo
sub WriteJS_GoogleMaps_LocateByAddress(zoomLevel, prefix, JsParent) 
	call WriteJS_GoogleMaps_LocateByAddress_Ex(zoomLevel, prefix, JsParent, "")
end sub

sub WriteJS_GoogleMaps_LocateByAddress_Ex(zoomLevel, prefix, JsParent, suffix)

	if cIntero(Application("GMapVersion"))<>3  then 'GMapVersion = 2
		%>
		<script type="text/javascript">
			function LocateByAddress(address){
				var geocoder;
				var markers=[];        //Will be used to temp store markers.
		
				if (GBrowserIsCompatible()) {
					var map = new GMap2(document.getElementById("gmaps<%= suffix %>"));
			
					if(geocoder==null) {
						geocoder = new GClientGeocoder();
						
						geocoder.getLatLng(address, function(newPoint){
														if (newPoint != null){
															var marker = addMarker(map, null, false, newPoint, 0, 0, 10, 
																				   'coordinate della selezione attuale:\nlatitudine: ' + newPoint.lat() + '\nlongitudine: ' + newPoint.lng(), null);
															map.setCenter(newPoint, <%= zoomLevel %>);
															<% if JsParent<>"" then %>
																<%= JsParent %><%= prefix %>_google_maps_SetCoords<%= suffix %>(newPoint.lat(), newPoint.lng())
															<% end if %>
														}
														else{
															alert('Nessun punto individuato con l\'indirizzo:"' + address + '"');
															<%= JsParent %><%= prefix %>_google_maps_RESET();
														}
													}
										   );
					 }
			
				}
			}
		</script>
	<%
	else 'GMapVersion = 3
	%>
		<script type="text/javascript">
			function LocateByAddress(address){
				var geocoder = new google.maps.Geocoder();
				geocoder.geocode( {'address': address}, function(results,status) {
					if (status == google.maps.GeocoderStatus.OK) {
						<% if JsParent<>"" then %>
							<%= JsParent %><%= prefix %>_google_maps_SetCoords<%= suffix %>(results[0].geometry.location.lat(), results[0].geometry.location.lng())
						<% else %>
							var options = {
								zoom: <%= zoomLevel %>,
								center: results[0].geometry.location,
								mapTypeId: google.maps.MapTypeId.ROADMAP
							};
							var map = new google.maps.Map(document.getElementById('gmaps<%= suffix %>'), options);
							var marker = new google.maps.Marker({position: results[0].geometry.location, map: map});
						<% end if %>
					} else {
						alert('Nessun punto individuato con l\'indirizzo:"' + address + '"');
					}
				});
			}
		</script>
	<%
	end if
end sub


%>
