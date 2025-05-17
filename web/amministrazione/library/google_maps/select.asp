<%@ Language=VBScript CODEPAGE=65001%>
<% Option Explicit %>
<% response.charset = "UTF-8" %>
<!--#INCLUDE FILE="Tools4GoogleMap.asp" -->
<%
'--------------------------------------------------------
sezione_testata = ChooseValueByAllLanguages(Session("LINGUA"), "selezione punto sulla mappa", "select the point on the map", "", "", "", "", "", "") %>
<!--#INCLUDE FILE="../Intestazione_Ridotta.asp" -->
<%'----------------------------------------------------- 
dim prefix, GoogleMapsKey, suffix
prefix = request("prefix")
suffix = request("suffix")

dim lat, lon
if request("lat") <> "" AND request("lon") <> "" then
	lat = request("lat")
	lon = request("lon")
else
	'default = veneto - next-aim
	lat = 45.47738342708687
	lon = 12.254530191421508
end if
					
'recupera chiave di abilitazione google maps.
GoogleMapsKey = GetGoogleMapsCode(NULL, NULL)

if cString(GoogleMapsKey)<>"" then 
	
	if cIntero(Application("GMapVersion"))<>3  then 'GMapVersion = 2
		%>
		<style type="text/css">
			@import url("http://www.google.com/uds/css/gsearch.css");
			@import url("http://www.google.com/uds/solutions/localsearch/gmlocalsearch.css");
		</style>
		
		<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=<%= GoogleMapsKey %>" type="text/javascript"></script>
		<script src="http://www.google.com/uds/api?file=uds.js&amp;v=1.0" type="text/javascript"></script>
		<script src="http://www.google.com/uds/solutions/localsearch/gmlocalsearch.js" type="text/javascript"></script>    
		<script type="text/javascript">
			var map;

			function addMakerDraggable(map, point) {
				var marker = addMarker(map, null, true, point, 0, 0, 10, '', null);
				
				GEvent.addDomListener(marker, "dblclick", function() {
					var p = marker.getPoint();
					opener.<%= prefix %>_google_maps_SetCoords<%= suffix %>(p.lat(), p.lng())
					window.close();
				});
			}

			document.body.onload = function() {
				if (GBrowserIsCompatible()) {
					map = new GMap2(document.getElementById("gmaps<%= suffix %>"));
					map.enableScrollWheelZoom();
					map.addControl(new GLargeMapControl());
					map.addControl(new GMapTypeControl());
					var point;
					point = new GLatLng(toNumber('<%= lat %>'), toNumber('<%= lon %>'));
					addMakerDraggable(map, point)
					map.setCenter(point, 10);
					
					map.addControl(new google.maps.LocalSearch({ onMarkersSetCallback : searchMarkersSet }), new GControlPosition(G_ANCHOR_BOTTOM_RIGHT, new GSize(10, 20)));
				}
			}
			
			document.body.onunload = function(){GUnload()};
			
			function searchMarkersSet(markers) {
				for (var i = 0; i < markers.length; i++) {
					addMakerDraggable(map, markers[i].marker.getPoint());
				}
			}
		</script>
		<%
	else
		%>
		<script type="text/javascript" src="https://maps.googleapis.com/maps/api/js?key=<%= GoogleMapsKey %>&sensor=false"></script>
		<script type="text/javascript">
			function initialize() {
	
				var coord = new google.maps.LatLng(toNumber('<%= lat %>'), toNumber('<%= lon %>'));
				var mapOptions = {
					center: coord,
					zoom: toNumber('<%= ZOOM_LEVEL_PREVIEW %>'),
					mapTypeId: google.maps.MapTypeId.ROADMAP        
				};
				var map = new google.maps.Map(document.getElementById("gmaps<%= suffix %>"), mapOptions);
				var marker = new google.maps.Marker({ position: coord,
													  map: map,
													  clickable : true,
													  draggable : true,
													  title: 'Trascina questo segnaposto nel punto desiderato,\n infine fai doppio click su di esso.' });
				
				google.maps.event.addListener(marker, "dblclick", function (e) {
																	var p = marker.getPosition();
																	opener.<%= prefix %>_google_maps_SetCoords<%= suffix %>(p.lat(), p.lng())
																	window.close();
																});
			}
			
			window.onload = initialize;
		</script>
		<%
	end if
end if %>
<div id="content_ridotto">
	<table cellspacing="1" cellpadding="0" class="tabella_madre" style="margin-bottom:5px;">
		<caption class="border"><%= ChooseValueByAllLanguages(Session("LINGUA"), "Seleziona il punto sulla mappa con ", "Select the point on the map with ", "", "", "", "", "", "")%>
			<a href="http://mappe.google.it" target="_blank" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Apre Google Maps in una nuova finestra", "Open Google Maps in a new window", "", "", "", "", "", "")%>">Google Maps</a></caption>
		<tr>
			<td>
				<% if cString(GoogleMapsKey)<>"" then %>
					<div id="gmaps<%= suffix %>" style="width: 100%; height: 500px"></div>
				<% else %>
					<div class="errore">
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Google Maps non attivo.", "Google Maps offline.", "", "", "", "", "", "")%><br>
						<%= ChooseValueByAllLanguages(Session("LINGUA"), "Contattare il supporto tecnico.", "Please contact technical support.", "", "", "", "", "", "")%>
					</div>
				<% end if %>
			</td>
		</tr>
		<tr>
			<td class="note">
				<%= ChooseValueByAllLanguages(Session("LINGUA"), "Per selezionare il punto sulla mappa trascinare il segnaposto e, quando lo si &egrave; posizionato correttamente, fare doppio click su di esso.", "Drag the marker to select the point on the map, and when it is positioned correctly, double click on it.", "", "", "", "", "", "")%>
			</td>
		</tr>
		<tr>
			<td colspan="4" class="footer">
				<a class="button" href="javascript:window.close();" title="<%= ChooseValueByAllLanguages(Session("LINGUA"), "Chiudi la finestra", "Close the window", "", "", "", "", "", "")%>" <%= ACTIVE_STATUS %>>
					<%= ChooseValueByAllLanguages(Session("LINGUA"), "CHIUDI", "CLOSE", "", "", "", "", "", "")%>
				</a>
			</td>
		</tr>
	</table>
</div>
</body>
</html>
<script language="JavaScript" type="text/javascript">
	FitWindowSize(this);
</script>