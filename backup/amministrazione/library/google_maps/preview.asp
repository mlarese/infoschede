<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="Tools4GoogleMap.asp" -->
<% dim GoogleMapsKey
'recupera chiave di abilitazione google maps.
GoogleMapsKey = GetGoogleMapsCode(NULL, NULL) 

dim lat, lon, prefix, defaultPosition, suffix
prefix = request("prefix")
suffix = request("suffix")
defaultPosition = false
if request("lat") <> "" AND request("lon") <> "" then
	lat = request("lat")
	lon = request("lon")
else
	'default = next-aim
	defaultPosition = true
	lat = 45.47738342708687
	lon = 12.254530191421508
end if
%>
<html>
	<head>
   		<meta http-equiv="content-type" content="text/html; charset=utf-8"/>
	    <title>Google Maps Preview</title>
		<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
		<% if cString(GoogleMapsKey)<>"" then
			'goggle maps attivo
			
			if cIntero(Application("GMapVersion"))<>3  then 'GMapVersion = 2
			%>
			   	<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=<%= GoogleMapsKey %>" type="text/javascript"></script>
			    <script type="text/javascript">
			    	function initialize() {
			      		if (GBrowserIsCompatible()) {
			        		var map = new GMap2(document.getElementById("gmaps"));
							var point;
							point = new GLatLng(toNumber('<%= lat %>'), toNumber('<%= lon %>'));
							<% if not defaultPosition then %>
								var marker = addMarker(map, null, false, point, 0, 0, 10, 
													   'coordinate della selezione attuale:\nlatitudine: ' + point.lat() + '\nlongitudine: ' + point.lng(), null);
								
								GEvent.addDomListener(marker, "dblclick", function() {
									OpenAutoPositionedScrollWindow(window.parent.<%= prefix %>_google_maps_GetHref<%= suffix %>('', 'select.asp'), 'gmaps_select', 700, 500, true);
								});
							<%	end if %>
							map.setCenter(point, <%= ZOOM_LEVEL_PREVIEW %>);
			      		}
			    	}
			    </script>
				<%
				CALL WriteJS_GoogleMaps_LocateByAddress_Ex(ZOOM_LEVEL_PREVIEW, prefix, "window.parent.", suffix)
				%>
				</head>
				<body onload="initialize()" onunload="GUnload()" style="margin: 0px;">
					<div id="gmaps" style="width: 100%; height: 100%"></div>
				</body>
			<%
			else 'GMapVersion = 3
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
						var map = new google.maps.Map(document.getElementById("gmaps"), mapOptions);
						var marker = new google.maps.Marker({ position: coord,
															map: map, 
															title: 'coordinate della selezione attuale:\nlatitudine: ' + coord.lat() + '\nlongitudine: ' + coord.lng() });
					}
				</script>
				<%
				CALL WriteJS_GoogleMaps_LocateByAddress_Ex(ZOOM_LEVEL_PREVIEW, prefix, "window.parent.", suffix)
				%>
				</head>
				<body onload="initialize()" style="margin:0px;">
					<div id="gmaps" style="width:100%; height:100%"></div>
				</body>
			<%end if%>
		<% else %>
				<link rel="stylesheet" type="text/css" href="../stili.css">
			</head>
			<body>
				<div class="errore">
					Google Maps non attivo.<br>
					Contattare il supporto tecnico.
				</div>
			</body>
		<% end if %>
</html>