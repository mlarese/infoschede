<!--#INCLUDE FILE="../Tools.asp" -->
<!--#INCLUDE FILE="Tools4GoogleMap.asp" -->
<% 	dim conn, rs
	set conn = server.createobject("adodb.connection")
	conn.open Application("DATA_ConnectionString")
	
	dim currentDomain
	if instr(1,Request.ServerVariables("HTTPS"),"on",vbTextCompare) then
		currentDomain = "https://" & LCase(Application("SERVER_NAME"))
	else
		currentDomain = "http://" & LCase(Application("SERVER_NAME"))
	end if
	
	'parametri
	dim lingua
	dim key, contattoId, lat, lng, descrId
	dim iconUrl, iconWidth, iconHeight, iconAnchorX, iconAnchorY
	dim centerLatitude, centerLongitude, centerZoom
	dim localSearch, enableScroll
	
	lingua = server.HtmlEncode(request.querystring("lingua"))
	
	key = server.HtmlEncode(request.querystring("key"))
	contattoId = CIntero(request.querystring("CID"))
	lat = CRealNull(request.querystring("lat"))
	lng = CRealNull(request.querystring("lng"))
	descrId = JSFilter(request.querystring("descrId"))
	
	iconUrl = JSFilter(request.querystring("iconUrl"))
	iconWidth = CRealNull(request.querystring("iconWidth"))
	iconHeight = CRealNull(request.querystring("iconHeight"))
	iconAnchorX = CRealNull(request.querystring("iconAnchorX"))
	iconAnchorY = CRealNull(request.querystring("iconAnchorY"))
	
	centerLatitude = CRealNull(request.querystring("centerLatitude"))
	centerLongitude = CRealNull(request.querystring("centerLongitude"))
	centerZoom = CInteger(request.querystring("centerZoom"))
	
	localSearch = LCase(request.querystring("localSearch")) = "true"
	enableScroll = LCase(request.querystring("enableScroll")) <> "false"
	
	'gestione contatto
	if contattoId > 0 then
		set rs = server.createobject("adodb.recordset")
		rs.open "SELECT * FROM tb_indirizzario WHERE idElencoIndirizzi = "& cIntero(contattoId), conn, adOpenStatic, adLockReadOnly
		
		if IsNull(lat) OR IsNull(lng) then
			lat = CRealNull(rs("google_maps_latitudine"))
			lng = CRealNull(rs("google_maps_longitudine"))
		end if
	end if %>
<html>
<head>
   	<meta http-equiv="content-type" content="text/html; charset=utf-8"/>
    <title>Google Maps</title>
	<style type="text/css">
		@import url("http://www.google.com/uds/css/gsearch.css");
   		@import url("http://www.google.com/uds/solutions/localsearch/gmlocalsearch.css");
	</style>
	<% 	if GetNextWebCurrentVersion(conn, NULL) >= 5 then %>
	<link href="<%= currentDomain %>/App_Themes/Default/Stili.css" type="text/css" rel="stylesheet" />
	<% 	else %>
	<link href="<%= currentDomain %>/stili.css" type="text/css" rel="stylesheet" />
	<% 	end if %>
	<SCRIPT LANGUAGE="javascript"  src="../utils.js" type="text/javascript"></SCRIPT>
   	<script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=<%= GetGoogleMapsCode(conn, NULL) %>&amp;hl=<%= lingua %>" type="text/javascript"></script>
	<script src="http://www.google.com/uds/api?file=uds.js&amp;v=1.0&amp;hl=<%= lingua %>" type="text/javascript"></script>
    <script src="http://www.google.com/uds/solutions/localsearch/gmlocalsearch.js" type="text/javascript"></script>
</head>
<% 	if IsNull(lat) OR IsNull(lng) then %>
<body style="margin: 0px;">
	<div class="gmaps_no"><%= ChooseValueByAllLanguages(lingua, "Mappa non trovata", "Maps not found", "Karte nicht gefunden", "Carte introuvable", "Mapa que no se encuentra", "Карта не найдена", "地图未找到", "Mapa não encontrado") %></div>
<% 	else %>
<body onload="initialize()" onunload="GUnload()" style="margin: 0px;">
	<script type="text/javascript">
		var map;
		
    	function initialize() {
      		if (GBrowserIsCompatible()) {
        		map = new GMap2(document.getElementById("gmaps"));
				<% 	if enableScroll then %>
				map.enableScrollWheelZoom();
				<% 	end if %>
				map.addControl(new GLargeMapControl());
                map.addControl(new GMapTypeControl());
				
				var icona = null;
				<% 	if iconUrl <> "" AND LCase(Left(iconUrl, Len(currentDomain))) = currentDomain _
					AND iconWidth > 0 AND iconHeight > 0 then %>
				icona = new GIcon();
                icona.image = '<%= iconUrl %>';
                icona.iconSize = new GSize(<%= iconWidth %>, <%= iconHeight %>);
                icona.iconAnchor = new GPoint(<%= IIF(iconAnchorX = null, iconWidth / 2, iconAnchorX) %>, <%= IIF(iconAnchorY = null, iconHeight / 2, iconAnchorY) %>);
                icona.infoWindowAnchor = new GPoint(<%= iconWidth %>, 0);
				<% 	end if %>
				
				var descr = null;
				<% 	if contattoId > 0 AND CString(descrId) = "" then %>
				descr = document.getElementById('box');
				<%	else
						if CString(descrId) <> "" AND _
						   Instr(1, Request.ServerVariables("ALL_HTTP"), "HTTP_REFERER:"& currentDomain, vbTextCompare) > 0 then %>
				if (window.parent.location.href.toLowerCase().substring(0, <%= Len(currentDomain) %>) == "<%= currentDomain %>") {
					descr = document.getElementById('box');
					descr.innerHTML = window.parent.document.getElementById('<%= descrId %>').innerHTML;
				}
				<%		end if
					end if %>
				
				var point = new GLatLng(toNumber('<%= lat %>'), toNumber('<%= lng %>'));
				addMarker(map, icona, false, point, 0, 0, <%= IIF(request.querystring("centerZoom") = "", ZOOM_LEVEL_NAVIGATION, centerZoom) %>, "", descr)
				
				<% 	if centerLatitude <> null AND centerLongitude <> null then %>
				map.setCenter(new GLatLng(toNumber('<%= centerLatitude %>'), toNumber('<%= centerLongitude %>'), <%= IIF(request.querystring("centerZoom") <> "", centerZoom, ZOOM_LEVEL_NAVIGATION) %>);
				<% 	end if %>
				
				<% 	if localSearch then %>
				map.addControl(new google.maps.LocalSearch(), new GControlPosition(G_ANCHOR_BOTTOM_RIGHT, new GSize(10, 20)));
				<% 	end if %>
      		}
    	}
		<% 	if localSearch then %>
		GSearch.setOnLoadCallback(initialize);
		<% 	end if %>
    </script>
	<% CALL WriteJS_GoogleMaps_LocateByAddress(IIF(request.querystring("centerZoom") <> "", centerZoom, ZOOM_LEVEL_NAVIGATION), prefix, "") %>
	
	<div id="box" class="gmaps_box" style="display: none;">
	<% 	if contattoId > 0 AND CString(descrId) = "" then %>
		<h1 style="font-family: verdana; font-size: 68.75%; margin: 0px;"><%= ContactFullName(rs) %></h1>
		<h2 style="font-family: verdana; font-size: 62.5%; font-weight: normal; margin: 0px;"><%= ContactAddress(rs) %></h2>
	<% 	end if %>
	</div>
	<div id="gmaps" class="gmaps_map" style="width: 100%; height: 100%;"></div>
<% 	end if %>
</body>
</html>

<% 	if contattoId > 0 then
		rs.close
	end if
	conn.close
	
	
	'filtra paramtri javascript
	Function JSFilter(str)
		JSFilter = server.HtmlEncode(Replace(str, "'", ""))
	End Function %>
