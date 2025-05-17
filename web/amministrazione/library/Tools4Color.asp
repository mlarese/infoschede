<% 
'.................................................................................................
'.................................................................................................
'COSTANTI
'.................................................................................................
'.................................................................................................
'definizione set di caratteri validi
const COLOR_CHARSET 		= "#0123456789ABCDEF"

'.................................................................................................
'.................................................................................................
'FUNZIONI E PROCEDURE
'.................................................................................................
'.................................................................................................


Sub WriteColorPicker_Input(form_name, field_name, field_value, field_style, show_reset, obbligatorio, onchange_JS_event)
	CALL WriteColorPicker_Input_Disable(form_name, field_name, field_value, field_style, show_reset, obbligatorio, onchange_JS_event, false)
End Sub

'.................................................................................................
'.. 	Creazione di un imput con funzione di scelta del colore in una nuova finestra
'.................................................................................................
sub WriteColorPicker_Input_Disable(form_name, field_name, field_value, field_style, show_reset, obbligatorio, onchange_JS_event, disabled)
	%>
	<script language="JavaScript" type="text/javascript">
		function <%= field_name %>_select(sender){
			var url = "<%= GetLibraryPath() %>PickerColor.asp?field_value=" + escape(document.<%= form_name %>.<%= field_name %>.value) + "&field_name=<%= field_name %>&form_name=<%= form_name %>";
			OpenAutoPositionedScrollWindow(url, sender.name, 400, 400, false);
		}
		
		function <%= field_name %>_change(sender){
			var color_preview = document.getElementById("<%= field_name %>_preview");
			color_preview.style.backgroundColor = sender.value;
			
			<% if onchange_JS_event<>"" then %>
				<%= onchange_JS_event %>
			<% end if %>
		}
		
		function <%= field_name %>_reset(sender){
			sender.value = "transparent";
			<%= field_name %>_change(sender);
		}
	</script>
	<table cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<span id="<%= field_name %>_preview" style="display:block; font-size:11px; border:1px solid #999; border-right:0px; width:14px; height:14px; background-color:<%= field_value %>;"
				<% 	if NOT disabled then %>
					 style="cursor:pointer;"
					 onclick="<%= field_name %>_select(<%= form_name %>.<%= field_name %>)" title="click per aprire la finestra di selezione del colore."
				<% 	end if %>>
					&nbsp;</span>
			</td>
			<td>
				<input type="text" maxlength="11" name="<%= field_name %>" id="<%= field_name %>_input" value="<%= field_value %>" 
				<% 	if disabled then %>
					   disabled
				<% 	else %>
					   ondblclick="<%= field_name %>_select(this)" onchange="<%= field_name %>_change(this)"
				<% 	end if %>
					   title="click per modificare manualmente il colore<%= vbCrLf %> doppio click per aprire la finestra di selezione del colore"
					   size="11" style="letter-spacing:1px; padding-left:3px;<%= field_style %>">
			</td>
			<% 	if NOT disabled then %>
			<td>
				<a href="javascript:void(0);" title="click per aprire la finestra di selezione del colore." <%= ACTIVE_STATUS %> 
				   class="button_input"  id="link_scegli_<%= field_name %>" onclick="<%= field_name %>_select(<%= form_name %>.<%= field_name %>)">
					SELEZIONA
				</a>
			</td>
			<%if show_reset then %>
				<td>
					<a href="javascript:void(0);" id="link_reset_<%= field_name %>" class="button_input" onclick="<%= field_name %>_reset(<%= form_name %>.<%= field_name %>)" title="cancella la selezione" <%= ACTIVE_STATUS %>>
						RESET
					</a>
				</td>
			<%end if
			end if
			if obbligatorio then %>
				<td>&nbsp;(*)</td>
			<% end if %>
		</tr>
	</table>
	<%
end sub


'.................................................................................................
'.. 	Funzione che verifica la validita' di un colore
'.................................................................................................
function IsColor(Color)
	Color = Ucase(cString(Color))
	
	if Color <> "TRANSPARENT" then
	
		'lunghezza massima colore: 6 + # (es: #FFFFFF) o 4 + # (es: #FFF)
		IsColor = ( (Len(Color)=7) OR (Len(Color)=4) )
			if not IsColor then Exit function
			
		'controlla se i caratteri sono corretti
		IsColor = CheckChar(Color, COLOR_CHARSET)
			if not IsColor then Exit function
	
		'controlla presenza di un solo #
		IsColor = ( (len(Color)-1) = len(replace(Color, "#", "")) )
	else
		IsColor = true
	end if
end function


'.................................................................................................
'.. 	procedura che scrive un piccolo riquadro di visualizzazione del colore
'.................................................................................................
sub WriteColor(color)
	CALL WriteColoreTipo(color, color)
end sub


'.................................................................................................
'.. 	procedura che scrive un piccolo riquadro di visualizzazione del colore
'.................................................................................................
sub WriteColoreTipo(color, label)
	if IsColor(color) then %>
		<span class="icona" style="background-color:<%= color %>;" title="<%=label%>">&nbsp;</span>
	<%end if
end sub
%>