
<% 
dim CKeditorIncludeWrited
CKeditorIncludeWrited = false

sub WriteCKeditorInclude() 
	if not CKeditorIncludeWrited then%>
		<script type="text/javascript">
			//setto il focus nel primo input della pagina
			var inputs;
			inputs = document.getElementsByTagName('input');
			if (inputs[0] != null) {
				setTimeout(function(){
					inputs[0].focus();
					window.scrollTo(0, 0);
				}, 2000);
			}

			var urlFileManager;
			var urlStandardCSS;
			urlFileManager = '<%=GetLibraryPath()%>filemanager.asp?STANDALONE=1&OBJECT_TYPE=<%=FILE_SYSTEM_FILE%>'+
							'&FILEMAN_AZ_ID=<%=Application("AZ_ID")%>&ABS_PATH=true&filter=images&selected=&F=\\images/&field_id=';
							
			urlStandardCSS = '<%=GetSiteUrl(null, null, null) & "/upload/" & Application("AZ_ID") & "/css/stili_testo.css"%>'

		</script>
		<script type="text/javascript" src="<%=GetLibraryPath()%>editorHTML/ckeditor/ckeditor_3_6_4/ckeditor.js"></script>
		<% CKeditorIncludeWrited = true
	end if
end sub

sub activateCKEditor(nameTextArea)
	CALL activateCKEditorComplete(nameTextArea, "")
end sub	
	
sub activateCKEditorComplete(nameTextArea, height)
	CALL  WriteCKeditorInclude()
	%>
	<script type="text/javascript">

		var editor_<%=nameTextArea%> = CKEDITOR.replace( '<%=nameTextArea%>',
														{	
															startupFocus : false,
															
															<% if cString(height) <> "" then %>
																height: '<%=height%>',
															<% end if %>
															
															on :
															{
																instanceReady : function( ev )
																{
																	// Output paragraphs as <p>Text</p>.
																	this.dataProcessor.writer.setRules( 'p',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'li',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'ol',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'ul',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'div',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'h1',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'h2',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'h3',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'h4',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'h5',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'h6',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'address',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'pre',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'table',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'caption',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'tbody',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'thead',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'th',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'td',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'tr',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																	this.dataProcessor.writer.setRules( 'hr',
																		{
																			indent : false,
																			breakBeforeOpen : false,
																			breakAfterOpen : false,
																			breakBeforeClose : false,
																			breakAfterClose : false
																		});
																}
															}
														});
	</script>
	<%
end sub

sub importStiliStandardCSS(nameTextArea)
	%>
	<script type="text/javascript">
		editor_<%=nameTextArea%>.config.contentsCss = urlStandardCSS;
	</script>
	<%
end sub



%>