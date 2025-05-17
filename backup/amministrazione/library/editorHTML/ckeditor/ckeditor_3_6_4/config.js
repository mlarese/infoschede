/*
Copyright (c) 2003-2012, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
	config.startupFocus = false;

	// Define changes to default configuration here. For example:
	// config.language = 'fr';
	config.uiColor = '#cccccc';
	
	config.filebrowserBrowseUrl = urlFileManager;
	config.filebrowserImageBrowseUrl = urlFileManager;
	//config.filebrowserFlashBrowseUrl = UrlFileManager;
	//config.filebrowserUploadUrl = UrlFileManager;
	//config.filebrowserImageUploadUrl = UrlFileManager;
	//config.filebrowserFlashUploadUrl = UrlFileManager;
	
	//config.contentsCss = urlStandardCSS;
	
	// pulisce l'html in uscita
	config.entities  = false;
	config.basicEntities = false;
	config.entities_greek = false;
	config.entities_latin = false;
	
	//config.enterMode = CKEDITOR.ENTER_BR;
	config.shiftEnterMode = CKEDITOR.ENTER_BR;
	
	config.toolbarStartupExpanded = false;
	config.toolbar = 'MyToolbar';
	config.toolbar_MyToolbar =
	[
		//'Cut','Copy',
		//{ name: 'tools', items : [ 'Source','-','ShowBlocks' ] },
		//{ name: 'document', items : [ 'NewPage' ] },
		//{ name: 'editing', items : [ 'Find','Replace','-','SelectAll' ] },
		//,'Subscript','Superscript'
		//,'-','BidiLtr','BidiRtl' 
		{ name: 'clipboard', items : ['PasteText','PasteFromWord','-','Undo','Redo' ] },
		{ name: 'basicstyles', items : [ 'Bold','Italic','Underline','Strike','-','RemoveFormat' ] },
		{ name: 'colors', items : [ 'TextColor','BGColor' ] },
		{ name: 'styles', items : [ 'Font','FontSize' ] },
		{ name: 'tools', items : ['Maximize'] },
		"/",
		{ name: 'styles', items : [ 'Format'] },
		{ name: 'paragraph', items : [ 'NumberedList','BulletedList','-','Outdent','Indent','-','JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'] },
		{ name: 'links', items : [ 'Link','Unlink','Anchor' ] },
		
		{ name: 'insert', items : [ 'Image','Table','HorizontalRule','Smiley','SpecialChar','Iframe' ] },
		,
		{ name: 'tools', items : [ 'Source'] }
	];

};



CKEDITOR.on( 'dialogDefinition', function( ev )
{
	// Take the dialog name and its definition from the event data.
	var dialogName = ev.data.name;
	var dialogDefinition = ev.data.definition;

    // Check if the definition is from the dialog we're
    // interested in (the 'image' dialog). This dialog name found using DevTools plugin
    if ( dialogName == 'image')
    {
       // Remove the 'Link' and 'Advanced' tabs from the 'Image' dialog.
       dialogDefinition.removeContents( 'advanced' );
	   dialogDefinition.removeContents( 'link' );

       // Get a reference to the 'Image Info' tab.
       var infoTab = dialogDefinition.getContents( 'info' );

       // Remove unnecessary widgets/elements from the 'Image Info' tab.        
       //infoTab.remove( 'txtHSpace');
       //infoTab.remove( 'txtVSpace');
    }
	
	if ( dialogName == 'link')
    {
		// Remove the 'Link' and 'Advanced' tabs from the 'Image' dialog.
       dialogDefinition.removeContents( 'advanced' );
	}
	

//});


// Giacomo 25/10/2012 - Trovato all'indirizzo http://cksource.com/forums/viewtopic.php?t=15490
// Modifica che permette di scegliere i px di margine per i quattro lati dell'immagine
//CKEDITOR.on( 'dialogDefinition', function( ev )
//{

   // Take the dialog name and its definition from the event data.
   //var dialogName = ev.data.name;
   //var dialogDefinition = ev.data.definition;

   
   // **************************************
   // IMAGE DIALOG
   // **************************************
   if ( dialogName == 'image' )
   {
	  
	  // **************************************
	  // IMAGE INFO TAB
	  // **************************************
	  var imageInfoTab = dialogDefinition.getContents( 'info' );
	  
	  // remove the hspace and vspace fields
	  imageInfoTab.remove('txtHSpace');
	  imageInfoTab.remove('txtVSpace');
	  
	  // setup constants and other vars (recreating some functionality)
	  var IMAGE = 1,
	  LINK = 2,
	  PREVIEW = 4,
	  CLEANUP = 8,
	  regexGetSize = /^\s*(\d+)((px)|\%)?\s*$/i,
	  regexGetSizeOrEmpty = /(^\s*(\d+)((px)|\%)?\s*$)|^$/i,
	  pxLengthRegex = /^\d+px$/;
	  
	  // function to update preview
	  var updatePreview = function( dialog )
	  {
		 //Don't load before onShow.
		 if ( !dialog.originalElement || !dialog.preview )
			return 1;

		 // Read attributes and update imagePreview;
		 dialog.commitContent( PREVIEW, dialog.preview );
		 return 0;
	  };
	  
	  // function to commit changes internally
	  
	  // Avoid recursions.
	  var incommit;

	  // Synchronous field values to other impacted fields is required, e.g. border
	  // size change should alter inline-style text as well.
	  function commitInternally( targetFields )
	  {
		 if ( incommit )
			return;

		 incommit = 1;

		 var dialog = this.getDialog(),
			element = dialog.imageElement;
		 if ( element )
		 {
			// Commit this field and broadcast to target fields.
			this.commit( IMAGE, element );

			targetFields = [].concat( targetFields );
			var length = targetFields.length,
			   field;
			for ( var i = 0; i < length; i++ )
			{
			   field = dialog.getContentElement.apply( dialog, targetFields[ i ].split( ':' ) );
			   // May cause recursion.
			   field && field.setup( IMAGE, element );
			}
		 }

		 incommit = 0;
	  }
	  
	  // new margin fields
	  imageInfoTab.add( {
		 
		 type : 'fieldset',
		 label: 'Margins',
		 children :
		 [
			 {
			   type : 'vbox',
			   padding : 1,
			   width : '100px',
			   label: 'Margins',
			   align: 'center',
			   children :
			   [
			   
				  // margin-top
				  {
					 type : 'text',
					 id : 'txtMarginTop',
					 width: '40px',
					 labelLayout : 'horizontal',
					 label : 'Top',
					 'default' : '',
					 onKeyUp : function()
					 {
						updatePreview( this.getDialog() );
					 },
					 onChange : function()
					 {
						commitInternally.call( this, 'advanced:txtdlgGenStyle' );
					 },
					 validate : CKEDITOR.dialog.validate.integer( ev.editor.lang.image.validateVSpace ),
					 setup : function( type, element )
					 {
						if ( type == IMAGE )
						{
						   var value,
							  marginTopPx,
							  marginTopStyle = element.getStyle( 'margin-top' );
			   
						   marginTopStyle = marginTopStyle && marginTopStyle.match( pxLengthRegex );
						   marginTopPx = parseInt( marginTopStyle, 10 );
						   value = marginTopPx;
						   isNaN( parseInt( value, 10 ) ) && ( value = element.getAttribute( 'vspace' ) );
			   
						   this.setValue( value );
						}
					 },
					 commit : function( type, element, internalCommit )
					 {
						var value = parseInt( this.getValue(), 10 );
						if ( type == IMAGE || type == PREVIEW )
						{
						   if ( !isNaN( value ) )
						   {
							  element.setStyle( 'margin-top', CKEDITOR.tools.cssLength( value ) );
						   }
						   else if ( !value && this.isChanged( ) )
						   {
							  element.removeStyle( 'margin-top' );
						   }
			   
						   if ( !internalCommit && type == IMAGE )
							  element.removeAttribute( 'vspace' );
						}
						else if ( type == CLEANUP )
						{
						   element.removeAttribute( 'vspace' );
						   element.removeStyle( 'margin-top' );
						}
					 }
				  }, // end margin-top
				  
				  // margin-right
				  {
					 type : 'text',
					 id : 'txtMarginRight',
					 width: '40px',
					 labelLayout : 'horizontal',
					 label : 'Right',
					 'default' : '',
					 onKeyUp : function()
					 {
						updatePreview( this.getDialog() );
					 },
					 onChange : function()
					 {
						commitInternally.call( this, 'advanced:txtdlgGenStyle' );
					 },
					 validate : CKEDITOR.dialog.validate.integer( ev.editor.lang.image.validateHSpace ),
					 setup : function( type, element )
					 {
						if ( type == IMAGE )
						{
						   var value,
							  marginRightPx,
							  marginRightStyle = element.getStyle( 'margin-right' );
			   
						   marginRightStyle = marginRightStyle && marginRightStyle.match( pxLengthRegex );
						   marginRightPx = parseInt( marginRightStyle, 10 );
						   value = marginRightPx;
						   isNaN( parseInt( value, 10 ) ) && ( value = element.getAttribute( 'hspace' ) );
			   
						   this.setValue( value );
						}
					 },
					 commit : function( type, element, internalCommit )
					 {
						var value = parseInt( this.getValue(), 10 );
						if ( type == IMAGE || type == PREVIEW )
						{
						   if ( !isNaN( value ) )
						   {
							  element.setStyle( 'margin-right', CKEDITOR.tools.cssLength( value ) );
						   }
						   else if ( !value && this.isChanged( ) )
						   {
							  element.removeStyle( 'margin-right' );
						   }
			   
						   if ( !internalCommit && type == IMAGE )
							  element.removeAttribute( 'hspace' );
						}
						else if ( type == CLEANUP )
						{
						   element.removeAttribute( 'hspace' );
						   element.removeStyle( 'margin-right' );
						}
					 }
				  }, // end margin-right
				  
				  // margin-bottom
				  {
					 type : 'text',
					 id : 'txtMarginBottom',
					 width: '40px',
					 labelLayout : 'horizontal',
					 label : 'Bottom',
					 'default' : '',
					 onKeyUp : function()
					 {
						updatePreview( this.getDialog() );
					 },
					 onChange : function()
					 {
						commitInternally.call( this, 'advanced:txtdlgGenStyle' );
					 },
					 validate : CKEDITOR.dialog.validate.integer( ev.editor.lang.image.validateVSpace ),
					 setup : function( type, element )
					 {
						if ( type == IMAGE )
						{
						   var value,
							  marginBottomPx,
							  marginBottomStyle = element.getStyle( 'margin-bottom' );
			   
						   marginBottomStyle = marginBottomStyle && marginBottomStyle.match( pxLengthRegex );
						   marginBottomPx = parseInt( marginBottomStyle, 10 );
						   value = marginBottomPx;
						   isNaN( parseInt( value, 10 ) ) && ( value = element.getAttribute( 'vspace' ) );
			   
						   this.setValue( value );
						}
					 },
					 commit : function( type, element, internalCommit )
					 {
						var value = parseInt( this.getValue(), 10 );
						if ( type == IMAGE || type == PREVIEW )
						{
						   if ( !isNaN( value ) )
						   {
							  element.setStyle( 'margin-bottom', CKEDITOR.tools.cssLength( value ) );
						   }
						   else if ( !value && this.isChanged( ) )
						   {
							  element.removeStyle( 'margin-bottom' );
						   }
			   
						   if ( !internalCommit && type == IMAGE )
							  element.removeAttribute( 'vspace' );
						}
						else if ( type == CLEANUP )
						{
						   element.removeAttribute( 'vspace' );
						   element.removeStyle( 'margin-bottom' );
						}
					 }
				  }, // end margin-bottom
				  
				  // margin-left
				  {
					 type : 'text',
					 id : 'txtMarginLeft',
					 width: '40px',
					 labelLayout : 'horizontal',
					 label : 'Left',
					 'default' : '',
					 onKeyUp : function()
					 {
						updatePreview( this.getDialog() );
					 },
					 onChange : function()
					 {
						commitInternally.call( this, 'advanced:txtdlgGenStyle' );
					 },
					 validate : CKEDITOR.dialog.validate.integer( ev.editor.lang.image.validateHSpace ),
					 setup : function( type, element )
					 {
						if ( type == IMAGE )
						{
						   var value,
							  marginLeftPx,
							  marginLeftStyle = element.getStyle( 'margin-left' );
			   
						   marginLeftStyle = marginLeftStyle && marginLeftStyle.match( pxLengthRegex );
						   marginLeftPx = parseInt( marginLeftStyle, 10 );
						   value = marginLeftPx;
						   isNaN( parseInt( value, 10 ) ) && ( value = element.getAttribute( 'hspace' ) );
			   
						   this.setValue( value );
						}
					 },
					 commit : function( type, element, internalCommit )
					 {
						var value = parseInt( this.getValue(), 10 );
						if ( type == IMAGE || type == PREVIEW )
						{
						   if ( !isNaN( value ) )
						   {
							  element.setStyle( 'margin-left', CKEDITOR.tools.cssLength( value ) );
						   }
						   else if ( !value && this.isChanged( ) )
						   {
							  element.removeStyle( 'margin-left' );
						   }
			   
						   if ( !internalCommit && type == IMAGE )
							  element.removeAttribute( 'hspace' );
						}
						else if ( type == CLEANUP )
						{
						   element.removeAttribute( 'hspace' );
						   element.removeStyle( 'margin-left' );
						}
					 }
				  } // end margin-left
			   ]
			}     
					
		 ]
			
	  }, 'txtBorder');

	  // this syntax chokes in Safari and others (I think "default" is reserved)
	  //imageInfoTab.get('txtBorder').default = '0';
	  
	  // this syntax works...
	  // set default border to zero
	  var imageTxtBorder = imageInfoTab.get('txtBorder');
	  imageTxtBorder['default'] = '0';
	  
   }
   
});


//http://theholyjava.wordpress.com/2011/04/01/ckeditor-collapsing-only-2nd-toolbar-rows-howto/
//Permette una gestione personalizzata della minimizzazione della toolbox
CKEDITOR.on('instanceReady', function(e) {

    function switchVisibilityAfter1stRow(toolbox, show)
    {
        var inFirstRow = true;
        var elements = toolbox.getChildren();
        var elementsCount = elements.count();
        var elementIndex = 0;
        var element = elements.getItem(elementIndex);
        for (; elementIndex < elementsCount; element = elements.getItem(++elementIndex))
        {
            inFirstRow = inFirstRow && !(element.is('div') && element.hasClass('cke_break'));

            if (!inFirstRow)
            {
                if (show) element.show(); else element.hide();
            }
        }
    }

    var editor = e.editor;
    var collapser = (function()
    {
        try
        {
            // We've HTML: td.cke_top {
            //  div.cke_toolbox {span.cke_toolbar, ... }
            //  , a.cke_toolbox_collapser }
            var firstToolbarId = editor.toolbox.toolbars[0].id;
            var firstToolbar = CKEDITOR.document.getById(firstToolbarId);
            var toolbox = firstToolbar.getParent();
            var collapser = toolbox.getNext();
            return collapser;
        }
        catch (e) {}
    })();

    // Copied from editor/_source/plugins/toolbar/plugin.js & modified
    editor.addCommand( 'toolbarCollapse',
    {

        exec : function( editor )
        {
            if (collapser == null) return;

            var toolbox = collapser.getPrevious(),
            contents = editor.getThemeSpace( 'contents' ),
            toolboxContainer = toolbox.getParent(),
            contentHeight = parseInt( contents.$.style.height, 10 ),
            previousHeight = toolboxContainer.$.offsetHeight,

            collapsed = toolbox.hasClass('iterate_tbx_hidden');//!toolbox.isVisible();

            if ( !collapsed )
            {
                switchVisibilityAfter1stRow(toolbox, false);    // toolbox.hide();
                toolbox.addClass('iterate_tbx_hidden');
                if (!toolbox.isVisible()) toolbox.show(); // necessary 1st time if initially collapsed

                collapser.addClass( 'cke_toolbox_collapser_min' );
                collapser.setAttribute( 'title', editor.lang.toolbarExpand );
            }
            else
            {
                switchVisibilityAfter1stRow(toolbox, true);    // toolbox.show();
                toolbox.removeClass('iterate_tbx_hidden');

                collapser.removeClass( 'cke_toolbox_collapser_min' );
                collapser.setAttribute( 'title', editor.lang.toolbarCollapse );
            }

            // Update collapser symbol.
            collapser.getFirst().setText( collapsed ?
                '\u25B2' :		// BLACK UP-POINTING TRIANGLE
                '\u25C0' );	// BLACK LEFT-POINTING TRIANGLE

            var dy = toolboxContainer.$.offsetHeight - previousHeight;
            contents.setStyle( 'height', ( contentHeight - dy ) + 'px' );

            editor.fire( 'resize' );
        },

        modes : {
            wysiwyg : 1,
            source : 1
        }
    } )

    // Make sure advanced toolbars initially collapsed
    editor.execCommand( 'toolbarCollapse' );
});


