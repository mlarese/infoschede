/*
Copyright (c) 2003-2012, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license


(function(){var a=function(b,c){var d=1,e=2,f=4,g=8,h=/^\s*(\d+)((px)|\%)?\s*$/i,i=/(^\s*(\d+)((px)|\%)?\s*$)|^$/i,j=/^\d+px$/,k=function(){var B=this.getValue(),C=this.getDialog(),D=B.match(h);if(D){if(D[2]=='%')p(C,false);B=D[1];}if(C.lockRatio){var E=C.originalElement;if(E.getCustomData('isReady')=='true')if(this.id=='txtHeight'){if(B&&B!='0')B=Math.round(E.$.width*(B/E.$.height));if(!isNaN(B))C.setValueOf('info','txtWidth',B);}else{if(B&&B!='0')B=Math.round(E.$.height*(B/E.$.width));if(!isNaN(B))C.setValueOf('info','txtHeight',B);}}l(C);},l=function(B){if(!B.originalElement||!B.preview)return 1;B.commitContent(f,B.preview);return 0;};function m(){var B=arguments,C=this.getContentElement('advanced','txtdlgGenStyle');C&&C.commit.apply(C,B);this.foreach(function(D){if(D.commit&&D.id!='txtdlgGenStyle')D.commit.apply(D,B);});};var n;function o(B){if(n)return;n=1;var C=this.getDialog(),D=C.imageElement;if(D){this.commit(d,D);B=[].concat(B);var E=B.length,F;for(var G=0;G<E;G++){F=C.getContentElement.apply(C,B[G].split(':'));F&&F.setup(d,D);}}n=0;};var p=function(B,C){if(!B.getContentElement('info','ratioLock'))return null;var D=B.originalElement;if(!D)return null;if(C=='check'){if(!B.userlockRatio&&D.getCustomData('isReady')=='true'){var E=B.getValueOf('info','txtWidth'),F=B.getValueOf('info','txtHeight'),G=D.$.width*1000/D.$.height,H=E*1000/F;B.lockRatio=false;if(!E&&!F)B.lockRatio=true;else if(!isNaN(G)&&!isNaN(H))if(Math.round(G)==Math.round(H))B.lockRatio=true;}}else if(C!=undefined)B.lockRatio=C;else{B.userlockRatio=1;B.lockRatio=!B.lockRatio;}var I=CKEDITOR.document.getById(w);if(B.lockRatio)I.removeClass('cke_btn_unlocked');else I.addClass('cke_btn_unlocked');I.setAttribute('aria-checked',B.lockRatio);if(CKEDITOR.env.hc){var J=I.getChild(0);J.setHtml(B.lockRatio?CKEDITOR.env.ie?'■':'▣':CKEDITOR.env.ie?'□':'▢');}return B.lockRatio;},q=function(B){var C=B.originalElement;if(C.getCustomData('isReady')=='true'){var D=B.getContentElement('info','txtWidth'),E=B.getContentElement('info','txtHeight');D&&D.setValue(C.$.width);E&&E.setValue(C.$.height);}l(B);},r=function(B,C){if(B!=d)return;function D(I,J){var K=I.match(h);if(K){if(K[2]=='%'){K[1]+='%';p(E,false);}return K[1];}return J;};var E=this.getDialog(),F='',G=this.id=='txtWidth'?'width':'height',H=C.getAttribute(G);if(H)F=D(H,F);F=D(C.getStyle(G),F);this.setValue(F);},s,t=function(){var B=this.originalElement;B.setCustomData('isReady','true');B.removeListener('load',t);B.removeListener('error',u);B.removeListener('abort',u);
CKEDITOR.document.getById(y).setStyle('display','none');if(!this.dontResetSize)q(this);if(this.firstLoad)CKEDITOR.tools.setTimeout(function(){p(this,'check');},0,this);this.firstLoad=false;this.dontResetSize=false;},u=function(){var D=this;var B=D.originalElement;B.removeListener('load',t);B.removeListener('error',u);B.removeListener('abort',u);var C=CKEDITOR.getUrl(b.skinPath+'images/noimage.png');if(D.preview)D.preview.setAttribute('src',C);CKEDITOR.document.getById(y).setStyle('display','none');p(D,false);},v=function(B){return CKEDITOR.tools.getNextId()+'_'+B;},w=v('btnLockSizes'),x=v('btnResetSize'),y=v('ImagePreviewLoader'),z=v('previewLink'),A=v('previewImage');return{title:b.lang.image[c=='image'?'title':'titleButton'],minWidth:420,minHeight:360,onShow:function(){var H=this;H.imageElement=false;H.linkElement=false;H.imageEditMode=false;H.linkEditMode=false;H.lockRatio=true;H.userlockRatio=0;H.dontResetSize=false;H.firstLoad=true;H.addLink=false;var B=H.getParentEditor(),C=B.getSelection(),D=C&&C.getSelectedElement(),E=D&&D.getAscendant('a');CKEDITOR.document.getById(y).setStyle('display','none');s=new CKEDITOR.dom.element('img',B.document);H.preview=CKEDITOR.document.getById(A);H.originalElement=B.document.createElement('img');H.originalElement.setAttribute('alt','');H.originalElement.setCustomData('isReady','false');if(E){H.linkElement=E;H.linkEditMode=true;var F=E.getChildren();if(F.count()==1){var G=F.getItem(0).getName();if(G=='img'||G=='input'){H.imageElement=F.getItem(0);if(H.imageElement.getName()=='img')H.imageEditMode='img';else if(H.imageElement.getName()=='input')H.imageEditMode='input';}}if(c=='image')H.setupContent(e,E);}if(D&&D.getName()=='img'&&!D.data('cke-realelement')||D&&D.getName()=='input'&&D.getAttribute('type')=='image'){H.imageEditMode=D.getName();H.imageElement=D;}if(H.imageEditMode){H.cleanImageElement=H.imageElement;H.imageElement=H.cleanImageElement.clone(true,true);H.setupContent(d,H.imageElement);}else H.imageElement=B.document.createElement('img');p(H,true);if(!CKEDITOR.tools.trim(H.getValueOf('info','txtUrl'))){H.preview.removeAttribute('src');H.preview.setStyle('display','none');}},onOk:function(){var C=this;if(C.imageEditMode){var B=C.imageEditMode;if(c=='image'&&B=='input'&&confirm(b.lang.image.button2Img)){B='img';C.imageElement=b.document.createElement('img');C.imageElement.setAttribute('alt','');b.insertElement(C.imageElement);}else if(c!='image'&&B=='img'&&confirm(b.lang.image.img2Button)){B='input';C.imageElement=b.document.createElement('input');
C.imageElement.setAttributes({type:'image',alt:''});b.insertElement(C.imageElement);}else{C.imageElement=C.cleanImageElement;delete C.cleanImageElement;}}else{if(c=='image')C.imageElement=b.document.createElement('img');else{C.imageElement=b.document.createElement('input');C.imageElement.setAttribute('type','image');}C.imageElement.setAttribute('alt','');}if(!C.linkEditMode)C.linkElement=b.document.createElement('a');C.commitContent(d,C.imageElement);C.commitContent(e,C.linkElement);if(!C.imageElement.getAttribute('style'))C.imageElement.removeAttribute('style');if(!C.imageEditMode){if(C.addLink){if(!C.linkEditMode){b.insertElement(C.linkElement);C.linkElement.append(C.imageElement,false);}else b.insertElement(C.imageElement);}else b.insertElement(C.imageElement);}else if(!C.linkEditMode&&C.addLink){b.insertElement(C.linkElement);C.imageElement.appendTo(C.linkElement);}else if(C.linkEditMode&&!C.addLink){b.getSelection().selectElement(C.linkElement);b.insertElement(C.imageElement);}},onLoad:function(){var C=this;if(c!='image')C.hidePage('Link');var B=C._.element.getDocument();if(C.getContentElement('info','ratioLock')){C.addFocusable(B.getById(x),5);C.addFocusable(B.getById(w),5);}C.commitContent=m;},onHide:function(){var B=this;if(B.preview)B.commitContent(g,B.preview);if(B.originalElement){B.originalElement.removeListener('load',t);B.originalElement.removeListener('error',u);B.originalElement.removeListener('abort',u);B.originalElement.remove();B.originalElement=false;}delete B.imageElement;},contents:[{id:'info',label:b.lang.image.infoTab,accessKey:'I',elements:[{type:'vbox',padding:0,children:[{type:'hbox',widths:['280px','110px'],align:'right',children:[{id:'txtUrl',type:'text',label:b.lang.common.url,required:true,onChange:function(){var B=this.getDialog(),C=this.getValue();if(C.length>0){B=this.getDialog();var D=B.originalElement;B.preview.removeStyle('display');D.setCustomData('isReady','false');var E=CKEDITOR.document.getById(y);if(E)E.setStyle('display','');D.on('load',t,B);D.on('error',u,B);D.on('abort',u,B);D.setAttribute('src',C);s.setAttribute('src',C);B.preview.setAttribute('src',s.$.src);l(B);}else if(B.preview){B.preview.removeAttribute('src');B.preview.setStyle('display','none');}},setup:function(B,C){if(B==d){var D=C.data('cke-saved-src')||C.getAttribute('src'),E=this;this.getDialog().dontResetSize=true;E.setValue(D);E.setInitValue();}},commit:function(B,C){var D=this;if(B==d&&(D.getValue()||D.isChanged())){C.data('cke-saved-src',D.getValue());
C.setAttribute('src',D.getValue());}else if(B==g){C.setAttribute('src','');C.removeAttribute('src');}},validate:CKEDITOR.dialog.validate.notEmpty(b.lang.image.urlMissing)},{type:'button',id:'browse',style:'display:inline-block;margin-top:10px;',align:'center',label:b.lang.common.browseServer,hidden:true,filebrowser:'info:txtUrl'}]}]},{id:'txtAlt',type:'text',label:b.lang.image.alt,accessKey:'T','default':'',onChange:function(){l(this.getDialog());},setup:function(B,C){if(B==d)this.setValue(C.getAttribute('alt'));},commit:function(B,C){var D=this;if(B==d){if(D.getValue()||D.isChanged())C.setAttribute('alt',D.getValue());}else if(B==f)C.setAttribute('alt',D.getValue());else if(B==g)C.removeAttribute('alt');}},{type:'hbox',children:[{id:'basic',type:'vbox',children:[{type:'hbox',widths:['50%','50%'],children:[{type:'vbox',padding:1,children:[{type:'text',width:'40px',id:'txtWidth',label:b.lang.common.width,onKeyUp:k,onChange:function(){o.call(this,'advanced:txtdlgGenStyle');},validate:function(){var B=this.getValue().match(i),C=!!(B&&parseInt(B[1],10)!==0);if(!C)alert(b.lang.common.invalidWidth);return C;},setup:r,commit:function(B,C,D){var E=this.getValue();if(B==d){if(E)C.setStyle('width',CKEDITOR.tools.cssLength(E));else C.removeStyle('width');!D&&C.removeAttribute('width');}else if(B==f){var F=E.match(h);if(!F){var G=this.getDialog().originalElement;if(G.getCustomData('isReady')=='true')C.setStyle('width',G.$.width+'px');}else C.setStyle('width',CKEDITOR.tools.cssLength(E));}else if(B==g){C.removeAttribute('width');C.removeStyle('width');}}},{type:'text',id:'txtHeight',width:'40px',label:b.lang.common.height,onKeyUp:k,onChange:function(){o.call(this,'advanced:txtdlgGenStyle');},validate:function(){var B=this.getValue().match(i),C=!!(B&&parseInt(B[1],10)!==0);if(!C)alert(b.lang.common.invalidHeight);return C;},setup:r,commit:function(B,C,D){var E=this.getValue();if(B==d){if(E)C.setStyle('height',CKEDITOR.tools.cssLength(E));else C.removeStyle('height');!D&&C.removeAttribute('height');}else if(B==f){var F=E.match(h);if(!F){var G=this.getDialog().originalElement;if(G.getCustomData('isReady')=='true')C.setStyle('height',G.$.height+'px');}else C.setStyle('height',CKEDITOR.tools.cssLength(E));}else if(B==g){C.removeAttribute('height');C.removeStyle('height');}}}]},{id:'ratioLock',type:'html',style:'margin-top:30px;width:40px;height:40px;',onLoad:function(){var B=CKEDITOR.document.getById(x),C=CKEDITOR.document.getById(w);if(B){B.on('click',function(D){q(this);D.data&&D.data.preventDefault();
},this.getDialog());B.on('mouseover',function(){this.addClass('cke_btn_over');},B);B.on('mouseout',function(){this.removeClass('cke_btn_over');},B);}if(C){C.on('click',function(D){var I=this;var E=p(I),F=I.originalElement,G=I.getValueOf('info','txtWidth');if(F.getCustomData('isReady')=='true'&&G){var H=F.$.height/F.$.width*G;if(!isNaN(H)){I.setValueOf('info','txtHeight',Math.round(H));l(I);}}D.data&&D.data.preventDefault();},this.getDialog());C.on('mouseover',function(){this.addClass('cke_btn_over');},C);C.on('mouseout',function(){this.removeClass('cke_btn_over');},C);}},html:'<div><a href="javascript:void(0)" tabindex="-1" title="'+b.lang.image.lockRatio+'" class="cke_btn_locked" id="'+w+'" role="checkbox"><span class="cke_icon"></span><span class="cke_label">'+b.lang.image.lockRatio+'</span></a>'+'<a href="javascript:void(0)" tabindex="-1" title="'+b.lang.image.resetSize+'" class="cke_btn_reset" id="'+x+'" role="button"><span class="cke_label">'+b.lang.image.resetSize+'</span></a>'+'</div>'}]},{type:'vbox',padding:1,children:[{type:'text',id:'txtBorder',width:'60px',label:b.lang.image.border,'default':'',onKeyUp:function(){l(this.getDialog());},onChange:function(){o.call(this,'advanced:txtdlgGenStyle');},validate:CKEDITOR.dialog.validate.integer(b.lang.image.validateBorder),setup:function(B,C){if(B==d){var D,E=C.getStyle('border-width');E=E&&E.match(/^(\d+px)(?: \1 \1 \1)?$/);D=E&&parseInt(E[1],10);isNaN(parseInt(D,10))&&(D=C.getAttribute('border'));this.setValue(D);}},commit:function(B,C,D){var E=parseInt(this.getValue(),10);if(B==d||B==f){if(!isNaN(E)){C.setStyle('border-width',CKEDITOR.tools.cssLength(E));C.setStyle('border-style','solid');}else if(!E&&this.isChanged()){C.removeStyle('border-width');C.removeStyle('border-style');C.removeStyle('border-color');}if(!D&&B==d)C.removeAttribute('border');}else if(B==g){C.removeAttribute('border');C.removeStyle('border-width');C.removeStyle('border-style');C.removeStyle('border-color');}}},{type:'text',id:'txtHSpace',width:'60px',label:b.lang.image.hSpace,'default':'',onKeyUp:function(){l(this.getDialog());},onChange:function(){o.call(this,'advanced:txtdlgGenStyle');},validate:CKEDITOR.dialog.validate.integer(b.lang.image.validateHSpace),setup:function(B,C){if(B==d){var D,E,F,G=C.getStyle('margin-left'),H=C.getStyle('margin-right');G=G&&G.match(j);H=H&&H.match(j);E=parseInt(G,10);F=parseInt(H,10);D=E==F&&E;isNaN(parseInt(D,10))&&(D=C.getAttribute('hspace'));this.setValue(D);}},commit:function(B,C,D){var E=parseInt(this.getValue(),10);
if(B==d||B==f){if(!isNaN(E)){C.setStyle('margin-left',CKEDITOR.tools.cssLength(E));C.setStyle('margin-right',CKEDITOR.tools.cssLength(E));}else if(!E&&this.isChanged()){C.removeStyle('margin-left');C.removeStyle('margin-right');}if(!D&&B==d)C.removeAttribute('hspace');}else if(B==g){C.removeAttribute('hspace');C.removeStyle('margin-left');C.removeStyle('margin-right');}}},{type:'text',id:'txtVSpace',width:'60px',label:b.lang.image.vSpace,'default':'',onKeyUp:function(){l(this.getDialog());},onChange:function(){o.call(this,'advanced:txtdlgGenStyle');},validate:CKEDITOR.dialog.validate.integer(b.lang.image.validateVSpace),setup:function(B,C){if(B==d){var D,E,F,G=C.getStyle('margin-top'),H=C.getStyle('margin-bottom');G=G&&G.match(j);H=H&&H.match(j);E=parseInt(G,10);F=parseInt(H,10);D=E==F&&E;isNaN(parseInt(D,10))&&(D=C.getAttribute('vspace'));this.setValue(D);}},commit:function(B,C,D){var E=parseInt(this.getValue(),10);if(B==d||B==f){if(!isNaN(E)){C.setStyle('margin-top',CKEDITOR.tools.cssLength(E));C.setStyle('margin-bottom',CKEDITOR.tools.cssLength(E));}else if(!E&&this.isChanged()){C.removeStyle('margin-top');C.removeStyle('margin-bottom');}if(!D&&B==d)C.removeAttribute('vspace');}else if(B==g){C.removeAttribute('vspace');C.removeStyle('margin-top');C.removeStyle('margin-bottom');}}},{id:'cmbAlign',type:'select',widths:['35%','65%'],style:'width:90px',label:b.lang.common.align,'default':'',items:[[b.lang.common.notSet,''],[b.lang.common.alignLeft,'left'],[b.lang.common.alignRight,'right']],onChange:function(){l(this.getDialog());o.call(this,'advanced:txtdlgGenStyle');},setup:function(B,C){if(B==d){var D=C.getStyle('float');switch(D){case 'inherit':case 'none':D='';}!D&&(D=(C.getAttribute('align')||'').toLowerCase());this.setValue(D);}},commit:function(B,C,D){var E=this.getValue();if(B==d||B==f){if(E)C.setStyle('float',E);else C.removeStyle('float');if(!D&&B==d){E=(C.getAttribute('align')||'').toLowerCase();switch(E){case 'left':case 'right':C.removeAttribute('align');}}}else if(B==g)C.removeStyle('float');}}]}]},{type:'vbox',height:'250px',children:[{type:'html',id:'htmlPreview',style:'width:95%;',html:'<div>'+CKEDITOR.tools.htmlEncode(b.lang.common.preview)+'<br>'+'<div id="'+y+'" class="ImagePreviewLoader" style="display:none"><div class="loading">&nbsp;</div></div>'+'<div class="ImagePreviewBox"><table><tr><td>'+'<a href="javascript:void(0)" target="_blank" onclick="return false;" id="'+z+'">'+'<img id="'+A+'" alt="" /></a>'+(b.config.image_previewText||'Prova. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas feugiat consequat diam. Maecenas metus. Vivamus diam purus, cursus a, commodo non, facilisis vitae, nulla. Aenean dictum lacinia tortor. Nunc iaculis, nibh non iaculis aliquam, orci felis euismod neque, sed ornare massa mauris sed velit. Nulla pretium mi et risus. Fusce mi pede, tempor id, cursus ac, ullamcorper nec, enim. Sed tortor. Curabitur molestie. Duis velit augue, condimentum at, ultrices a, luctus ut, orci. Donec pellentesque egestas eros. Integer cursus, augue in cursus faucibus, eros pede bibendum sem, in tempus tellus justo quis ligula. Etiam eget tortor. Vestibulum rutrum, est ut placerat elementum, lectus nisl aliquam velit, tempor aliquam eros nunc nonummy metus. In eros metus, gravida a, gravida sed, lobortis id, turpis. Ut ultrices, ipsum at venenatis fringilla, sem nulla lacinia tellus, eget aliquet turpis mauris non enim. Nam turpis. Suspendisse lacinia. Curabitur ac tortor ut ipsum egestas elementum. Nunc imperdiet gravida mauris.')+'</td></tr></table></div></div>'}]}]}]},{id:'Link',label:b.lang.link.title,padding:0,elements:[{id:'txtUrl',type:'text',label:b.lang.common.url,style:'width: 100%','default':'',setup:function(B,C){if(B==e){var D=C.data('cke-saved-href');
if(!D)D=C.getAttribute('href');this.setValue(D);}},commit:function(B,C){var E=this;if(B==e)if(E.getValue()||E.isChanged()){var D=decodeURI(E.getValue());C.data('cke-saved-href',D);C.setAttribute('href',D);if(E.getValue()||!b.config.image_removeLinkByEmptyURL)E.getDialog().addLink=true;}}},{type:'button',id:'browse',filebrowser:{action:'Browse',target:'Link:txtUrl',url:b.config.filebrowserImageBrowseLinkUrl},style:'float:right',hidden:true,label:b.lang.common.browseServer},{id:'cmbTarget',type:'select',label:b.lang.common.target,'default':'',items:[[b.lang.common.notSet,''],[b.lang.common.targetNew,'_blank'],[b.lang.common.targetTop,'_top'],[b.lang.common.targetSelf,'_self'],[b.lang.common.targetParent,'_parent']],setup:function(B,C){if(B==e)this.setValue(C.getAttribute('target')||'');},commit:function(B,C){if(B==e)if(this.getValue()||this.isChanged())C.setAttribute('target',this.getValue());}}]},{id:'Upload',hidden:true,filebrowser:'uploadButton',label:b.lang.image.upload,elements:[{type:'file',id:'upload',label:b.lang.image.btnUpload,style:'height:40px',size:38},{type:'fileButton',id:'uploadButton',filebrowser:'info:txtUrl',label:b.lang.image.btnUpload,'for':['Upload','upload']}]},{id:'advanced',label:b.lang.common.advancedTab,elements:[{type:'hbox',widths:['50%','25%','25%'],children:[{type:'text',id:'linkId',label:b.lang.common.id,setup:function(B,C){if(B==d)this.setValue(C.getAttribute('id'));},commit:function(B,C){if(B==d)if(this.getValue()||this.isChanged())C.setAttribute('id',this.getValue());}},{id:'cmbLangDir',type:'select',style:'width : 100px;',label:b.lang.common.langDir,'default':'',items:[[b.lang.common.notSet,''],[b.lang.common.langDirLtr,'ltr'],[b.lang.common.langDirRtl,'rtl']],setup:function(B,C){if(B==d)this.setValue(C.getAttribute('dir'));},commit:function(B,C){if(B==d)if(this.getValue()||this.isChanged())C.setAttribute('dir',this.getValue());}},{type:'text',id:'txtLangCode',label:b.lang.common.langCode,'default':'',setup:function(B,C){if(B==d)this.setValue(C.getAttribute('lang'));},commit:function(B,C){if(B==d)if(this.getValue()||this.isChanged())C.setAttribute('lang',this.getValue());}}]},{type:'text',id:'txtGenLongDescr',label:b.lang.common.longDescr,setup:function(B,C){if(B==d)this.setValue(C.getAttribute('longDesc'));},commit:function(B,C){if(B==d)if(this.getValue()||this.isChanged())C.setAttribute('longDesc',this.getValue());}},{type:'hbox',widths:['50%','50%'],children:[{type:'text',id:'txtGenClass',label:b.lang.common.cssClass,'default':'',setup:function(B,C){if(B==d)this.setValue(C.getAttribute('class'));
},commit:function(B,C){if(B==d)if(this.getValue()||this.isChanged())C.setAttribute('class',this.getValue());}},{type:'text',id:'txtGenTitle',label:b.lang.common.advisoryTitle,'default':'',onChange:function(){l(this.getDialog());},setup:function(B,C){if(B==d)this.setValue(C.getAttribute('title'));},commit:function(B,C){var D=this;if(B==d){if(D.getValue()||D.isChanged())C.setAttribute('title',D.getValue());}else if(B==f)C.setAttribute('title',D.getValue());else if(B==g)C.removeAttribute('title');}}]},{type:'text',id:'txtdlgGenStyle',label:b.lang.common.cssStyle,validate:CKEDITOR.dialog.validate.inlineStyle(b.lang.common.invalidInlineStyle),'default':'',setup:function(B,C){if(B==d){var D=C.getAttribute('style');if(!D&&C.$.style.cssText)D=C.$.style.cssText;this.setValue(D);var E=C.$.style.height,F=C.$.style.width,G=(E?E:'').match(h),H=(F?F:'').match(h);this.attributesInStyle={height:!!G,width:!!H};}},onChange:function(){o.call(this,['info:cmbFloat','info:cmbAlign','info:txtVSpace','info:txtHSpace','info:txtBorder','info:txtWidth','info:txtHeight']);l(this);},commit:function(B,C){if(B==d&&(this.getValue()||this.isChanged()))C.setAttribute('style',this.getValue());}}]}]};};CKEDITOR.dialog.add('image',function(b){return a(b,'image');});CKEDITOR.dialog.add('imagebutton',function(b){return a(b,'imagebutton');});})();
*/

/*
Copyright (c) 2003-2012, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

(function()
{
	var imageDialog = function( editor, dialogType )
	{
		// Load image preview.
		var IMAGE = 1,
			LINK = 2,
			PREVIEW = 4,
			CLEANUP = 8,
			regexGetSize = /^\s*(\d+)((px)|\%)?\s*$/i,
			regexGetSizeOrEmpty = /(^\s*(\d+)((px)|\%)?\s*$)|^$/i,
			pxLengthRegex = /^\d+px$/;

		var onSizeChange = function()
		{
			var value = this.getValue(),	// This = input element.
				dialog = this.getDialog(),
				aMatch  =  value.match( regexGetSize );	// Check value
			if ( aMatch )
			{
				if ( aMatch[2] == '%' )			// % is allowed - > unlock ratio.
					switchLockRatio( dialog, false );	// Unlock.
				value = aMatch[1];
			}

			// Only if ratio is locked
			if ( dialog.lockRatio )
			{
				var oImageOriginal = dialog.originalElement;
				if ( oImageOriginal.getCustomData( 'isReady' ) == 'true' )
				{
					if ( this.id == 'txtHeight' )
					{
						if ( value && value != '0' )
							value = Math.round( oImageOriginal.$.width * ( value  / oImageOriginal.$.height ) );
						if ( !isNaN( value ) )
							dialog.setValueOf( 'info', 'txtWidth', value );
					}
					else		//this.id = txtWidth.
					{
						if ( value && value != '0' )
							value = Math.round( oImageOriginal.$.height * ( value  / oImageOriginal.$.width ) );
						if ( !isNaN( value ) )
							dialog.setValueOf( 'info', 'txtHeight', value );
					}
				}
			}
			updatePreview( dialog );
		};

		var updatePreview = function( dialog )
		{
			//Don't load before onShow.
			if ( !dialog.originalElement || !dialog.preview )
				return 1;

			// Read attributes and update imagePreview;
			dialog.commitContent( PREVIEW, dialog.preview );
			return 0;
		};

		// Custom commit dialog logic, where we're intended to give inline style
		// field (txtdlgGenStyle) higher priority to avoid overwriting styles contribute
		// by other fields.
		function commitContent()
		{
			var args = arguments;
			var inlineStyleField = this.getContentElement( 'advanced', 'txtdlgGenStyle' );
			inlineStyleField && inlineStyleField.commit.apply( inlineStyleField, args );

			this.foreach( function( widget )
			{
				if ( widget.commit &&  widget.id != 'txtdlgGenStyle' )
					widget.commit.apply( widget, args );
			});
		}

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

		var switchLockRatio = function( dialog, value )
		{
			if ( !dialog.getContentElement( 'info', 'ratioLock' ) )
				return null;
				
			var oImageOriginal = dialog.originalElement;

			// Dialog may already closed. (#5505)
			if( !oImageOriginal )
				return null;

			// Check image ratio and original image ratio, but respecting user's preference.
			if ( value == 'check' )
			{
				if ( !dialog.userlockRatio && oImageOriginal.getCustomData( 'isReady' ) == 'true'  )
				{
					var width = dialog.getValueOf( 'info', 'txtWidth' ),
						height = dialog.getValueOf( 'info', 'txtHeight' ),
						originalRatio = oImageOriginal.$.width * 1000 / oImageOriginal.$.height,
						thisRatio = width * 1000 / height;
					dialog.lockRatio  = false;		// Default: unlock ratio

					if ( !width && !height )
						dialog.lockRatio = true;
					else if ( !isNaN( originalRatio ) && !isNaN( thisRatio ) )
					{
						if ( Math.round( originalRatio ) == Math.round( thisRatio ) )
							dialog.lockRatio = true;
					}
				}
			}
			else if ( value != undefined )
				dialog.lockRatio = value;
			else
			{
				dialog.userlockRatio = 1;
				dialog.lockRatio = !dialog.lockRatio;
			}

			var ratioButton = CKEDITOR.document.getById( btnLockSizesId );
			if ( dialog.lockRatio )
				ratioButton.removeClass( 'cke_btn_unlocked' );
			else
				ratioButton.addClass( 'cke_btn_unlocked' );

			ratioButton.setAttribute( 'aria-checked', dialog.lockRatio );

			// Ratio button hc presentation - WHITE SQUARE / BLACK SQUARE
			if ( CKEDITOR.env.hc )
			{
				var icon = ratioButton.getChild( 0 );
				icon.setHtml(  dialog.lockRatio ? CKEDITOR.env.ie ? '\u25A0': '\u25A3' : CKEDITOR.env.ie ? '\u25A1' : '\u25A2' );
			}

			return dialog.lockRatio;
		};

		var resetSize = function( dialog )
		{
			var oImageOriginal = dialog.originalElement;
			if ( oImageOriginal.getCustomData( 'isReady' ) == 'true' )
			{
				var widthField = dialog.getContentElement( 'info', 'txtWidth' ),
					heightField = dialog.getContentElement( 'info', 'txtHeight' );
				widthField && widthField.setValue( oImageOriginal.$.width );
				heightField && heightField.setValue( oImageOriginal.$.height );
			}
			updatePreview( dialog );
		};

		var setupDimension = function( type, element )
		{
			if ( type != IMAGE )
				return;

			function checkDimension( size, defaultValue )
			{
				var aMatch  =  size.match( regexGetSize );
				if ( aMatch )
				{
					if ( aMatch[2] == '%' )				// % is allowed.
					{
						aMatch[1] += '%';
						switchLockRatio( dialog, false );	// Unlock ratio
					}
					return aMatch[1];
				}
				return defaultValue;
			}

			var dialog = this.getDialog(),
				value = '',
				dimension = this.id == 'txtWidth' ? 'width' : 'height',
				size = element.getAttribute( dimension );

			if ( size )
				value = checkDimension( size, value );
			value = checkDimension( element.getStyle( dimension ), value );

			this.setValue( value );
		};

		var previewPreloader;

		var onImgLoadEvent = function()
		{
			// Image is ready.
			var original = this.originalElement;
			original.setCustomData( 'isReady', 'true' );
			original.removeListener( 'load', onImgLoadEvent );
			original.removeListener( 'error', onImgLoadErrorEvent );
			original.removeListener( 'abort', onImgLoadErrorEvent );

			// Hide loader
			CKEDITOR.document.getById( imagePreviewLoaderId ).setStyle( 'display', 'none' );

			// New image -> new domensions
			if ( !this.dontResetSize )
				resetSize( this );

			if ( this.firstLoad )
				CKEDITOR.tools.setTimeout( function(){ switchLockRatio( this, 'check' ); }, 0, this );

			this.firstLoad = false;
			this.dontResetSize = false;
		};

		var onImgLoadErrorEvent = function()
		{
			// Error. Image is not loaded.
			var original = this.originalElement;
			original.removeListener( 'load', onImgLoadEvent );
			original.removeListener( 'error', onImgLoadErrorEvent );
			original.removeListener( 'abort', onImgLoadErrorEvent );

			// Set Error image.
			var noimage = CKEDITOR.getUrl( editor.skinPath + 'images/noimage.png' );

			if ( this.preview )
				this.preview.setAttribute( 'src', noimage );

			// Hide loader
			CKEDITOR.document.getById( imagePreviewLoaderId ).setStyle( 'display', 'none' );
			switchLockRatio( this, false );	// Unlock.
		};

		var numbering = function( id )
			{
				return CKEDITOR.tools.getNextId() + '_' + id;
			},
			btnLockSizesId = numbering( 'btnLockSizes' ),
			btnResetSizeId = numbering( 'btnResetSize' ),
			imagePreviewLoaderId = numbering( 'ImagePreviewLoader' ),
			previewLinkId = numbering( 'previewLink' ),
			previewImageId = numbering( 'previewImage' );

		return {
			title : editor.lang.image[ dialogType == 'image' ? 'title' : 'titleButton' ],
			minWidth : 420,
			minHeight : 360,
			onShow : function()
			{
				this.imageElement = false;
				this.linkElement = false;

				// Default: create a new element.
				this.imageEditMode = false;
				this.linkEditMode = false;

				this.lockRatio = true;
				this.userlockRatio = 0;
				this.dontResetSize = false;
				this.firstLoad = true;
				this.addLink = false;

				var editor = this.getParentEditor(),
					sel = editor.getSelection(),
					element = sel && sel.getSelectedElement(),
					link = element && element.getAscendant( 'a' );

				//Hide loader.
				CKEDITOR.document.getById( imagePreviewLoaderId ).setStyle( 'display', 'none' );
				// Create the preview before setup the dialog contents.
				previewPreloader = new CKEDITOR.dom.element( 'img', editor.document );
				this.preview = CKEDITOR.document.getById( previewImageId );

				// Copy of the image
				this.originalElement = editor.document.createElement( 'img' );
				this.originalElement.setAttribute( 'alt', '' );
				this.originalElement.setCustomData( 'isReady', 'false' );

				if ( link )
				{
					this.linkElement = link;
					this.linkEditMode = true;

					// Look for Image element.
					var linkChildren = link.getChildren();
					if ( linkChildren.count() == 1 )			// 1 child.
					{
						var childTagName = linkChildren.getItem( 0 ).getName();
						if ( childTagName == 'img' || childTagName == 'input' )
						{
							this.imageElement = linkChildren.getItem( 0 );
							if ( this.imageElement.getName() == 'img' )
								this.imageEditMode = 'img';
							else if ( this.imageElement.getName() == 'input' )
								this.imageEditMode = 'input';
						}
					}
					// Fill out all fields.
					if ( dialogType == 'image' )
						this.setupContent( LINK, link );
				}

				if ( element && element.getName() == 'img' && !element.data( 'cke-realelement' )
					|| element && element.getName() == 'input' && element.getAttribute( 'type' ) == 'image' )
				{
					this.imageEditMode = element.getName();
					this.imageElement = element;
				}

				if ( this.imageEditMode )
				{
					// Use the original element as a buffer from  since we don't want
					// temporary changes to be committed, e.g. if the dialog is canceled.
					this.cleanImageElement = this.imageElement;
					this.imageElement = this.cleanImageElement.clone( true, true );

					// Fill out all fields.
					this.setupContent( IMAGE, this.imageElement );
				}
				else
					this.imageElement =  editor.document.createElement( 'img' );

				// Refresh LockRatio button
				switchLockRatio ( this, true );

				// Dont show preview if no URL given.
				if ( !CKEDITOR.tools.trim( this.getValueOf( 'info', 'txtUrl' ) ) )
				{
					this.preview.removeAttribute( 'src' );
					this.preview.setStyle( 'display', 'none' );
				}
			},
			onOk : function()
			{
				// Edit existing Image.
				if ( this.imageEditMode )
				{
					var imgTagName = this.imageEditMode;

					// Image dialog and Input element.
					if ( dialogType == 'image' && imgTagName == 'input' && confirm( editor.lang.image.button2Img ) )
					{
						// Replace INPUT-> IMG
						imgTagName = 'img';
						this.imageElement = editor.document.createElement( 'img' );
						this.imageElement.setAttribute( 'alt', '' );
						editor.insertElement( this.imageElement );
					}
					// ImageButton dialog and Image element.
					else if ( dialogType != 'image' && imgTagName == 'img' && confirm( editor.lang.image.img2Button ))
					{
						// Replace IMG -> INPUT
						imgTagName = 'input';
						this.imageElement = editor.document.createElement( 'input' );
						this.imageElement.setAttributes(
							{
								type : 'image',
								alt : ''
							}
						);
						editor.insertElement( this.imageElement );
					}
					else
					{
						// Restore the original element before all commits.
						this.imageElement = this.cleanImageElement;
						delete this.cleanImageElement;
					}
				}
				else	// Create a new image.
				{
					// Image dialog -> create IMG element.
					if ( dialogType == 'image' )
						this.imageElement = editor.document.createElement( 'img' );
					else
					{
						this.imageElement = editor.document.createElement( 'input' );
						this.imageElement.setAttribute ( 'type' ,'image' );
					}
					this.imageElement.setAttribute( 'alt', '' );
				}

				// Create a new link.
				if ( !this.linkEditMode )
					this.linkElement = editor.document.createElement( 'a' );

				// Set attributes.
				this.commitContent( IMAGE, this.imageElement );
				this.commitContent( LINK, this.linkElement );

				// Remove empty style attribute.
				if ( !this.imageElement.getAttribute( 'style' ) )
					this.imageElement.removeAttribute( 'style' );

				// Insert a new Image.
				if ( !this.imageEditMode )
				{
					if ( this.addLink )
					{
						//Insert a new Link.
						if ( !this.linkEditMode )
						{
							editor.insertElement( this.linkElement );
							this.linkElement.append( this.imageElement, false );
						}
						else	 //Link already exists, image not.
							editor.insertElement( this.imageElement );
					}
					else
						editor.insertElement( this.imageElement );
				}
				else		// Image already exists.
				{
					//Add a new link element.
					if ( !this.linkEditMode && this.addLink )
					{
						editor.insertElement( this.linkElement );
						this.imageElement.appendTo( this.linkElement );
					}
					//Remove Link, Image exists.
					else if ( this.linkEditMode && !this.addLink )
					{
						editor.getSelection().selectElement( this.linkElement );
						editor.insertElement( this.imageElement );
					}
				}
			},
			onLoad : function()
			{
				//Modifica - Giacomo
				//urlFileManager = urlFileManager + this.getContentElement( 'info', 'txtUrl' )._.inputId;
				this.getContentElement( 'info', 'browse' ).filebrowser.url += this.getContentElement( 'info', 'txtUrl' )._.inputId;
				//alert(this.getContentElement( 'info', 'browse' ).filebrowser.url);				
			
				if ( dialogType != 'image' )
					this.hidePage( 'Link' );		//Hide Link tab.
				var doc = this._.element.getDocument();

				if ( this.getContentElement( 'info', 'ratioLock' ) )
				{
					this.addFocusable( doc.getById( btnResetSizeId ), 5 );
					this.addFocusable( doc.getById( btnLockSizesId ), 5 );
				}

				this.commitContent = commitContent;
			},
			onHide : function()
			{			
				if ( this.preview )
					this.commitContent( CLEANUP, this.preview );

				if ( this.originalElement )
				{
					this.originalElement.removeListener( 'load', onImgLoadEvent );
					this.originalElement.removeListener( 'error', onImgLoadErrorEvent );
					this.originalElement.removeListener( 'abort', onImgLoadErrorEvent );
					this.originalElement.remove();
					this.originalElement = false;		// Dialog is closed.
				}

				delete this.imageElement;
			},
			contents : [
				{
					id : 'info',
					label : editor.lang.image.infoTab,
					accessKey : 'I',
					elements :
					[
						{
							type : 'vbox',
							padding : 0,
							children :
							[
								{
									type : 'hbox',
									widths : [ '280px', '110px' ],
									align : 'right',
									children :
									[
										{
											id : 'txtUrl',
											type : 'text',
											label : editor.lang.common.url,
											required: true,
											onChange : function()
											{
												var dialog = this.getDialog(),
													newUrl = this.getValue();

												//Update original image
												if ( newUrl.length > 0 )	//Prevent from load before onShow
												{
													dialog = this.getDialog();
													var original = dialog.originalElement;

													dialog.preview.removeStyle( 'display' );

													original.setCustomData( 'isReady', 'false' );
													// Show loader
													var loader = CKEDITOR.document.getById( imagePreviewLoaderId );
													if ( loader )
														loader.setStyle( 'display', '' );

													original.on( 'load', onImgLoadEvent, dialog );
													original.on( 'error', onImgLoadErrorEvent, dialog );
													original.on( 'abort', onImgLoadErrorEvent, dialog );
													original.setAttribute( 'src', newUrl );

													// Query the preloader to figure out the url impacted by based href.
													previewPreloader.setAttribute( 'src', newUrl );
													dialog.preview.setAttribute( 'src', previewPreloader.$.src );
													updatePreview( dialog );
												}
												// Dont show preview if no URL given.
												else if ( dialog.preview )
												{
													dialog.preview.removeAttribute( 'src' );
													dialog.preview.setStyle( 'display', 'none' );
												}
											},
											setup : function( type, element )
											{
												if ( type == IMAGE )
												{
													var url = element.data( 'cke-saved-src' ) || element.getAttribute( 'src' );
													var field = this;

													this.getDialog().dontResetSize = true;

													field.setValue( url );		// And call this.onChange()
													// Manually set the initial value.(#4191)
													field.setInitValue();
												}
											},
											commit : function( type, element )
											{
												if ( type == IMAGE && ( this.getValue() || this.isChanged() ) )
												{
													element.data( 'cke-saved-src', this.getValue() );
													element.setAttribute( 'src', this.getValue() );
												}
												else if ( type == CLEANUP )
												{
													element.setAttribute( 'src', '' );	// If removeAttribute doesn't work.
													element.removeAttribute( 'src' );
												}
											},
											validate : CKEDITOR.dialog.validate.notEmpty( editor.lang.image.urlMissing )
										},
										{
											type : 'button',
											id : 'browse',
											// v-align with the 'txtUrl' field.
											// TODO: We need something better than a fixed size here.
											style : 'display:inline-block;margin-top:10px;',
											align : 'center',
											label : editor.lang.common.browseServer,
											hidden : true,
											filebrowser : 'info:txtUrl',
											target : 'info:txtUrl'
											//hidden : false,
											//filebrowser :
											//{
											//	action : 'Browse',
											//	target: 'info:txtUrl',
											//	url: UrlFileManager
											//}
										}								
									]
								}
							]
						},
						{
							id : 'txtAlt',
							type : 'text',
							label : editor.lang.image.alt,
							accessKey : 'T',
							'default' : '',
							onChange : function()
							{
								updatePreview( this.getDialog() );
							},
							setup : function( type, element )
							{
								if ( type == IMAGE )
									this.setValue( element.getAttribute( 'alt' ) );
							},
							commit : function( type, element )
							{
								if ( type == IMAGE )
								{
									if ( this.getValue() || this.isChanged() )
										element.setAttribute( 'alt', this.getValue() );
								}
								else if ( type == PREVIEW )
								{
									element.setAttribute( 'alt', this.getValue() );
								}
								else if ( type == CLEANUP )
								{
									element.removeAttribute( 'alt' );
								}
							}
						},
						{
							type : 'hbox',
							children :
							[
								{
									id : 'basic',
									type : 'vbox',
									children :
									[
										{
											type : 'hbox',
											widths : [ '50%', '50%' ],
											children :
											[
												{
													type : 'vbox',
													padding : 1,
													children :
													[
														{
															type : 'text',
															width: '40px',
															id : 'txtWidth',
															label : editor.lang.common.width,
															onKeyUp : onSizeChange,
															onChange : function()
															{
																commitInternally.call( this, 'advanced:txtdlgGenStyle' );
															},
															validate : function()
															{
																var aMatch  =  this.getValue().match( regexGetSizeOrEmpty ),
																	isValid = !!( aMatch && parseInt( aMatch[1], 10 ) !== 0 );
																if ( !isValid )
																	alert( editor.lang.common.invalidWidth );
																return isValid;
															},
															setup : setupDimension,
															commit : function( type, element, internalCommit )
															{
																var value = this.getValue();
																if ( type == IMAGE )
																{
																	if ( value )
																		element.setStyle( 'width', CKEDITOR.tools.cssLength( value ) );
																	else
																		element.removeStyle( 'width' );

																	!internalCommit && element.removeAttribute( 'width' );
																}
																else if ( type == PREVIEW )
																{
																	var aMatch = value.match( regexGetSize );
																	if ( !aMatch )
																	{
																		var oImageOriginal = this.getDialog().originalElement;
																		if ( oImageOriginal.getCustomData( 'isReady' ) == 'true' )
																			element.setStyle( 'width',  oImageOriginal.$.width + 'px');
																	}
																	else
																		element.setStyle( 'width', CKEDITOR.tools.cssLength( value ) );
																}
																else if ( type == CLEANUP )
																{
																	element.removeAttribute( 'width' );
																	element.removeStyle( 'width' );
																}
															}
														},
														{
															type : 'text',
															id : 'txtHeight',
															width: '40px',
															label : editor.lang.common.height,
															onKeyUp : onSizeChange,
															onChange : function()
															{
																commitInternally.call( this, 'advanced:txtdlgGenStyle' );
															},
															validate : function()
															{
																var aMatch = this.getValue().match( regexGetSizeOrEmpty ),
																	isValid = !!( aMatch && parseInt( aMatch[1], 10 ) !== 0 );
																if ( !isValid )
																	alert( editor.lang.common.invalidHeight );
																return isValid;
															},
															setup : setupDimension,
															commit : function( type, element, internalCommit )
															{
																var value = this.getValue();
																if ( type == IMAGE )
																{
																	if ( value )
																		element.setStyle( 'height', CKEDITOR.tools.cssLength( value ) );
																	else
																		element.removeStyle( 'height' );

																	!internalCommit && element.removeAttribute( 'height' );
																}
																else if ( type == PREVIEW )
																{
																	var aMatch = value.match( regexGetSize );
																	if ( !aMatch )
																	{
																		var oImageOriginal = this.getDialog().originalElement;
																		if ( oImageOriginal.getCustomData( 'isReady' ) == 'true' )
																			element.setStyle( 'height', oImageOriginal.$.height + 'px' );
																	}
																	else
																		element.setStyle( 'height',  CKEDITOR.tools.cssLength( value ) );
																}
																else if ( type == CLEANUP )
																{
																	element.removeAttribute( 'height' );
																	element.removeStyle( 'height' );
																}
															}
														}
													]
												},
												{
													id : 'ratioLock',
													type : 'html',
													style : 'margin-top:30px;width:40px;height:40px;',
													onLoad : function()
													{
														// Activate Reset button
														var	resetButton = CKEDITOR.document.getById( btnResetSizeId ),
															ratioButton = CKEDITOR.document.getById( btnLockSizesId );
														if ( resetButton )
														{
															resetButton.on( 'click', function( evt )
																{
																	resetSize( this );
																	evt.data && evt.data.preventDefault();
																}, this.getDialog() );
															resetButton.on( 'mouseover', function()
																{
																	this.addClass( 'cke_btn_over' );
																}, resetButton );
															resetButton.on( 'mouseout', function()
																{
																	this.removeClass( 'cke_btn_over' );
																}, resetButton );
														}
														// Activate (Un)LockRatio button
														if ( ratioButton )
														{
															ratioButton.on( 'click', function(evt)
																{
																	var locked = switchLockRatio( this ),
																		oImageOriginal = this.originalElement,
																		width = this.getValueOf( 'info', 'txtWidth' );

																	if ( oImageOriginal.getCustomData( 'isReady' ) == 'true' && width )
																	{
																		var height = oImageOriginal.$.height / oImageOriginal.$.width * width;
																		if ( !isNaN( height ) )
																		{
																			this.setValueOf( 'info', 'txtHeight', Math.round( height ) );
																			updatePreview( this );
																		}
																	}
																	evt.data && evt.data.preventDefault();
																}, this.getDialog() );
															ratioButton.on( 'mouseover', function()
																{
																	this.addClass( 'cke_btn_over' );
																}, ratioButton );
															ratioButton.on( 'mouseout', function()
																{
																	this.removeClass( 'cke_btn_over' );
																}, ratioButton );
														}
													},
													html : '<div>'+
														'<a href="javascript:void(0)" tabindex="-1" title="' + editor.lang.image.lockRatio +
														'" class="cke_btn_locked" id="' + btnLockSizesId + '" role="checkbox"><span class="cke_icon"></span><span class="cke_label">' + editor.lang.image.lockRatio + '</span></a>' +
														'<a href="javascript:void(0)" tabindex="-1" title="' + editor.lang.image.resetSize +
														'" class="cke_btn_reset" id="' + btnResetSizeId + '" role="button"><span class="cke_label">' + editor.lang.image.resetSize + '</span></a>'+
														'</div>'
												}
											]
										},
										{
											type : 'vbox',
											padding : 1,
											children :
											[
												{
													type : 'text',
													id : 'txtBorder',
													width: '60px',
													label : editor.lang.image.border,
													'default' : '',
													onKeyUp : function()
													{
														updatePreview( this.getDialog() );
													},
													onChange : function()
													{
														commitInternally.call( this, 'advanced:txtdlgGenStyle' );
													},
													validate : CKEDITOR.dialog.validate.integer( editor.lang.image.validateBorder ),
													setup : function( type, element )
													{
														if ( type == IMAGE )
														{
															var value,
																borderStyle = element.getStyle( 'border-width' );
															borderStyle = borderStyle && borderStyle.match( /^(\d+px)(?: \1 \1 \1)?$/ );
															value = borderStyle && parseInt( borderStyle[ 1 ], 10 );
															isNaN ( parseInt( value, 10 ) ) && ( value = element.getAttribute( 'border' ) );
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
																element.setStyle( 'border-width', CKEDITOR.tools.cssLength( value ) );
																element.setStyle( 'border-style', 'solid' );
															}
															else if ( !value && this.isChanged() )
															{
																element.removeStyle( 'border-width' );
																element.removeStyle( 'border-style' );
																element.removeStyle( 'border-color' );
															}

															if ( !internalCommit && type == IMAGE )
																element.removeAttribute( 'border' );
														}
														else if ( type == CLEANUP )
														{
															element.removeAttribute( 'border' );
															element.removeStyle( 'border-width' );
															element.removeStyle( 'border-style' );
															element.removeStyle( 'border-color' );
														}
													}
												},
												{
													type : 'text',
													id : 'txtHSpace',
													width: '60px',
													label : editor.lang.image.hSpace,
													'default' : '',
													onKeyUp : function()
													{
														updatePreview( this.getDialog() );
													},
													onChange : function()
													{
														commitInternally.call( this, 'advanced:txtdlgGenStyle' );
													},
													validate : CKEDITOR.dialog.validate.integer( editor.lang.image.validateHSpace ),
													setup : function( type, element )
													{
														if ( type == IMAGE )
														{
															var value,
																marginLeftPx,
																marginRightPx,
																marginLeftStyle = element.getStyle( 'margin-left' ),
																marginRightStyle = element.getStyle( 'margin-right' );

															marginLeftStyle = marginLeftStyle && marginLeftStyle.match( pxLengthRegex );
															marginRightStyle = marginRightStyle && marginRightStyle.match( pxLengthRegex );
															marginLeftPx = parseInt( marginLeftStyle, 10 );
															marginRightPx = parseInt( marginRightStyle, 10 );

															value = ( marginLeftPx == marginRightPx ) && marginLeftPx;
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
																element.setStyle( 'margin-right', CKEDITOR.tools.cssLength( value ) );
															}
															else if ( !value && this.isChanged( ) )
															{
																element.removeStyle( 'margin-left' );
																element.removeStyle( 'margin-right' );
															}

															if ( !internalCommit && type == IMAGE )
																element.removeAttribute( 'hspace' );
														}
														else if ( type == CLEANUP )
														{
															element.removeAttribute( 'hspace' );
															element.removeStyle( 'margin-left' );
															element.removeStyle( 'margin-right' );
														}
													}
												},
												{
													type : 'text',
													id : 'txtVSpace',
													width : '60px',
													label : editor.lang.image.vSpace,
													'default' : '',
													onKeyUp : function()
													{
														updatePreview( this.getDialog() );
													},
													onChange : function()
													{
														commitInternally.call( this, 'advanced:txtdlgGenStyle' );
													},
													validate : CKEDITOR.dialog.validate.integer( editor.lang.image.validateVSpace ),
													setup : function( type, element )
													{
														if ( type == IMAGE )
														{
															var value,
																marginTopPx,
																marginBottomPx,
																marginTopStyle = element.getStyle( 'margin-top' ),
																marginBottomStyle = element.getStyle( 'margin-bottom' );

															marginTopStyle = marginTopStyle && marginTopStyle.match( pxLengthRegex );
															marginBottomStyle = marginBottomStyle && marginBottomStyle.match( pxLengthRegex );
															marginTopPx = parseInt( marginTopStyle, 10 );
															marginBottomPx = parseInt( marginBottomStyle, 10 );

															value = ( marginTopPx == marginBottomPx ) && marginTopPx;
															isNaN ( parseInt( value, 10 ) ) && ( value = element.getAttribute( 'vspace' ) );
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
																element.setStyle( 'margin-bottom', CKEDITOR.tools.cssLength( value ) );
															}
															else if ( !value && this.isChanged( ) )
															{
																element.removeStyle( 'margin-top' );
																element.removeStyle( 'margin-bottom' );
															}

															if ( !internalCommit && type == IMAGE )
																element.removeAttribute( 'vspace' );
														}
														else if ( type == CLEANUP )
														{
															element.removeAttribute( 'vspace' );
															element.removeStyle( 'margin-top' );
															element.removeStyle( 'margin-bottom' );
														}
													}
												},
												{
													id : 'cmbAlign',
													type : 'select',
													widths : [ '35%','65%' ],
													style : 'width:90px',
													label : editor.lang.common.align,
													'default' : '',
													items :
													[
														[ editor.lang.common.notSet , ''],
														[ editor.lang.common.alignLeft , 'left'],
														[ editor.lang.common.alignRight , 'right']
														// Backward compatible with v2 on setup when specified as attribute value,
														// while these values are no more available as select options.
														//	[ editor.lang.image.alignAbsBottom , 'absBottom'],
														//	[ editor.lang.image.alignAbsMiddle , 'absMiddle'],
														//  [ editor.lang.image.alignBaseline , 'baseline'],
														//  [ editor.lang.image.alignTextTop , 'text-top'],
														//  [ editor.lang.image.alignBottom , 'bottom'],
														//  [ editor.lang.image.alignMiddle , 'middle'],
														//  [ editor.lang.image.alignTop , 'top']
													],
													onChange : function()
													{
														updatePreview( this.getDialog() );
														commitInternally.call( this, 'advanced:txtdlgGenStyle' );
													},
													setup : function( type, element )
													{
														if ( type == IMAGE )
														{
															var value = element.getStyle( 'float' );
															switch( value )
															{
																// Ignore those unrelated values.
																case 'inherit':
																case 'none':
																	value = '';
															}

															!value && ( value = ( element.getAttribute( 'align' ) || '' ).toLowerCase() );
															this.setValue( value );
														}
													},
													commit : function( type, element, internalCommit )
													{
														var value = this.getValue();
														if ( type == IMAGE || type == PREVIEW )
														{
															if ( value )
																element.setStyle( 'float', value );
															else
																element.removeStyle( 'float' );

															if ( !internalCommit && type == IMAGE )
															{
																value = ( element.getAttribute( 'align' ) || '' ).toLowerCase();
																switch( value )
																{
																	// we should remove it only if it matches "left" or "right",
																	// otherwise leave it intact.
																	case 'left':
																	case 'right':
																		element.removeAttribute( 'align' );
																}
															}
														}
														else if ( type == CLEANUP )
															element.removeStyle( 'float' );

													}
												}
											]
										}
									]
								},
								{
									type : 'vbox',
									height : '250px',
									children :
									[
										{
											type : 'html',
											id : 'htmlPreview',
											style : 'width:95%;',
											html : '<div>' + CKEDITOR.tools.htmlEncode( editor.lang.common.preview ) +'<br>'+
											'<div id="' + imagePreviewLoaderId + '" class="ImagePreviewLoader" style="display:none"><div class="loading">&nbsp;</div></div>'+
											'<div class="ImagePreviewBox"><table><tr><td>'+
											'<a href="javascript:void(0)" target="_blank" onclick="return false;" id="' + previewLinkId + '">'+
											'<img id="' + previewImageId + '" alt="" /></a>' +
											( editor.config.image_previewText ||
											'' ) +
											'</td></tr></table></div></div>'
										}
									]
								}
							]
						}
					]
				},
				{
					id : 'Link',
					label : editor.lang.link.title,
					padding : 0,
					elements :
					[
						{
							id : 'txtUrl',
							type : 'text',
							label : editor.lang.common.url,
							style : 'width: 100%',
							'default' : '',
							setup : function( type, element )
							{
								if ( type == LINK )
								{
									var href = element.data( 'cke-saved-href' );
									if ( !href )
										href = element.getAttribute( 'href' );
									this.setValue( href );
								}
							},
							commit : function( type, element )
							{
								if ( type == LINK )
								{
									if ( this.getValue() || this.isChanged() )
									{
										var url = decodeURI( this.getValue() );
										element.data( 'cke-saved-href', url );
										element.setAttribute( 'href', url );

										if ( this.getValue() || !editor.config.image_removeLinkByEmptyURL )
											this.getDialog().addLink = true;
									}
								}
							}
						},
						{
							type : 'button',
							id : 'browse',
							filebrowser :
							{
								action : 'Browse',
								target: 'Link:txtUrl',
								url: editor.config.filebrowserImageBrowseLinkUrl
							},
							style : 'float:right',
							hidden : true,
							label : editor.lang.common.browseServer
						},
						{
							id : 'cmbTarget',
							type : 'select',
							label : editor.lang.common.target,
							'default' : '',
							items :
							[
								[ editor.lang.common.notSet , ''],
								[ editor.lang.common.targetNew , '_blank'],
								[ editor.lang.common.targetTop , '_top'],
								[ editor.lang.common.targetSelf , '_self'],
								[ editor.lang.common.targetParent , '_parent']
							],
							setup : function( type, element )
							{
								if ( type == LINK )
									this.setValue( element.getAttribute( 'target' ) || '' );
							},
							commit : function( type, element )
							{
								if ( type == LINK )
								{
									if ( this.getValue() || this.isChanged() )
										element.setAttribute( 'target', this.getValue() );
								}
							}
						}
					]
				},
				{
					id : 'Upload',
					hidden : true,
					filebrowser : 'uploadButton',
					label : editor.lang.image.upload,
					elements :
					[
						{
							type : 'file',
							id : 'upload',
							label : editor.lang.image.btnUpload,
							style: 'height:40px',
							size : 38
						},
						{
							type : 'fileButton',
							id : 'uploadButton',
							filebrowser : 'info:txtUrl',
							label : editor.lang.image.btnUpload,
							'for' : [ 'Upload', 'upload' ]
						}
					]
				},
				{
					id : 'advanced',
					label : editor.lang.common.advancedTab,
					elements :
					[
						{
							type : 'hbox',
							widths : [ '50%', '25%', '25%' ],
							children :
							[
								{
									type : 'text',
									id : 'linkId',
									label : editor.lang.common.id,
									setup : function( type, element )
									{
										if ( type == IMAGE )
											this.setValue( element.getAttribute( 'id' ) );
									},
									commit : function( type, element )
									{
										if ( type == IMAGE )
										{
											if ( this.getValue() || this.isChanged() )
												element.setAttribute( 'id', this.getValue() );
										}
									}
								},
								{
									id : 'cmbLangDir',
									type : 'select',
									style : 'width : 100px;',
									label : editor.lang.common.langDir,
									'default' : '',
									items :
									[
										[ editor.lang.common.notSet, '' ],
										[ editor.lang.common.langDirLtr, 'ltr' ],
										[ editor.lang.common.langDirRtl, 'rtl' ]
									],
									setup : function( type, element )
									{
										if ( type == IMAGE )
											this.setValue( element.getAttribute( 'dir' ) );
									},
									commit : function( type, element )
									{
										if ( type == IMAGE )
										{
											if ( this.getValue() || this.isChanged() )
												element.setAttribute( 'dir', this.getValue() );
										}
									}
								},
								{
									type : 'text',
									id : 'txtLangCode',
									label : editor.lang.common.langCode,
									'default' : '',
									setup : function( type, element )
									{
										if ( type == IMAGE )
											this.setValue( element.getAttribute( 'lang' ) );
									},
									commit : function( type, element )
									{
										if ( type == IMAGE )
										{
											if ( this.getValue() || this.isChanged() )
												element.setAttribute( 'lang', this.getValue() );
										}
									}
								}
							]
						},
						{
							type : 'text',
							id : 'txtGenLongDescr',
							label : editor.lang.common.longDescr,
							setup : function( type, element )
							{
								if ( type == IMAGE )
									this.setValue( element.getAttribute( 'longDesc' ) );
							},
							commit : function( type, element )
							{
								if ( type == IMAGE )
								{
									if ( this.getValue() || this.isChanged() )
										element.setAttribute( 'longDesc', this.getValue() );
								}
							}
						},
						{
							type : 'hbox',
							widths : [ '50%', '50%' ],
							children :
							[
								{
									type : 'text',
									id : 'txtGenClass',
									label : editor.lang.common.cssClass,
									'default' : '',
									setup : function( type, element )
									{
										if ( type == IMAGE )
											this.setValue( element.getAttribute( 'class' ) );
									},
									commit : function( type, element )
									{
										if ( type == IMAGE )
										{
											if ( this.getValue() || this.isChanged() )
												element.setAttribute( 'class', this.getValue() );
										}
									}
								},
								{
									type : 'text',
									id : 'txtGenTitle',
									label : editor.lang.common.advisoryTitle,
									'default' : '',
									onChange : function()
									{
										updatePreview( this.getDialog() );
									},
									setup : function( type, element )
									{
										if ( type == IMAGE )
											this.setValue( element.getAttribute( 'title' ) );
									},
									commit : function( type, element )
									{
										if ( type == IMAGE )
										{
											if ( this.getValue() || this.isChanged() )
												element.setAttribute( 'title', this.getValue() );
										}
										else if ( type == PREVIEW )
										{
											element.setAttribute( 'title', this.getValue() );
										}
										else if ( type == CLEANUP )
										{
											element.removeAttribute( 'title' );
										}
									}
								}
							]
						},
						{
							type : 'text',
							id : 'txtdlgGenStyle',
							label : editor.lang.common.cssStyle,
							validate : CKEDITOR.dialog.validate.inlineStyle( editor.lang.common.invalidInlineStyle ),
							'default' : '',
							setup : function( type, element )
							{
								if ( type == IMAGE )
								{
									var genStyle = element.getAttribute( 'style' );
									if ( !genStyle && element.$.style.cssText )
										genStyle = element.$.style.cssText;
									this.setValue( genStyle );

									var height = element.$.style.height,
										width = element.$.style.width,
										aMatchH  = ( height ? height : '' ).match( regexGetSize ),
										aMatchW  = ( width ? width : '').match( regexGetSize );

									this.attributesInStyle =
									{
										height : !!aMatchH,
										width : !!aMatchW
									};
								}
							},
							onChange : function ()
							{
								commitInternally.call( this,
									[ 'info:cmbFloat', 'info:cmbAlign',
									  'info:txtVSpace', 'info:txtHSpace',
									  'info:txtBorder',
									  'info:txtWidth', 'info:txtHeight' ] );
								updatePreview( this );
							},
							commit : function( type, element )
							{
								if ( type == IMAGE && ( this.getValue() || this.isChanged() ) )
								{
									element.setAttribute( 'style', this.getValue() );
								}
							}
						}
					]
				}
			]
		};
	};

	CKEDITOR.dialog.add( 'image', function( editor )
		{	
			return imageDialog( editor, 'image' );	
		});

	CKEDITOR.dialog.add( 'imagebutton', function( editor )
		{
			return imageDialog( editor, 'imagebutton' );
		});
		
		
})();
