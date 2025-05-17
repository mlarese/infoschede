/*
Copyright (c) 2003-2012, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

//CKEDITOR.dialog.add('link',function(a){var b=CKEDITOR.plugins.link,c=function(){var F=this.getDialog(),G=F.getContentElement('target','popupFeatures'),H=F.getContentElement('target','linkTargetName'),I=this.getValue();if(!G||!H)return;G=G.getElement();G.hide();H.setValue('');switch(I){case 'frame':H.setLabel(a.lang.link.targetFrameName);H.getElement().show();break;case 'popup':G.show();H.setLabel(a.lang.link.targetPopupName);H.getElement().show();break;default:H.setValue(I);H.getElement().hide();break;}},d=function(){var F=this.getDialog(),G=['urlOptions','anchorOptions','emailOptions'],H=this.getValue(),I=F.definition.getContents('upload'),J=I&&I.hidden;if(H=='url'){if(a.config.linkShowTargetTab)F.showPage('target');if(!J)F.showPage('upload');}else{F.hidePage('target');if(!J)F.hidePage('upload');}for(var K=0;K<G.length;K++){var L=F.getContentElement('info',G[K]);if(!L)continue;L=L.getElement().getParent().getParent();if(G[K]==H+'Options')L.show();else L.hide();}F.layout();},e=/^javascript:/,f=/^mailto:([^?]+)(?:\?(.+))?$/,g=/subject=([^;?:@&=$,\/]*)/,h=/body=([^;?:@&=$,\/]*)/,i=/^#(.*)$/,j=/^((?:http|https|ftp|news):\/\/)?(.*)$/,k=/^(_(?:self|top|parent|blank))$/,l=/^javascript:void\(location\.href='mailto:'\+String\.fromCharCode\(([^)]+)\)(?:\+'(.*)')?\)$/,m=/^javascript:([^(]+)\(([^)]+)\)$/,n=/\s*window.open\(\s*this\.href\s*,\s*(?:'([^']*)'|null)\s*,\s*'([^']*)'\s*\)\s*;\s*return\s*false;*\s*/,o=/(?:^|,)([^=]+)=(\d+|yes|no)/gi,p=function(F,G){var H=G&&(G.data('cke-saved-href')||G.getAttribute('href'))||'',I,J,K,L,M={};if(I=H.match(e))if(y=='encode')H=H.replace(l,function(ae,af,ag){return 'mailto:'+String.fromCharCode.apply(String,af.split(','))+(ag&&w(ag));});else if(y)H.replace(m,function(ae,af,ag){if(af==z.name){M.type='email';var ah=M.email={},ai=/[^,\s]+/g,aj=/(^')|('$)/g,ak=ag.match(ai),al=ak.length,am,an;for(var ao=0;ao<al;ao++){an=decodeURIComponent(w(ak[ao].replace(aj,'')));am=z.params[ao].toLowerCase();ah[am]=an;}ah.address=[ah.name,ah.domain].join('@');}});if(!M.type)if(K=H.match(i)){M.type='anchor';M.anchor={};M.anchor.name=M.anchor.id=K[1];}else if(J=H.match(f)){var N=H.match(g),O=H.match(h);M.type='email';var P=M.email={};P.address=J[1];N&&(P.subject=decodeURIComponent(N[1]));O&&(P.body=decodeURIComponent(O[1]));}else if(H&&(L=H.match(j))){M.type='url';M.url={};M.url.protocol=L[1];M.url.url=L[2];}else M.type='url';if(G){var Q=G.getAttribute('target');M.target={};M.adv={};if(!Q){var R=G.data('cke-pa-onclick')||G.getAttribute('onclick'),S=R&&R.match(n);
//if(S){M.target.type='popup';M.target.name=S[1];var T;while(T=o.exec(S[2])){if((T[2]=='yes'||T[2]=='1')&&!(T[1] in {height:1,width:1,top:1,left:1}))M.target[T[1]]=true;else if(isFinite(T[2]))M.target[T[1]]=T[2];}}}else{var U=Q.match(k);if(U)M.target.type=M.target.name=Q;else{M.target.type='frame';M.target.name=Q;}}var V=this,W=function(ae,af){var ag=G.getAttribute(af);if(ag!==null)M.adv[ae]=ag||'';};W('advId','id');W('advLangDir','dir');W('advAccessKey','accessKey');M.adv.advName=G.data('cke-saved-name')||G.getAttribute('name')||'';W('advLangCode','lang');W('advTabIndex','tabindex');W('advTitle','title');W('advContentType','type');CKEDITOR.plugins.link.synAnchorSelector?M.adv.advCSSClasses=C(G):W('advCSSClasses','class');W('advCharset','charset');W('advStyles','style');W('advRel','rel');}var X=M.anchors=[],Y,Z,aa;if(CKEDITOR.plugins.link.emptyAnchorFix){var ab=F.document.getElementsByTag('a');for(Y=0,Z=ab.count();Y<Z;Y++){aa=ab.getItem(Y);if(aa.data('cke-saved-name')||aa.hasAttribute('name'))X.push({name:aa.data('cke-saved-name')||aa.getAttribute('name'),id:aa.getAttribute('id')});}}else{var ac=new CKEDITOR.dom.nodeList(F.document.$.anchors);for(Y=0,Z=ac.count();Y<Z;Y++){aa=ac.getItem(Y);X[Y]={name:aa.getAttribute('name'),id:aa.getAttribute('id')};}}if(CKEDITOR.plugins.link.fakeAnchor){var ad=F.document.getElementsByTag('img');for(Y=0,Z=ad.count();Y<Z;Y++){if(aa=CKEDITOR.plugins.link.tryRestoreFakeAnchor(F,ad.getItem(Y)))X.push({name:aa.getAttribute('name'),id:aa.getAttribute('id')});}}this._.selectedElement=G;return M;},q=function(F,G){if(G[F])this.setValue(G[F][this.id]||'');},r=function(F){return q.call(this,'target',F);},s=function(F){return q.call(this,'adv',F);},t=function(F,G){if(!G[F])G[F]={};G[F][this.id]=this.getValue()||'';},u=function(F){return t.call(this,'target',F);},v=function(F){return t.call(this,'adv',F);};function w(F){return F.replace(/\\'/g,"'");};function x(F){return F.replace(/'/g,'\\$&');};var y=a.config.emailProtection||'';if(y&&y!='encode'){var z={};y.replace(/^([^(]+)\(([^)]+)\)$/,function(F,G,H){z.name=G;z.params=[];H.replace(/[^,\s]+/g,function(I){z.params.push(I);});});}function A(F){var G,H=z.name,I=z.params,J,K;G=[H,'('];for(var L=0;L<I.length;L++){J=I[L].toLowerCase();K=F[J];L>0&&G.push(',');G.push("'",K?x(encodeURIComponent(F[J])):'',"'");}G.push(')');return G.join('');};function B(F){var G,H=F.length,I=[];for(var J=0;J<H;J++){G=F.charCodeAt(J);I.push(G);}return 'String.fromCharCode('+I.join(',')+')';};function C(F){var G=F.getAttribute('class');
//return G?G.replace(/\s*(?:cke_anchor_empty|cke_anchor)(?:\s*$)?/g,''):'';};var D=a.lang.common,E=a.lang.link;return{title:E.title,minWidth:350,minHeight:230,contents:[{id:'info',label:E.info,title:E.info,elements:[{id:'linkType',type:'select',label:E.type,'default':'url',items:[[E.toUrl,'url'],[E.toAnchor,'anchor'],[E.toEmail,'email']],onChange:d,setup:function(F){if(F.type)this.setValue(F.type);},commit:function(F){F.type=this.getValue();}},{type:'vbox',id:'urlOptions',children:[{type:'hbox',widths:['25%','75%'],children:[{id:'protocol',type:'select',label:D.protocol,'default':'http://',items:[['http://‎','http://'],['https://‎','https://'],['ftp://‎','ftp://'],['news://‎','news://'],[E.other,'']],setup:function(F){if(F.url)this.setValue(F.url.protocol||'');},commit:function(F){if(!F.url)F.url={};F.url.protocol=this.getValue();}},{type:'text',id:'url',label:D.url,required:true,onLoad:function(){this.allowOnChange=true;},onKeyUp:function(){var K=this;K.allowOnChange=false;var F=K.getDialog().getContentElement('info','protocol'),G=K.getValue(),H=/^(http|https|ftp|news):\/\/(?=.)/i,I=/^((javascript:)|[#\/\.\?])/i,J=H.exec(G);if(J){K.setValue(G.substr(J[0].length));F.setValue(J[0].toLowerCase());}else if(I.test(G))F.setValue('');K.allowOnChange=true;},onChange:function(){if(this.allowOnChange)this.onKeyUp();},validate:function(){var F=this.getDialog();if(F.getContentElement('info','linkType')&&F.getValueOf('info','linkType')!='url')return true;if(this.getDialog().fakeObj)return true;var G=CKEDITOR.dialog.validate.notEmpty(E.noUrl);return G.apply(this);},setup:function(F){this.allowOnChange=false;if(F.url)this.setValue(F.url.url);this.allowOnChange=true;},commit:function(F){this.onChange();if(!F.url)F.url={};F.url.url=this.getValue();this.allowOnChange=false;}}],setup:function(F){if(!this.getDialog().getContentElement('info','linkType'))this.getElement().show();}},{type:'button',id:'browse',hidden:'true',filebrowser:'info:url',label:D.browseServer}]},{type:'vbox',id:'anchorOptions',width:260,align:'center',padding:0,children:[{type:'fieldset',id:'selectAnchorText',label:E.selectAnchor,setup:function(F){if(F.anchors.length>0)this.getElement().show();else this.getElement().hide();},children:[{type:'hbox',id:'selectAnchor',children:[{type:'select',id:'anchorName','default':'',label:E.anchorName,style:'width: 100%;',items:[['']],setup:function(F){var I=this;I.clear();I.add('');for(var G=0;G<F.anchors.length;G++){if(F.anchors[G].name)I.add(F.anchors[G].name);}if(F.anchor)I.setValue(F.anchor.name);
//var H=I.getDialog().getContentElement('info','linkType');if(H&&H.getValue()=='email')I.focus();},commit:function(F){if(!F.anchor)F.anchor={};F.anchor.name=this.getValue();}},{type:'select',id:'anchorId','default':'',label:E.anchorId,style:'width: 100%;',items:[['']],setup:function(F){var H=this;H.clear();H.add('');for(var G=0;G<F.anchors.length;G++){if(F.anchors[G].id)H.add(F.anchors[G].id);}if(F.anchor)H.setValue(F.anchor.id);},commit:function(F){if(!F.anchor)F.anchor={};F.anchor.id=this.getValue();}}],setup:function(F){if(F.anchors.length>0)this.getElement().show();else this.getElement().hide();}}]},{type:'html',id:'noAnchors',style:'text-align: center;',html:'<div role="note" tabIndex="-1">'+CKEDITOR.tools.htmlEncode(E.noAnchors)+'</div>',focus:true,setup:function(F){if(F.anchors.length<1)this.getElement().show();else this.getElement().hide();}}],setup:function(F){if(!this.getDialog().getContentElement('info','linkType'))this.getElement().hide();}},{type:'vbox',id:'emailOptions',padding:1,children:[{type:'text',id:'emailAddress',label:E.emailAddress,required:true,validate:function(){var F=this.getDialog();if(!F.getContentElement('info','linkType')||F.getValueOf('info','linkType')!='email')return true;var G=CKEDITOR.dialog.validate.notEmpty(E.noEmail);return G.apply(this);},setup:function(F){if(F.email)this.setValue(F.email.address);var G=this.getDialog().getContentElement('info','linkType');if(G&&G.getValue()=='email')this.select();},commit:function(F){if(!F.email)F.email={};F.email.address=this.getValue();}},{type:'text',id:'emailSubject',label:E.emailSubject,setup:function(F){if(F.email)this.setValue(F.email.subject);},commit:function(F){if(!F.email)F.email={};F.email.subject=this.getValue();}},{type:'textarea',id:'emailBody',label:E.emailBody,rows:3,'default':'',setup:function(F){if(F.email)this.setValue(F.email.body);},commit:function(F){if(!F.email)F.email={};F.email.body=this.getValue();}}],setup:function(F){if(!this.getDialog().getContentElement('info','linkType'))this.getElement().hide();}}]},{id:'target',label:E.target,title:E.target,elements:[{type:'hbox',widths:['50%','50%'],children:[{type:'select',id:'linkTargetType',label:D.target,'default':'notSet',style:'width : 100%;',items:[[D.notSet,'notSet'],[E.targetFrame,'frame'],[E.targetPopup,'popup'],[D.targetNew,'_blank'],[D.targetTop,'_top'],[D.targetSelf,'_self'],[D.targetParent,'_parent']],onChange:c,setup:function(F){if(F.target)this.setValue(F.target.type||'notSet');c.call(this);},commit:function(F){if(!F.target)F.target={};
//F.target.type=this.getValue();}},{type:'text',id:'linkTargetName',label:E.targetFrameName,'default':'',setup:function(F){if(F.target)this.setValue(F.target.name);},commit:function(F){if(!F.target)F.target={};F.target.name=this.getValue().replace(/\W/gi,'');}}]},{type:'vbox',width:'100%',align:'center',padding:2,id:'popupFeatures',children:[{type:'fieldset',label:E.popupFeatures,children:[{type:'hbox',children:[{type:'checkbox',id:'resizable',label:E.popupResizable,setup:r,commit:u},{type:'checkbox',id:'status',label:E.popupStatusBar,setup:r,commit:u}]},{type:'hbox',children:[{type:'checkbox',id:'location',label:E.popupLocationBar,setup:r,commit:u},{type:'checkbox',id:'toolbar',label:E.popupToolbar,setup:r,commit:u}]},{type:'hbox',children:[{type:'checkbox',id:'menubar',label:E.popupMenuBar,setup:r,commit:u},{type:'checkbox',id:'fullscreen',label:E.popupFullScreen,setup:r,commit:u}]},{type:'hbox',children:[{type:'checkbox',id:'scrollbars',label:E.popupScrollBars,setup:r,commit:u},{type:'checkbox',id:'dependent',label:E.popupDependent,setup:r,commit:u}]},{type:'hbox',children:[{type:'text',widths:['50%','50%'],labelLayout:'horizontal',label:D.width,id:'width',setup:r,commit:u},{type:'text',labelLayout:'horizontal',widths:['50%','50%'],label:E.popupLeft,id:'left',setup:r,commit:u}]},{type:'hbox',children:[{type:'text',labelLayout:'horizontal',widths:['50%','50%'],label:D.height,id:'height',setup:r,commit:u},{type:'text',labelLayout:'horizontal',label:E.popupTop,widths:['50%','50%'],id:'top',setup:r,commit:u}]}]}]}]},{id:'upload',label:E.upload,title:E.upload,hidden:true,filebrowser:'uploadButton',elements:[{type:'file',id:'upload',label:D.upload,style:'height:40px',size:29},{type:'fileButton',id:'uploadButton',label:D.uploadSubmit,filebrowser:'info:url','for':['upload','upload']}]},{id:'advanced',label:E.advanced,title:E.advanced,elements:[{type:'vbox',padding:1,children:[{type:'hbox',widths:['45%','35%','20%'],children:[{type:'text',id:'advId',label:E.id,setup:s,commit:v},{type:'select',id:'advLangDir',label:E.langDir,'default':'',style:'width:110px',items:[[D.notSet,''],[E.langDirLTR,'ltr'],[E.langDirRTL,'rtl']],setup:s,commit:v},{type:'text',id:'advAccessKey',width:'80px',label:E.acccessKey,maxLength:1,setup:s,commit:v}]},{type:'hbox',widths:['45%','35%','20%'],children:[{type:'text',label:E.name,id:'advName',setup:s,commit:v},{type:'text',label:E.langCode,id:'advLangCode',width:'110px','default':'',setup:s,commit:v},{type:'text',label:E.tabIndex,id:'advTabIndex',width:'80px',maxLength:5,setup:s,commit:v}]}]},{type:'vbox',padding:1,children:[{type:'hbox',widths:['45%','55%'],children:[{type:'text',label:E.advisoryTitle,'default':'',id:'advTitle',setup:s,commit:v},{type:'text',label:E.advisoryContentType,'default':'',id:'advContentType',setup:s,commit:v}]},{type:'hbox',widths:['45%','55%'],children:[{type:'text',label:E.cssClasses,'default':'',id:'advCSSClasses',setup:s,commit:v},{type:'text',label:E.charset,'default':'',id:'advCharset',setup:s,commit:v}]},{type:'hbox',widths:['45%','55%'],children:[{type:'text',label:E.rel,'default':'',id:'advRel',setup:s,commit:v},{type:'text',label:E.styles,'default':'',id:'advStyles',validate:CKEDITOR.dialog.validate.inlineStyle(a.lang.common.invalidInlineStyle),setup:s,commit:v}]}]}]}],onShow:function(){var F=this.getParentEditor(),G=F.getSelection(),H=null;
//if((H=b.getSelectedLink(F))&&H.hasAttribute('href'))G.selectElement(H);else H=null;this.setupContent(p.apply(this,[F,H]));},onOk:function(){var F={},G=[],H={},I=this,J=this.getParentEditor();this.commitContent(H);switch(H.type||'url'){case 'url':var K=H.url&&H.url.protocol!=undefined?H.url.protocol:'http://',L=H.url&&CKEDITOR.tools.trim(H.url.url)||'';F['data-cke-saved-href']=L.indexOf('/')===0?L:K+L;break;case 'anchor':var M=H.anchor&&H.anchor.name,N=H.anchor&&H.anchor.id;F['data-cke-saved-href']='#'+(M||N||'');break;case 'email':var O,P=H.email,Q=P.address;switch(y){case '':case 'encode':var R=encodeURIComponent(P.subject||''),S=encodeURIComponent(P.body||''),T=[];R&&T.push('subject='+R);S&&T.push('body='+S);T=T.length?'?'+T.join('&'):'';if(y=='encode'){O=["javascript:void(location.href='mailto:'+",B(Q)];T&&O.push("+'",x(T),"'");O.push(')');}else O=['mailto:',Q,T];break;default:var U=Q.split('@',2);P.name=U[0];P.domain=U[1];O=['javascript:',A(P)];}F['data-cke-saved-href']=O.join('');break;}if(H.target)if(H.target.type=='popup'){var V=["window.open(this.href, '",H.target.name||'',"', '"],W=['resizable','status','location','toolbar','menubar','fullscreen','scrollbars','dependent'],X=W.length,Y=function(ai){if(H.target[ai])W.push(ai+'='+H.target[ai]);};for(var Z=0;Z<X;Z++)W[Z]=W[Z]+(H.target[W[Z]]?'=yes':'=no');Y('width');Y('left');Y('height');Y('top');V.push(W.join(','),"'); return false;");F['data-cke-pa-onclick']=V.join('');G.push('target');}else{if(H.target.type!='notSet'&&H.target.name)F.target=H.target.name;else G.push('target');G.push('data-cke-pa-onclick','onclick');}if(H.adv){var aa=function(ai,aj){var ak=H.adv[ai];if(ak)F[aj]=ak;else G.push(aj);};aa('advId','id');aa('advLangDir','dir');aa('advAccessKey','accessKey');if(H.adv.advName)F.name=F['data-cke-saved-name']=H.adv.advName;else G=G.concat(['data-cke-saved-name','name']);aa('advLangCode','lang');aa('advTabIndex','tabindex');aa('advTitle','title');aa('advContentType','type');aa('advCSSClasses','class');aa('advCharset','charset');aa('advStyles','style');aa('advRel','rel');}var ab=J.getSelection();F.href=F['data-cke-saved-href'];if(!this._.selectedElement){var ac=ab.getRanges(true);if(ac.length==1&&ac[0].collapsed){var ad=new CKEDITOR.dom.text(H.type=='email'?H.email.address:F['data-cke-saved-href'],J.document);ac[0].insertNode(ad);ac[0].selectNodeContents(ad);ab.selectRanges(ac);}var ae=new CKEDITOR.style({element:'a',attributes:F});ae.type=CKEDITOR.STYLE_INLINE;ae.apply(J.document);}else{var af=this._.selectedElement,ag=af.data('cke-saved-href'),ah=af.getHtml();
//af.setAttributes(F);af.removeAttributes(G);if(H.adv&&H.adv.advName&&CKEDITOR.plugins.link.synAnchorSelector)af.addClass(af.getChildCount()?'cke_anchor':'cke_anchor_empty');if(ag==ah||H.type=='email'&&ah.indexOf('@')!=-1)af.setHtml(H.type=='email'?H.email.address:F['data-cke-saved-href']);ab.selectElement(af);delete this._.selectedElement;}},onLoad:function(){if(!a.config.linkShowAdvancedTab)this.hidePage('advanced');if(!a.config.linkShowTargetTab)this.hidePage('target');},onFocus:function(){var F=this.getContentElement('info','linkType'),G;if(F&&F.getValue()=='url'){G=this.getContentElement('info','url');G.select();}}};});


/*
Copyright (c) 2003-2012, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.dialog.add( 'link', function( editor )
{
	var plugin = CKEDITOR.plugins.link;
	// Handles the event when the "Target" selection box is changed.
	var targetChanged = function()
	{
		var dialog = this.getDialog(),
			popupFeatures = dialog.getContentElement( 'target', 'popupFeatures' ),
			targetName = dialog.getContentElement( 'target', 'linkTargetName' ),
			value = this.getValue();

		if ( !popupFeatures || !targetName )
			return;

		popupFeatures = popupFeatures.getElement();
		popupFeatures.hide();
		targetName.setValue( '' );

		switch ( value )
 		{
			case 'frame' :
				targetName.setLabel( editor.lang.link.targetFrameName );
				targetName.getElement().show();
				break;
			case 'popup' :
				popupFeatures.show();
				targetName.setLabel( editor.lang.link.targetPopupName );
				targetName.getElement().show();
				break;
			default :
				targetName.setValue( value );
				targetName.getElement().hide();
				break;
 		}

	};

	// Handles the event when the "Type" selection box is changed.
	var linkTypeChanged = function()
	{
		var dialog = this.getDialog(),
			partIds = [ 'urlOptions', 'anchorOptions', 'emailOptions' ],
			typeValue = this.getValue(),
			uploadTab = dialog.definition.getContents( 'upload' ),
			uploadInitiallyHidden = uploadTab && uploadTab.hidden;

		if ( typeValue == 'url' )
		{
			if ( editor.config.linkShowTargetTab )
				dialog.showPage( 'target' );
			if ( !uploadInitiallyHidden )
				dialog.showPage( 'upload' );
		}
		else
		{
			dialog.hidePage( 'target' );
			if ( !uploadInitiallyHidden )
				dialog.hidePage( 'upload' );
		}

		for ( var i = 0 ; i < partIds.length ; i++ )
		{
			var element = dialog.getContentElement( 'info', partIds[i] );
			if ( !element )
				continue;

			element = element.getElement().getParent().getParent();
			if ( partIds[i] == typeValue + 'Options' )
				element.show();
			else
				element.hide();
		}

		dialog.layout();
	};

	// Loads the parameters in a selected link to the link dialog fields.
	var javascriptProtocolRegex = /^javascript:/,
		emailRegex = /^mailto:([^?]+)(?:\?(.+))?$/,
		emailSubjectRegex = /subject=([^;?:@&=$,\/]*)/,
		emailBodyRegex = /body=([^;?:@&=$,\/]*)/,
		anchorRegex = /^#(.*)$/,
		urlRegex = /^((?:http|https|ftp|news):\/\/)?(.*)$/,
		selectableTargets = /^(_(?:self|top|parent|blank))$/,
		encodedEmailLinkRegex = /^javascript:void\(location\.href='mailto:'\+String\.fromCharCode\(([^)]+)\)(?:\+'(.*)')?\)$/,
		functionCallProtectedEmailLinkRegex = /^javascript:([^(]+)\(([^)]+)\)$/;

	var popupRegex =
		/\s*window.open\(\s*this\.href\s*,\s*(?:'([^']*)'|null)\s*,\s*'([^']*)'\s*\)\s*;\s*return\s*false;*\s*/;
	var popupFeaturesRegex = /(?:^|,)([^=]+)=(\d+|yes|no)/gi;

	var parseLink = function( editor, element )
	{
		var href = ( element  && ( element.data( 'cke-saved-href' ) || element.getAttribute( 'href' ) ) ) || '',
		 	javascriptMatch,
			emailMatch,
			anchorMatch,
			urlMatch,
			retval = {};

		if ( ( javascriptMatch = href.match( javascriptProtocolRegex ) ) )
		{
			if ( emailProtection == 'encode' )
			{
				href = href.replace( encodedEmailLinkRegex,
						function ( match, protectedAddress, rest )
						{
							return 'mailto:' +
							       String.fromCharCode.apply( String, protectedAddress.split( ',' ) ) +
							       ( rest && unescapeSingleQuote( rest ) );
						});
			}
			// Protected email link as function call.
			else if ( emailProtection )
			{
				href.replace( functionCallProtectedEmailLinkRegex, function( match, funcName, funcArgs )
				{
					if ( funcName == compiledProtectionFunction.name )
					{
						retval.type = 'email';
						var email = retval.email = {};

						var paramRegex = /[^,\s]+/g,
							paramQuoteRegex = /(^')|('$)/g,
							paramsMatch = funcArgs.match( paramRegex ),
							paramsMatchLength = paramsMatch.length,
							paramName,
							paramVal;

						for ( var i = 0; i < paramsMatchLength; i++ )
						{
							paramVal = decodeURIComponent( unescapeSingleQuote( paramsMatch[ i ].replace( paramQuoteRegex, '' ) ) );
							paramName = compiledProtectionFunction.params[ i ].toLowerCase();
							email[ paramName ] = paramVal;
						}
						email.address = [ email.name, email.domain ].join( '@' );
					}
				} );
			}
		}

		if ( !retval.type )
		{
			if ( ( anchorMatch = href.match( anchorRegex ) ) )
			{
				retval.type = 'anchor';
				retval.anchor = {};
				retval.anchor.name = retval.anchor.id = anchorMatch[1];
			}
			// Protected email link as encoded string.
			else if ( ( emailMatch = href.match( emailRegex ) ) )
			{
				var subjectMatch = href.match( emailSubjectRegex ),
					bodyMatch = href.match( emailBodyRegex );

				retval.type = 'email';
				var email = ( retval.email = {} );
				email.address = emailMatch[ 1 ];
				subjectMatch && ( email.subject = decodeURIComponent( subjectMatch[ 1 ] ) );
				bodyMatch && ( email.body = decodeURIComponent( bodyMatch[ 1 ] ) );
			}
			// urlRegex matches empty strings, so need to check for href as well.
			else if (  href && ( urlMatch = href.match( urlRegex ) ) )
			{
				retval.type = 'url';
				retval.url = {};
				retval.url.protocol = urlMatch[1];
				retval.url.url = urlMatch[2];
			}
			else
				retval.type = 'url';
		}

		// Load target and popup settings.
		if ( element )
		{
			var target = element.getAttribute( 'target' );
			retval.target = {};
			retval.adv = {};

			// IE BUG: target attribute is an empty string instead of null in IE if it's not set.
			if ( !target )
			{
				var onclick = element.data( 'cke-pa-onclick' ) || element.getAttribute( 'onclick' ),
					onclickMatch = onclick && onclick.match( popupRegex );
				if ( onclickMatch )
				{
					retval.target.type = 'popup';
					retval.target.name = onclickMatch[1];

					var featureMatch;
					while ( ( featureMatch = popupFeaturesRegex.exec( onclickMatch[2] ) ) )
					{
						// Some values should remain numbers (#7300)
						if ( ( featureMatch[2] == 'yes' || featureMatch[2] == '1' ) && !( featureMatch[1] in { height:1, width:1, top:1, left:1 } ) )
							retval.target[ featureMatch[1] ] = true;
						else if ( isFinite( featureMatch[2] ) )
							retval.target[ featureMatch[1] ] = featureMatch[2];
					}
				}
			}
			else
			{
				var targetMatch = target.match( selectableTargets );
				if ( targetMatch )
					retval.target.type = retval.target.name = target;
				else
				{
					retval.target.type = 'frame';
					retval.target.name = target;
				}
			}

			var me = this;
			var advAttr = function( inputName, attrName )
			{
				var value = element.getAttribute( attrName );
				if ( value !== null )
					retval.adv[ inputName ] = value || '';
			};
			advAttr( 'advId', 'id' );
			advAttr( 'advLangDir', 'dir' );
			advAttr( 'advAccessKey', 'accessKey' );

			retval.adv.advName =
				element.data( 'cke-saved-name' )
				|| element.getAttribute( 'name' )
				|| '';
			advAttr( 'advLangCode', 'lang' );
			advAttr( 'advTabIndex', 'tabindex' );
			advAttr( 'advTitle', 'title' );
			advAttr( 'advContentType', 'type' );
			CKEDITOR.plugins.link.synAnchorSelector ?
				retval.adv.advCSSClasses = getLinkClass( element )
				: advAttr( 'advCSSClasses', 'class' );
			advAttr( 'advCharset', 'charset' );
			advAttr( 'advStyles', 'style' );
			advAttr( 'advRel', 'rel' );
		}

		// Find out whether we have any anchors in the editor.
		var anchors = retval.anchors = [],
			i, count, item;

		// For some browsers we set contenteditable="false" on anchors, making document.anchors not to include them, so we must traverse the links manually (#7893).
		if ( CKEDITOR.plugins.link.emptyAnchorFix )
		{
			var links = editor.document.getElementsByTag( 'a' );
			for ( i = 0, count = links.count(); i < count; i++ )
			{
				item = links.getItem( i );
				if ( item.data( 'cke-saved-name' ) || item.hasAttribute( 'name' ) )
					anchors.push( { name : item.data( 'cke-saved-name' ) || item.getAttribute( 'name' ), id : item.getAttribute( 'id' ) } );
			}
		}
		else
		{
			var anchorList = new CKEDITOR.dom.nodeList( editor.document.$.anchors );
			for ( i = 0, count = anchorList.count(); i < count; i++ )
			{
				item = anchorList.getItem( i );
				anchors[ i ] = { name : item.getAttribute( 'name' ), id : item.getAttribute( 'id' ) };
			}
		}

		if ( CKEDITOR.plugins.link.fakeAnchor )
		{
			var imgs = editor.document.getElementsByTag( 'img' );
			for ( i = 0, count = imgs.count(); i < count; i++ )
			{
				if ( ( item = CKEDITOR.plugins.link.tryRestoreFakeAnchor( editor, imgs.getItem( i ) ) ) )
					anchors.push( { name : item.getAttribute( 'name' ), id : item.getAttribute( 'id' ) } );
			}
		}

		// Record down the selected element in the dialog.
		this._.selectedElement = element;
		return retval;
	};

	var setupParams = function( page, data )
	{
		if ( data[page] )
			this.setValue( data[page][this.id] || '' );
	};

	var setupPopupParams = function( data )
	{
		return setupParams.call( this, 'target', data );
	};

	var setupAdvParams = function( data )
	{
		return setupParams.call( this, 'adv', data );
	};

	var commitParams = function( page, data )
	{
		if ( !data[page] )
			data[page] = {};

		data[page][this.id] = this.getValue() || '';
	};

	var commitPopupParams = function( data )
	{
		return commitParams.call( this, 'target', data );
	};

	var commitAdvParams = function( data )
	{
		return commitParams.call( this, 'adv', data );
	};

	function unescapeSingleQuote( str )
	{
		return str.replace( /\\'/g, '\'' );
	}

	function escapeSingleQuote( str )
	{
		return str.replace( /'/g, '\\$&' );
	}

	var emailProtection = editor.config.emailProtection || '';

	// Compile the protection function pattern.
	if ( emailProtection && emailProtection != 'encode' )
	{
		var compiledProtectionFunction = {};

		emailProtection.replace( /^([^(]+)\(([^)]+)\)$/, function( match, funcName, params )
		{
			compiledProtectionFunction.name = funcName;
			compiledProtectionFunction.params = [];
			params.replace( /[^,\s]+/g, function( param )
			{
				compiledProtectionFunction.params.push( param );
			} );
		} );
	}

	function protectEmailLinkAsFunction( email )
	{
		var retval,
			name = compiledProtectionFunction.name,
			params = compiledProtectionFunction.params,
			paramName,
			paramValue;

		retval = [ name, '(' ];
		for ( var i = 0; i < params.length; i++ )
		{
			paramName = params[ i ].toLowerCase();
			paramValue = email[ paramName ];

			i > 0 && retval.push( ',' );
			retval.push( '\'',
						 paramValue ?
						 escapeSingleQuote( encodeURIComponent( email[ paramName ] ) )
						 : '',
						 '\'');
		}
		retval.push( ')' );
		return retval.join( '' );
	}

	function protectEmailAddressAsEncodedString( address )
	{
		var charCode,
			length = address.length,
			encodedChars = [];
		for ( var i = 0; i < length; i++ )
		{
			charCode = address.charCodeAt( i );
			encodedChars.push( charCode );
		}
		return 'String.fromCharCode(' + encodedChars.join( ',' ) + ')';
	}

	function getLinkClass( ele )
	{
		var className = ele.getAttribute( 'class' );
		return className ? className.replace( /\s*(?:cke_anchor_empty|cke_anchor)(?:\s*$)?/g, '' ) : '';
	}

	var commonLang = editor.lang.common,
		linkLang = editor.lang.link;

	return {
		title : linkLang.title,
		minWidth : 350,
		minHeight : 230,
		contents : [
			{
				id : 'info',
				label : linkLang.info,
				title : linkLang.info,
				elements :
				[
					{
						id : 'linkType',
						type : 'select',
						label : linkLang.type,
						'default' : 'url',
						items :
						[
							[ linkLang.toUrl, 'url' ],
							[ linkLang.toAnchor, 'anchor' ],
							[ linkLang.toEmail, 'email' ]
						],
						onChange : linkTypeChanged,
						setup : function( data )
						{
							if ( data.type )
								this.setValue( data.type );
						},
						commit : function( data )
						{
							data.type = this.getValue();
						}
					},
					{
						type : 'vbox',
						id : 'urlOptions',
						children :
						[
							{
								type : 'hbox',
								widths : [ '25%', '75%' ],
								children :
								[
									{
										id : 'protocol',
										type : 'select',
										label : commonLang.protocol,
										'default' : 'http://',
										items :
										[
											// Force 'ltr' for protocol names in BIDI. (#5433)
											[ 'http://\u200E', 'http://' ],
											[ 'https://\u200E', 'https://' ],
											[ 'ftp://\u200E', 'ftp://' ],
											[ 'news://\u200E', 'news://' ],
											[ linkLang.other , '' ]
										],
										setup : function( data )
										{
											if ( data.url )
												this.setValue( data.url.protocol || '' );
										},
										commit : function( data )
										{
											if ( !data.url )
												data.url = {};

											data.url.protocol = this.getValue();
										}
									},
									{
										type : 'text',
										id : 'url',
										label : commonLang.url,
										required: true,
										onLoad : function ()
										{
											this.allowOnChange = true;
										},
										onKeyUp : function()
										{
											this.allowOnChange = false;
											var	protocolCmb = this.getDialog().getContentElement( 'info', 'protocol' ),
												url = this.getValue(),
												urlOnChangeProtocol = /^(http|https|ftp|news):\/\/(?=.)/i,
												urlOnChangeTestOther = /^((javascript:)|[#\/\.\?])/i;

											var protocol = urlOnChangeProtocol.exec( url );
											if ( protocol )
											{
												this.setValue( url.substr( protocol[ 0 ].length ) );
												protocolCmb.setValue( protocol[ 0 ].toLowerCase() );
											}
											else if ( urlOnChangeTestOther.test( url ) )
												protocolCmb.setValue( '' );

											this.allowOnChange = true;
										},
										onChange : function()
										{
											if ( this.allowOnChange )		// Dont't call on dialog load.
												this.onKeyUp();
										},
										validate : function()
										{
											var dialog = this.getDialog();

											if ( dialog.getContentElement( 'info', 'linkType' ) &&
													dialog.getValueOf( 'info', 'linkType' ) != 'url' )
												return true;

											if ( this.getDialog().fakeObj )	// Edit Anchor.
												return true;

											var func = CKEDITOR.dialog.validate.notEmpty( linkLang.noUrl );
											return func.apply( this );
										},
										setup : function( data )
										{
											this.allowOnChange = false;
											if ( data.url )
												this.setValue( data.url.url );
											this.allowOnChange = true;

										},
										commit : function( data )
										{
											// IE will not trigger the onChange event if the mouse has been used
											// to carry all the operations #4724
											this.onChange();

											if ( !data.url )
												data.url = {};

											data.url.url = this.getValue();
											this.allowOnChange = false;
										}
									}
								],
								setup : function( data )
								{
									if ( !this.getDialog().getContentElement( 'info', 'linkType' ) )
										this.getElement().show();
								}
							},
							{
								type : 'button',
								id : 'browse',
								hidden : 'true',
								filebrowser : 'info:url',
								label : commonLang.browseServer
							}
						]
					},
					{
						type : 'vbox',
						id : 'anchorOptions',
						width : 260,
						align : 'center',
						padding : 0,
						children :
						[
							{
								type : 'fieldset',
								id : 'selectAnchorText',
								label : linkLang.selectAnchor,
								setup : function( data )
								{
									if ( data.anchors.length > 0 )
										this.getElement().show();
									else
										this.getElement().hide();
								},
								children :
								[
									{
										type : 'hbox',
										id : 'selectAnchor',
										children :
										[
											{
												type : 'select',
												id : 'anchorName',
												'default' : '',
												label : linkLang.anchorName,
												style : 'width: 100%;',
												items :
												[
													[ '' ]
												],
												setup : function( data )
												{
													this.clear();
													this.add( '' );
													for ( var i = 0 ; i < data.anchors.length ; i++ )
													{
														if ( data.anchors[i].name )
															this.add( data.anchors[i].name );
													}

													if ( data.anchor )
														this.setValue( data.anchor.name );

													var linkType = this.getDialog().getContentElement( 'info', 'linkType' );
													if ( linkType && linkType.getValue() == 'email' )
														this.focus();
												},
												commit : function( data )
												{
													if ( !data.anchor )
														data.anchor = {};

													data.anchor.name = this.getValue();
												}
											},
											{
												type : 'select',
												id : 'anchorId',
												'default' : '',
												label : linkLang.anchorId,
												style : 'width: 100%;',
												items :
												[
													[ '' ]
												],
												setup : function( data )
												{
													this.clear();
													this.add( '' );
													for ( var i = 0 ; i < data.anchors.length ; i++ )
													{
														if ( data.anchors[i].id )
															this.add( data.anchors[i].id );
													}

													if ( data.anchor )
														this.setValue( data.anchor.id );
												},
												commit : function( data )
												{
													if ( !data.anchor )
														data.anchor = {};

													data.anchor.id = this.getValue();
												}
											}
										],
										setup : function( data )
										{
											if ( data.anchors.length > 0 )
												this.getElement().show();
											else
												this.getElement().hide();
										}
									}
								]
							},
							{
								type : 'html',
								id : 'noAnchors',
								style : 'text-align: center;',
								html : '<div role="note" tabIndex="-1">' + CKEDITOR.tools.htmlEncode( linkLang.noAnchors ) + '</div>',
								// Focus the first element defined in above html.
								focus : true,
								setup : function( data )
								{
									if ( data.anchors.length < 1 )
										this.getElement().show();
									else
										this.getElement().hide();
								}
							}
						],
						setup : function( data )
						{
							if ( !this.getDialog().getContentElement( 'info', 'linkType' ) )
								this.getElement().hide();
						}
					},
					{
						type :  'vbox',
						id : 'emailOptions',
						padding : 1,
						children :
						[
							{
								type : 'text',
								id : 'emailAddress',
								label : linkLang.emailAddress,
								required : true,
								validate : function()
								{
									var dialog = this.getDialog();

									if ( !dialog.getContentElement( 'info', 'linkType' ) ||
											dialog.getValueOf( 'info', 'linkType' ) != 'email' )
										return true;

									var func = CKEDITOR.dialog.validate.notEmpty( linkLang.noEmail );
									return func.apply( this );
								},
								setup : function( data )
								{
									if ( data.email )
										this.setValue( data.email.address );

									var linkType = this.getDialog().getContentElement( 'info', 'linkType' );
									if ( linkType && linkType.getValue() == 'email' )
										this.select();
								},
								commit : function( data )
								{
									if ( !data.email )
										data.email = {};

									data.email.address = this.getValue();
								}
							},
							{
								type : 'text',
								id : 'emailSubject',
								label : linkLang.emailSubject,
								setup : function( data )
								{
									if ( data.email )
										this.setValue( data.email.subject );
								},
								commit : function( data )
								{
									if ( !data.email )
										data.email = {};

									data.email.subject = this.getValue();
								}
							},
							{
								type : 'textarea',
								id : 'emailBody',
								label : linkLang.emailBody,
								rows : 3,
								'default' : '',
								setup : function( data )
								{
									if ( data.email )
										this.setValue( data.email.body );
								},
								commit : function( data )
								{
									if ( !data.email )
										data.email = {};

									data.email.body = this.getValue();
								}
							}
						],
						setup : function( data )
						{
							if ( !this.getDialog().getContentElement( 'info', 'linkType' ) )
								this.getElement().hide();
						}
					}
				]
			},
			{
				id : 'target',
				label : linkLang.target,
				title : linkLang.target,
				elements :
				[
					{
						type : 'hbox',
						widths : [ '50%', '50%' ],
						children :
						[
							{
								type : 'select',
								id : 'linkTargetType',
								label : commonLang.target,
								'default' : 'notSet',
								style : 'width : 100%;',
								'items' :
								[
									[ commonLang.notSet, 'notSet' ],
									[ linkLang.targetFrame, 'frame' ],
									[ linkLang.targetPopup, 'popup' ],
									[ commonLang.targetNew, '_blank' ],
									[ commonLang.targetTop, '_top' ],
									[ commonLang.targetSelf, '_self' ],
									[ commonLang.targetParent, '_parent' ]
								],
								onChange : targetChanged,
								setup : function( data )
								{
									if ( data.target )
										this.setValue( data.target.type || 'notSet' );
									targetChanged.call( this );
								},
								commit : function( data )
								{
									if ( !data.target )
										data.target = {};

									data.target.type = this.getValue();
								}
							},
							{
								type : 'text',
								id : 'linkTargetName',
								label : linkLang.targetFrameName,
								'default' : '',
								setup : function( data )
								{
									if ( data.target )
										this.setValue( data.target.name );
								},
								commit : function( data )
								{
									if ( !data.target )
										data.target = {};

									data.target.name = this.getValue().replace(/\W/gi, '');
								}
							}
						]
					},
					{
						type : 'vbox',
						width : '100%',
						align : 'center',
						padding : 2,
						id : 'popupFeatures',
						children :
						[
							{
								type : 'fieldset',
								label : linkLang.popupFeatures,
								children :
								[
									{
										type : 'hbox',
										children :
										[
											{
												type : 'checkbox',
												id : 'resizable',
												label : linkLang.popupResizable,
												setup : setupPopupParams,
												commit : commitPopupParams
											},
											{
												type : 'checkbox',
												id : 'status',
												label : linkLang.popupStatusBar,
												setup : setupPopupParams,
												commit : commitPopupParams

											}
										]
									},
									{
										type : 'hbox',
										children :
										[
											{
												type : 'checkbox',
												id : 'location',
												label : linkLang.popupLocationBar,
												setup : setupPopupParams,
												commit : commitPopupParams

											},
											{
												type : 'checkbox',
												id : 'toolbar',
												label : linkLang.popupToolbar,
												setup : setupPopupParams,
												commit : commitPopupParams

											}
										]
									},
									{
										type : 'hbox',
										children :
										[
											{
												type : 'checkbox',
												id : 'menubar',
												label : linkLang.popupMenuBar,
												setup : setupPopupParams,
												commit : commitPopupParams

											},
											{
												type : 'checkbox',
												id : 'fullscreen',
												label : linkLang.popupFullScreen,
												setup : setupPopupParams,
												commit : commitPopupParams

											}
										]
									},
									{
										type : 'hbox',
										children :
										[
											{
												type : 'checkbox',
												id : 'scrollbars',
												label : linkLang.popupScrollBars,
												setup : setupPopupParams,
												commit : commitPopupParams

											},
											{
												type : 'checkbox',
												id : 'dependent',
												label : linkLang.popupDependent,
												setup : setupPopupParams,
												commit : commitPopupParams

											}
										]
									},
									{
										type : 'hbox',
										children :
										[
											{
												type :  'text',
												widths : [ '50%', '50%' ],
												labelLayout : 'horizontal',
												label : commonLang.width,
												id : 'width',
												setup : setupPopupParams,
												commit : commitPopupParams

											},
											{
												type :  'text',
												labelLayout : 'horizontal',
												widths : [ '50%', '50%' ],
												label : linkLang.popupLeft,
												id : 'left',
												setup : setupPopupParams,
												commit : commitPopupParams

											}
										]
									},
									{
										type : 'hbox',
										children :
										[
											{
												type :  'text',
												labelLayout : 'horizontal',
												widths : [ '50%', '50%' ],
												label : commonLang.height,
												id : 'height',
												setup : setupPopupParams,
												commit : commitPopupParams

											},
											{
												type :  'text',
												labelLayout : 'horizontal',
												label : linkLang.popupTop,
												widths : [ '50%', '50%' ],
												id : 'top',
												setup : setupPopupParams,
												commit : commitPopupParams

											}
										]
									}
								]
							}
						]
					}
				]
			},
			{
				id : 'upload',
				label : linkLang.upload,
				title : linkLang.upload,
				hidden : true,
				filebrowser : 'uploadButton',
				elements :
				[
					{
						type : 'file',
						id : 'upload',
						label : commonLang.upload,
						style: 'height:40px',
						size : 29
					},
					{
						type : 'fileButton',
						id : 'uploadButton',
						label : commonLang.uploadSubmit,
						filebrowser : 'info:url',
						'for' : [ 'upload', 'upload' ]
					}
				]
			},
			{
				id : 'advanced',
				label : linkLang.advanced,
				title : linkLang.advanced,
				elements :
				[
					{
						type : 'vbox',
						padding : 1,
						children :
						[
							{
								type : 'hbox',
								widths : [ '45%', '35%', '20%' ],
								children :
								[
									{
										type : 'text',
										id : 'advId',
										label : linkLang.id,
										setup : setupAdvParams,
										commit : commitAdvParams
									},
									{
										type : 'select',
										id : 'advLangDir',
										label : linkLang.langDir,
										'default' : '',
										style : 'width:110px',
										items :
										[
											[ commonLang.notSet, '' ],
											[ linkLang.langDirLTR, 'ltr' ],
											[ linkLang.langDirRTL, 'rtl' ]
										],
										setup : setupAdvParams,
										commit : commitAdvParams
									},
									{
										type : 'text',
										id : 'advAccessKey',
										width : '80px',
										label : linkLang.acccessKey,
										maxLength : 1,
										setup : setupAdvParams,
										commit : commitAdvParams

									}
								]
							},
							{
								type : 'hbox',
								widths : [ '45%', '35%', '20%' ],
								children :
								[
									{
										type : 'text',
										label : linkLang.name,
										id : 'advName',
										setup : setupAdvParams,
										commit : commitAdvParams

									},
									{
										type : 'text',
										label : linkLang.langCode,
										id : 'advLangCode',
										width : '110px',
										'default' : '',
										setup : setupAdvParams,
										commit : commitAdvParams

									},
									{
										type : 'text',
										label : linkLang.tabIndex,
										id : 'advTabIndex',
										width : '80px',
										maxLength : 5,
										setup : setupAdvParams,
										commit : commitAdvParams

									}
								]
							}
						]
					},
					{
						type : 'vbox',
						padding : 1,
						children :
						[
							{
								type : 'hbox',
								widths : [ '45%', '55%' ],
								children :
								[
									{
										type : 'text',
										label : linkLang.advisoryTitle,
										'default' : '',
										id : 'advTitle',
										setup : setupAdvParams,
										commit : commitAdvParams

									},
									{
										type : 'text',
										label : linkLang.advisoryContentType,
										'default' : '',
										id : 'advContentType',
										setup : setupAdvParams,
										commit : commitAdvParams

									}
								]
							},
							{
								type : 'hbox',
								widths : [ '45%', '55%' ],
								children :
								[
									{
										type : 'text',
										label : linkLang.cssClasses,
										'default' : '',
										id : 'advCSSClasses',
										setup : setupAdvParams,
										commit : commitAdvParams

									},
									{
										type : 'text',
										label : linkLang.charset,
										'default' : '',
										id : 'advCharset',
										setup : setupAdvParams,
										commit : commitAdvParams

									}
								]
							},
							{
								type : 'hbox',
								widths : [ '45%', '55%' ],
								children :
								[
									{
										type : 'text',
										label : linkLang.rel,
										'default' : '',
										id : 'advRel',
										setup : setupAdvParams,
										commit : commitAdvParams
									},
									{
										type : 'text',
										label : linkLang.styles,
										'default' : '',
										id : 'advStyles',
										validate : CKEDITOR.dialog.validate.inlineStyle( editor.lang.common.invalidInlineStyle ),
										setup : setupAdvParams,
										commit : commitAdvParams
									}
								]
							}
						]
					}
				]
			}
		],
		onShow : function()
		{
			var editor = this.getParentEditor(),
				selection = editor.getSelection(),
				element = null;

			// Fill in all the relevant fields if there's already one link selected.
			if ( ( element = plugin.getSelectedLink( editor ) ) && element.hasAttribute( 'href' ) )
				selection.selectElement( element );
			else
				element = null;

			this.setupContent( parseLink.apply( this, [ editor, element ] ) );
		},
		onOk : function()
		{
			var attributes = {},
				removeAttributes = [],
				data = {},
				me = this,
				editor = this.getParentEditor();

			this.commitContent( data );

			// Compose the URL.
			switch ( data.type || 'url' )
			{
				case 'url':
					var protocol = ( data.url && data.url.protocol != undefined ) ? data.url.protocol : 'http://',
						url = ( data.url && CKEDITOR.tools.trim( data.url.url ) ) || '';
					attributes[ 'data-cke-saved-href' ] = ( url.indexOf( '/' ) === 0 ) ? url : protocol + url;
					break;
				case 'anchor':
					var name = ( data.anchor && data.anchor.name ),
						id = ( data.anchor && data.anchor.id );
					attributes[ 'data-cke-saved-href' ] = '#' + ( name || id || '' );
					break;
				case 'email':

					var linkHref,
						email = data.email,
						address = email.address;

					switch( emailProtection )
					{
						case '' :
						case 'encode' :
						{
							var subject = encodeURIComponent( email.subject || '' ),
								body = encodeURIComponent( email.body || '' );

							// Build the e-mail parameters first.
							var argList = [];
							subject && argList.push( 'subject=' + subject );
							body && argList.push( 'body=' + body );
							argList = argList.length ? '?' + argList.join( '&' ) : '';

							if ( emailProtection == 'encode' )
							{
								linkHref = [ 'javascript:void(location.href=\'mailto:\'+',
											 protectEmailAddressAsEncodedString( address ) ];
								// parameters are optional.
								argList && linkHref.push( '+\'', escapeSingleQuote( argList ), '\'' );

								linkHref.push( ')' );
							}
							else
								linkHref = [ 'mailto:', address, argList ];

							break;
						}
						default :
						{
							// Separating name and domain.
							var nameAndDomain = address.split( '@', 2 );
							email.name = nameAndDomain[ 0 ];
							email.domain = nameAndDomain[ 1 ];

							linkHref = [ 'javascript:', protectEmailLinkAsFunction( email ) ];
						}
					}

					attributes[ 'data-cke-saved-href' ] = linkHref.join( '' );
					break;
			}

			// Popups and target.
			if ( data.target )
			{
				if ( data.target.type == 'popup' )
				{
					var onclickList = [ 'window.open(this.href, \'',
							data.target.name || '', '\', \'' ];
					var featureList = [ 'resizable', 'status', 'location', 'toolbar', 'menubar', 'fullscreen',
							'scrollbars', 'dependent' ];
					var featureLength = featureList.length;
					var addFeature = function( featureName )
					{
						if ( data.target[ featureName ] )
							featureList.push( featureName + '=' + data.target[ featureName ] );
					};

					for ( var i = 0 ; i < featureLength ; i++ )
						featureList[i] = featureList[i] + ( data.target[ featureList[i] ] ? '=yes' : '=no' ) ;
					addFeature( 'width' );
					addFeature( 'left' );
					addFeature( 'height' );
					addFeature( 'top' );

					onclickList.push( featureList.join( ',' ), '\'); return false;' );
					attributes[ 'data-cke-pa-onclick' ] = onclickList.join( '' );

					// Add the "target" attribute. (#5074)
					removeAttributes.push( 'target' );
				}
				else
				{
					if ( data.target.type != 'notSet' && data.target.name )
						attributes.target = data.target.name;
					else
						removeAttributes.push( 'target' );

					removeAttributes.push( 'data-cke-pa-onclick', 'onclick' );
				}
			}

			// Advanced attributes.
			if ( data.adv )
			{
				var advAttr = function( inputName, attrName )
				{
					var value = data.adv[ inputName ];
					if ( value )
						attributes[attrName] = value;
					else
						removeAttributes.push( attrName );
				};

				advAttr( 'advId', 'id' );
				advAttr( 'advLangDir', 'dir' );
				advAttr( 'advAccessKey', 'accessKey' );

				if ( data.adv[ 'advName' ] )
					attributes[ 'name' ] = attributes[ 'data-cke-saved-name' ] = data.adv[ 'advName' ];
				else
					removeAttributes = removeAttributes.concat( [ 'data-cke-saved-name', 'name' ] );

				advAttr( 'advLangCode', 'lang' );
				advAttr( 'advTabIndex', 'tabindex' );
				advAttr( 'advTitle', 'title' );
				advAttr( 'advContentType', 'type' );
				advAttr( 'advCSSClasses', 'class' );
				advAttr( 'advCharset', 'charset' );
				advAttr( 'advStyles', 'style' );
				advAttr( 'advRel', 'rel' );
			}


			var selection = editor.getSelection();

			// Browser need the "href" fro copy/paste link to work. (#6641)
			attributes.href = attributes[ 'data-cke-saved-href' ];

			if ( !this._.selectedElement )
			{
				// Create element if current selection is collapsed.
				var ranges = selection.getRanges( true );
				if ( ranges.length == 1 && ranges[0].collapsed )
				{
					// Short mailto link text view (#5736).
					var text = new CKEDITOR.dom.text( data.type == 'email' ?
							data.email.address : attributes[ 'data-cke-saved-href' ], editor.document );
					ranges[0].insertNode( text );
					ranges[0].selectNodeContents( text );
					selection.selectRanges( ranges );
				}

				// Apply style.
				var style = new CKEDITOR.style( { element : 'a', attributes : attributes } );
				style.type = CKEDITOR.STYLE_INLINE;		// need to override... dunno why.
				style.apply( editor.document );
			}
			else
			{
				// We're only editing an existing link, so just overwrite the attributes.
				var element = this._.selectedElement,
					href = element.data( 'cke-saved-href' ),
					textView = element.getHtml();

				element.setAttributes( attributes );
				element.removeAttributes( removeAttributes );

				if ( data.adv && data.adv.advName && CKEDITOR.plugins.link.synAnchorSelector )
					element.addClass( element.getChildCount() ? 'cke_anchor' : 'cke_anchor_empty' );

				// Update text view when user changes protocol (#4612).
				if ( href == textView || data.type == 'email' && textView.indexOf( '@' ) != -1 )
				{
					// Short mailto link text view (#5736).
					element.setHtml( data.type == 'email' ?
						data.email.address : attributes[ 'data-cke-saved-href' ] );
				}

				selection.selectElement( element );
				delete this._.selectedElement;
			}
		},
		onLoad : function()
		{
			if ( !editor.config.linkShowAdvancedTab )
				this.hidePage( 'advanced' );		//Hide Advanded tab.

			if ( !editor.config.linkShowTargetTab )
				this.hidePage( 'target' );		//Hide Target tab.

		},
		// Inital focus on 'url' field if link is of type URL.
		onFocus : function()
		{
			var linkType = this.getContentElement( 'info', 'linkType' ),
					urlField;
			if ( linkType && linkType.getValue() == 'url' )
			{
				urlField = this.getContentElement( 'info', 'url' );
				urlField.select();
			}
			
			//Modifica - Giacomo
			//alert(this.getContentElement( 'info', 'browse' ).filebrowser.url);
			this.getContentElement( 'info', 'browse' ).filebrowser.url += this.getContentElement( 'info', 'url' )._.inputId;
			//alert(this.getContentElement( 'info', 'browse' ).filebrowser.url);
		}
	};
});

/**
 * The e-mail address anti-spam protection option. The protection will be
 * applied when creating or modifying e-mail links through the editor interface.<br>
 * Two methods of protection can be choosed:
 * <ol>	<li>The e-mail parts (name, domain and any other query string) are
 *			assembled into a function call pattern. Such function must be
 *			provided by the developer in the pages that will use the contents.
 *		<li>Only the e-mail address is obfuscated into a special string that
 *			has no meaning for humans or spam bots, but which is properly
 *			rendered and accepted by the browser.</li></ol>
 * Both approaches require JavaScript to be enabled.
 * @name CKEDITOR.config.emailProtection
 * @since 3.1
 * @type String
 * @default '' (empty string = disabled)
 * @example
 * // href="mailto:tester@ckeditor.com?subject=subject&body=body"
 * config.emailProtection = '';
 * @example
 * // href="<a href=\"javascript:void(location.href=\'mailto:\'+String.fromCharCode(116,101,115,116,101,114,64,99,107,101,100,105,116,111,114,46,99,111,109)+\'?subject=subject&body=body\')\">e-mail</a>"
 * config.emailProtection = 'encode';
 * @example
 * // href="javascript:mt('tester','ckeditor.com','subject','body')"
 * config.emailProtection = 'mt(NAME,DOMAIN,SUBJECT,BODY)';
 */
