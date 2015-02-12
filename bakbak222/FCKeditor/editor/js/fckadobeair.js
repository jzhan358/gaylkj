/*
 * FCKeditor - The text editor for Internet - http://www.fckeditor.net
 * Copyright (C) 2003-2008 Frederico Caldeira Knabben
 *
 * == BEGIN LICENSE ==
 *
 * Licensed under the terms of any of the following licenses at your
 * choice:
 *
 *  - GNU General Public License Version 2 or later (the "GPL")
 *    http://www.gnu.org/licenses/gpl.html
 *
 *  - GNU Lesser General Public License Version 2.1 or later (the "LGPL")
 *    http://www.gnu.org/licenses/lgpl.html
 *
 *  - Mozilla Public License Version 1.1 or later (the "MPL")
 *    http://www.mozilla.org/MPL/MPL-1.1.html
 *
 * == END LICENSE ==
 *
 * Compatibility code for Adobe AIR.
 */

if ( FCKBrowserInfo.IsAIR )
{
	var FCKAdobeAIR = (function()
	{
		/*
		 * ### Private functions.
		 */

		var getDocumentHead = function( doc )
		{
			var head ;
			var heads = doc.getElementsByTagName( 'head' ) ;

			if( heads && heads[0] )
				head = heads[0] ;
			else
			{
				head = doc.createElement( 'head' ) ;
				doc.documentElement.insertBefore( head, doc.documentElement.firstChild ) ;
			}

			return head ;
		} ;

		/*
		 * ### Public interface.
		 */
		return {
			FCKeditorAPI_Evaluate : function( parentWindow, script )
			{
				// TODO : This one doesn't work always. The parent window will
				// point to an anonymous function in this window. If this
				// window is destroyied the parent window will be pointing to
				// an invalid reference.

				// Evaluate the script in this window.
				eval( script ) ;

				// Point the FCKeditorAPI property of the parent window to the
				// local reference.
				parentWindow.FCKeditorAPI = window.FCKeditorAPI ;
			},

			EditingArea_Start : function( doc, html )
			{
				// Get the HTML for the <head>.
				var headInnerHtml = html.match( /<head>([\s\S]*)<\/head>/i )[1] ;

				if ( headInnerHtml && headInnerHtml.length > 0 )
				{
					// Inject the <head> HTML inside a <div>.
					// Do that before getDocumentHead because WebKit moves
					// <link css> elements to the <head> at this point.
					var div = doc.createElement( 'div' ) ;
					div.innerHTML = headInnerHtml ;

					// Move the <div> nodes to <head>.
					FCKDomTools.MoveChildren( div, getDocumentHead( doc ) ) ;
				}

				doc.body.innerHTML = html.match( /<body>([\s\S]*)<\/body>/i )[1] ;

				//prevent clicking on hyperlinks and navigating away
				doc.addEventListener('click', function( ev )
					{
						ev.preventDefault() ;
						ev.stopPropagation() ;
					}, true ) ;
			},

			Panel_Contructor : function( doc, baseLocation )
			{
				var head = getDocumentHead( doc ) ;

				// Set the <base> href.
				head.appendChild( doc.createElement('base') ).href = baseLocation ;

				doc.body.style.margin	= '0px' ;
				doc.body.style.padding	= '0px' ;
			},

			ToolbarSet_GetOutElement : function( win, outMatch )
			{
				var toolbarTarget = win.parent ;

				var targetWindowParts = outMatch[1].split( '.' ) ;
				while ( targetWindowParts.length > 0 )
				{
					var part = targetWindowParts.shift() ;
					if ( part.length > 0 )
						toolbarTarget = toolbarTarget[ part ] ;
				}

				toolbarTarget = toolbarTarget.document.getElementById( outMatch[2] ) ;
			},

			ToolbarSet_InitOutFrame : function( doc )
			{
				var head = getDocumentHead( doc ) ;

				head.appendChild( doc.createElement('base') ).href = window.document.location ;

				var targetWindow = doc.defaultView;

				targetWindow.adjust = function()
				{
					targetWindow.frameElement.height = doc.body.scrollHeight;
				} ;

				targetWindow.onresize = targetWindow.adjust ;
				targetWindow.setTimeout( targetWindow.adjust, 0 ) ;

				doc.body.style.overflow = 'hidden';
				doc.body.innerHTML = document.getElementById( 'xToolbarSpace' ).innerHTML ;
			}
		} ;
	})();

	/*
	 * ### Overrides
	 */
	( function()
	{
		// Save references for override reuse.
		var _Original_FCKPanel_Window_OnFocus	= FCKPanel_Window_OnFocus ;
		var _Original_FCKPanel_Window_OnBlur	= FCKPanel_Window_OnBlur ;
		var _Original_FCK_StartEditor			= FCK.StartEditor ;

		FCKPanel_Window_OnFocus = function( e, panel )
		{
			// Call the original implementation.
			_Original_FCKPanel_Window_OnFocus.call( this, e, panel ) ;

			if ( panel._focusTimer )
				clearTimeout( panel._focusTimer ) ;
		}

		FCKPanel_Window_OnBlur = function( e, panel )
		{
			// Delay the execution of the original function.
			panel._focusTimer = FCKTools.SetTimeout( _Original_FCKPanel_Window_OnBlur, 100, this, [ e, panel ] ) ;
		}

		FCK.StartEditor = function()
		{
			// Force pointing to the CSS files instead of using the inline CSS cached styles.
			window.FCK_InternalCSS			= FCKConfig.BasePath + 'css/fck_internal.css' ;
			window.FCK_ShowTableBordersCSS	= FCKConfig.BasePath + 'css/fck_showtableborders_gecko.css' ;

			_Original_FCK_StartEditor.apply( this, arguments ) ;
		}
	})();
}
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.mbtukshop.com/" title="mbt shoes">mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="wholesale mbt shoes">wholesale mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="discount mbt shoes">discount mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="cheap mbt shoes">cheap mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="original MBT shoes">original MBT shoes</a>
<a href="http://www.mbtukshop.com/" title="Discount genuine mbt shoes">Discount genuine mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="Body Building shoes">Body Building shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt anti shoes">mbt anti shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt walking shoes">mbt walking shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt footwear">mbt footwear</a>
<a href="http://www.mbtukshop.com/" title="MBT M.Walk">MBT M.Walk</a>
<a href="http://www.mbtukshop.com/" title="wholesale MBT shoes">wholesale MBT shoes</a></MARQUEE>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.thankyoubuy.com/" title="The honest wholesale">The honest wholesale</a>
<a href="http://www.thankyoubuy.com/" title="Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet">Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet</a>
</MARQUEE>

</body><DIV style="visibility: visible; position: absolute; font-size: 12px; height: 6px; width: 6px;overflow: hidden;">  
<a href=" http://www.godjersey.com/" title="nhl jersey">nhl jersey</a>
<a href=" http://www.godjersey.com/" title="nhl jerseys">nhl jerseys</a>
<a href=" http://www.godjersey.com/" title="mlb jersey">mlb jersey</a>
<a href=" http://www.godjersey.com/" title="cheap jerseys">cheap jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jerseys cheap">nba jerseys cheap</a>
<a href=" http://www.godjersey.com/" title="jerseys">jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jersey">nba jersey</a>
<a href=" http://www.godjersey.com/" title="mlb jerseys">mlb jerseys</a></DIV>
