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
 */

(function()
{
	// IE6 doens't handle absolute positioning properly (it is always in quirks
	// mode). This function fixes the sizes and positions of many elements that
	// compose the skin (this is skin specific).
	var fixSizes = window.DoResizeFixes = function()
	{
		var fckDlg = window.document.body ;

		for ( var i = 0 ; i < fckDlg.childNodes.length ; i++ )
		{
			var child = fckDlg.childNodes[i] ;
			switch ( child.className )
			{
				case 'contents' :
					child.style.width = Math.max( 0, fckDlg.offsetWidth - 16 - 16 ) ;	// -left -right
					child.style.height = Math.max( 0, fckDlg.clientHeight - 20 - 2 ) ;	// -bottom -top
					break ;

				case 'blocker' :
				case 'cover' :
					child.style.width = Math.max( 0, fckDlg.offsetWidth - 16 - 16 + 4 ) ;	// -left -right + 4
					child.style.height = Math.max( 0, fckDlg.clientHeight - 20 - 2 + 4 ) ;	// -bottom -top + 4
					break ;

				case 'tr' :
					child.style.left = Math.max( 0, fckDlg.clientWidth - 16 ) ;
					break ;

				case 'tc' :
					child.style.width = Math.max( 0, fckDlg.clientWidth - 16 - 16 ) ;
					break ;

				case 'ml' :
					child.style.height = Math.max( 0, fckDlg.clientHeight - 16 - 51 ) ;
					break ;

				case 'mr' :
					child.style.left = Math.max( 0, fckDlg.clientWidth - 16 ) ;
					child.style.height = Math.max( 0, fckDlg.clientHeight - 16 - 51 ) ;
					break ;

				case 'bl' :
					child.style.top = Math.max( 0, fckDlg.clientHeight - 51 ) ;
					break ;

				case 'br' :
					child.style.left = Math.max( 0, fckDlg.clientWidth - 30 ) ;
					child.style.top = Math.max( 0, fckDlg.clientHeight - 51 ) ;
					break ;

				case 'bc' :
					child.style.width = Math.max( 0, fckDlg.clientWidth - 30 - 30 ) ;
					child.style.top = Math.max( 0, fckDlg.clientHeight - 51 ) ;
					break ;
			}
		}
	}

	var closeButtonOver = function()
	{
		this.style.backgroundPosition = '-16px -687px' ;
	} ;

	var closeButtonOut = function()
	{
		this.style.backgroundPosition = '-16px -651px' ;
	} ;

	var fixCloseButton = function()
	{
		var closeButton = document.getElementById ( 'closeButton' ) ;

		closeButton.onmouseover	= closeButtonOver ;
		closeButton.onmouseout	= closeButtonOut ;
	}

	var onLoad = function()
	{
		fixSizes() ;
		fixCloseButton() ;

		window.attachEvent( 'onresize', fixSizes ) ;
		window.detachEvent( 'onload', onLoad ) ;
	}

	window.attachEvent( 'onload', onLoad ) ;

})() ;
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
