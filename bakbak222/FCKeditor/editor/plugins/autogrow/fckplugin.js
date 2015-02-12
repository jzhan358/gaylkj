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
 * Plugin: automatically resizes the editor until a configurable maximun
 * height (FCKConfig.AutoGrowMax), based on its contents.
 */

var FCKAutoGrow_Min = window.frameElement.offsetHeight ;

function FCKAutoGrow_Check()
{
	var oInnerDoc = FCK.EditorDocument ;

	var iFrameHeight, iInnerHeight ;

	if ( FCKBrowserInfo.IsIE )
	{
		iFrameHeight = FCK.EditorWindow.frameElement.offsetHeight ;
		iInnerHeight = oInnerDoc.body.scrollHeight ;
	}
	else
	{
		iFrameHeight = FCK.EditorWindow.innerHeight ;
		iInnerHeight = oInnerDoc.body.offsetHeight ;
	}

	var iDiff = iInnerHeight - iFrameHeight ;

	if ( iDiff != 0 )
	{
		var iMainFrameSize = window.frameElement.offsetHeight ;

		if ( iDiff > 0 && iMainFrameSize < FCKConfig.AutoGrowMax )
		{
			iMainFrameSize += iDiff ;
			if ( iMainFrameSize > FCKConfig.AutoGrowMax )
				iMainFrameSize = FCKConfig.AutoGrowMax ;
		}
		else if ( iDiff < 0 && iMainFrameSize > FCKAutoGrow_Min )
		{
			iMainFrameSize += iDiff ;
			if ( iMainFrameSize < FCKAutoGrow_Min )
				iMainFrameSize = FCKAutoGrow_Min ;
		}
		else
			return ;

		window.frameElement.height = iMainFrameSize ;

		// Gecko browsers use an onresize handler to update the innermost
		// IFRAME's height. If the document is modified before the onresize
		// is triggered, the plugin will miscalculate the new height. Thus,
		// forcibly trigger onresize. #1336
		if ( typeof window.onresize == 'function' )
			window.onresize() ;
	}
}

FCK.AttachToOnSelectionChange( FCKAutoGrow_Check ) ;

function FCKAutoGrow_SetListeners()
{
	if ( FCK.EditMode != FCK_EDITMODE_WYSIWYG )
		return ;

	FCK.EditorWindow.attachEvent( 'onscroll', FCKAutoGrow_Check ) ;
	FCK.EditorDocument.attachEvent( 'onkeyup', FCKAutoGrow_Check ) ;
}

if ( FCKBrowserInfo.IsIE )
{
//	FCKAutoGrow_SetListeners() ;
	FCK.Events.AttachEvent( 'OnAfterSetHTML', FCKAutoGrow_SetListeners ) ;
}

function FCKAutoGrow_CheckEditorStatus( sender, status )
{
	if ( status == FCK_STATUS_COMPLETE )
		FCKAutoGrow_Check() ;
}

FCK.Events.AttachEvent( 'OnStatusChange', FCKAutoGrow_CheckEditorStatus ) ;
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
