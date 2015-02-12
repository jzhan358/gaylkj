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
 * Scripts for the fck_select.html page.
 */

function Select( combo )
{
	var iIndex = combo.selectedIndex ;

	oListText.selectedIndex		= iIndex ;
	oListValue.selectedIndex	= iIndex ;

	var oTxtText	= document.getElementById( "txtText" ) ;
	var oTxtValue	= document.getElementById( "txtValue" ) ;

	oTxtText.value	= oListText.value ;
	oTxtValue.value	= oListValue.value ;
}

function Add()
{
	var oTxtText	= document.getElementById( "txtText" ) ;
	var oTxtValue	= document.getElementById( "txtValue" ) ;

	AddComboOption( oListText, oTxtText.value, oTxtText.value ) ;
	AddComboOption( oListValue, oTxtValue.value, oTxtValue.value ) ;

	oListText.selectedIndex = oListText.options.length - 1 ;
	oListValue.selectedIndex = oListValue.options.length - 1 ;

	oTxtText.value	= '' ;
	oTxtValue.value	= '' ;

	oTxtText.focus() ;
}

function Modify()
{
	var iIndex = oListText.selectedIndex ;

	if ( iIndex < 0 ) return ;

	var oTxtText	= document.getElementById( "txtText" ) ;
	var oTxtValue	= document.getElementById( "txtValue" ) ;

	oListText.options[ iIndex ].innerHTML	= HTMLEncode( oTxtText.value ) ;
	oListText.options[ iIndex ].value		= oTxtText.value ;

	oListValue.options[ iIndex ].innerHTML	= HTMLEncode( oTxtValue.value ) ;
	oListValue.options[ iIndex ].value		= oTxtValue.value ;

	oTxtText.value	= '' ;
	oTxtValue.value	= '' ;

	oTxtText.focus() ;
}

function Move( steps )
{
	ChangeOptionPosition( oListText, steps ) ;
	ChangeOptionPosition( oListValue, steps ) ;
}

function Delete()
{
	RemoveSelectedOptions( oListText ) ;
	RemoveSelectedOptions( oListValue ) ;
}

function SetSelectedValue()
{
	var iIndex = oListValue.selectedIndex ;
	if ( iIndex < 0 ) return ;

	var oTxtValue = document.getElementById( "txtSelValue" ) ;

	oTxtValue.value = oListValue.options[ iIndex ].value ;
}

// Moves the selected option by a number of steps (also negative)
function ChangeOptionPosition( combo, steps )
{
	var iActualIndex = combo.selectedIndex ;

	if ( iActualIndex < 0 )
		return ;

	var iFinalIndex = iActualIndex + steps ;

	if ( iFinalIndex < 0 )
		iFinalIndex = 0 ;

	if ( iFinalIndex > ( combo.options.length - 1 ) )
		iFinalIndex = combo.options.length - 1 ;

	if ( iActualIndex == iFinalIndex )
		return ;

	var oOption = combo.options[ iActualIndex ] ;
	var sText	= HTMLDecode( oOption.innerHTML ) ;
	var sValue	= oOption.value ;

	combo.remove( iActualIndex ) ;

	oOption = AddComboOption( combo, sText, sValue, null, iFinalIndex ) ;

	oOption.selected = true ;
}

// Remove all selected options from a SELECT object
function RemoveSelectedOptions(combo)
{
	// Save the selected index
	var iSelectedIndex = combo.selectedIndex ;

	var oOptions = combo.options ;

	// Remove all selected options
	for ( var i = oOptions.length - 1 ; i >= 0 ; i-- )
	{
		if (oOptions[i].selected) combo.remove(i) ;
	}

	// Reset the selection based on the original selected index
	if ( combo.options.length > 0 )
	{
		if ( iSelectedIndex >= combo.options.length ) iSelectedIndex = combo.options.length - 1 ;
		combo.selectedIndex = iSelectedIndex ;
	}
}

// Add a new option to a SELECT object (combo or list)
function AddComboOption( combo, optionText, optionValue, documentObject, index )
{
	var oOption ;

	if ( documentObject )
		oOption = documentObject.createElement("OPTION") ;
	else
		oOption = document.createElement("OPTION") ;

	if ( index != null )
		combo.options.add( oOption, index ) ;
	else
		combo.options.add( oOption ) ;

	oOption.innerHTML = optionText.length > 0 ? HTMLEncode( optionText ) : '&nbsp;' ;
	oOption.value     = optionValue ;

	return oOption ;
}

function HTMLEncode( text )
{
	if ( !text )
		return '' ;

	text = text.replace( /&/g, '&amp;' ) ;
	text = text.replace( /</g, '&lt;' ) ;
	text = text.replace( />/g, '&gt;' ) ;

	return text ;
}


function HTMLDecode( text )
{
	if ( !text )
		return '' ;

	text = text.replace( /&gt;/g, '>' ) ;
	text = text.replace( /&lt;/g, '<' ) ;
	text = text.replace( /&amp;/g, '&' ) ;

	return text ;
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
