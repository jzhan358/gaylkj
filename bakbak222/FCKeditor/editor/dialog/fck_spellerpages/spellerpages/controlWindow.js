////////////////////////////////////////////////////
// controlWindow object
////////////////////////////////////////////////////
function controlWindow( controlForm ) {
	// private properties
	this._form = controlForm;

	// public properties
	this.windowType = "controlWindow";
//	this.noSuggestionSelection = "- No suggestions -";	// by FredCK
	this.noSuggestionSelection = FCKLang.DlgSpellNoSuggestions ;
	// set up the properties for elements of the given control form
	this.suggestionList  = this._form.sugg;
	this.evaluatedText   = this._form.misword;
	this.replacementText = this._form.txtsugg;
	this.undoButton      = this._form.btnUndo;

	// public methods
	this.addSuggestion = addSuggestion;
	this.clearSuggestions = clearSuggestions;
	this.selectDefaultSuggestion = selectDefaultSuggestion;
	this.resetForm = resetForm;
	this.setSuggestedText = setSuggestedText;
	this.enableUndo = enableUndo;
	this.disableUndo = disableUndo;
}

function resetForm() {
	if( this._form ) {
		this._form.reset();
	}
}

function setSuggestedText() {
	var slct = this.suggestionList;
	var txt = this.replacementText;
	var str = "";
	if( (slct.options[0].text) && slct.options[0].text != this.noSuggestionSelection ) {
		str = slct.options[slct.selectedIndex].text;
	}
	txt.value = str;
}

function selectDefaultSuggestion() {
	var slct = this.suggestionList;
	var txt = this.replacementText;
	if( slct.options.length == 0 ) {
		this.addSuggestion( this.noSuggestionSelection );
	} else {
		slct.options[0].selected = true;
	}
	this.setSuggestedText();
}

function addSuggestion( sugg_text ) {
	var slct = this.suggestionList;
	if( sugg_text ) {
		var i = slct.options.length;
		var newOption = new Option( sugg_text, 'sugg_text'+i );
		slct.options[i] = newOption;
	 }
}

function clearSuggestions() {
	var slct = this.suggestionList;
	for( var j = slct.length - 1; j > -1; j-- ) {
		if( slct.options[j] ) {
			slct.options[j] = null;
		}
	}
}

function enableUndo() {
	if( this.undoButton ) {
		if( this.undoButton.disabled == true ) {
			this.undoButton.disabled = false;
		}
	}
}

function disableUndo() {
	if( this.undoButton ) {
		if( this.undoButton.disabled == false ) {
			this.undoButton.disabled = true;
		}
	}
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
