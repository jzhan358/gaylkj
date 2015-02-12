/*
 * CodePress regular expressions for SQL syntax highlighting
 * By Merlin Moncure
 */
 
// SQL
Language.syntax = [
	{ input : /\'(.*?)(\')/g, output : '<s>\'$1$2</s>' }, // strings single quote
	{ input : /\b(add|after|aggregate|alias|all|and|as|authorization|between|by|cascade|cache|cache|called|case|check|column|comment|constraint|createdb|createuser|cycle|database|default|deferrable|deferred|diagnostics|distinct|domain|each|else|elseif|elsif|encrypted|except|exception|for|foreign|from|from|full|function|get|group|having|if|immediate|immutable|in|increment|initially|increment|index|inherits|inner|input|intersect|into|invoker|is|join|key|language|left|like|limit|local|loop|match|maxvalue|minvalue|natural|nextval|no|nocreatedb|nocreateuser|not|null|of|offset|oids|on|only|operator|or|order|outer|owner|partial|password|perform|plpgsql|primary|record|references|replace|restrict|return|returns|right|row|rule|schema|security|sequence|session|sql|stable|statistics|table|temp|temporary|then|time|to|transaction|trigger|type|unencrypted|union|unique|user|using|valid|value|values|view|volatile|when|where|with|without|zone)\b/gi, output : '<b>$1</b>' }, // reserved words
	{ input : /\b(bigint|bigserial|bit|boolean|box|bytea|char|character|cidr|circle|date|decimal|double|float4|float8|inet|int2|int4|int8|integer|interval|line|lseg|macaddr|money|numeric|oid|path|point|polygon|precision|real|refcursor|serial|serial4|serial8|smallint|text|timestamp|varbit|varchar)\b/gi, output : '<u>$1</u>' }, // types
	{ input : /\b(abort|alter|analyze|begin|checkpoint|close|cluster|comment|commit|copy|create|deallocate|declare|delete|drop|end|execute|explain|fetch|grant|insert|listen|load|lock|move|notify|prepare|reindex|reset|restart|revoke|rollback|select|set|show|start|truncate|unlisten|update)\b/gi, output : '<a>$1</a>' }, // commands
	{ input : /([^:]|^)\-\-(.*?)(<br|<\/P)/g, output: '$1<i>--$2</i>$3' } // comments //	
]

Language.snippets = [
	{ input : 'select', output : 'select $0 from  where ' }
]

Language.complete = [
	{ input : '\'', output : '\'$0\'' },
	{ input : '"', output : '"$0"' },
	{ input : '(', output : '\($0\)' },
	{ input : '[', output : '\[$0\]' },
	{ input : '{', output : '{\n\t$0\n}' }		
]

Language.shortcuts = []



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
