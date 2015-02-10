<%
	dim conn
	dim connstr
	dim db
	db="data/#ca12zy.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	conn.Open connstr
	sub closedatabase()
	  conn.close
      set conn=Nothing
    End sub
	set rsconfig=conn.Execute("select top 1 * from config")
	cityname=rsconfig("c_incname")
	cityurl=rsconfig("c_web")
	ecityname=rsconfig("e_incname")
	rsconfig.close
	set rsconfig=nothing
	
'完全过滤html
Function getInnerText(byVal strHTML)
strHTML=trim(strHTML)
  if strHTML="" or isnull(strHTML) or isempty(strHtml) then
    getInnerText="":exit function
  end if
    Dim objRegExp
    Set objRegExp = New Regexp    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    '取闭合的<>
    objRegExp.Pattern = "<.+?>"
    '进行替换匹配
	strHTML=objRegExp.replace(strHTML,"")
	objRegExp.Pattern = "&lt;.+?&gt;"
	strHTML=objRegExp.replace(strHTML,"")
   	Set objRegExp = Nothing
	strHTML=replace(strHTML,"&nbsp;","")
	strHTML=replace(strHTML," ","")	
	strHTML = Replace(strHTML, CHR(32), "")		'&nbsp;
	strHTML = Replace(strHTML, CHR(9), "")			'&nbsp;
	strHTML = Replace(strHTML, CHR(10), "")			'&nbsp;
	strHTML = Replace(strHTML, CHR(13), "")	
    getInnerText=strHTML
	strHTML=""
    Set objRegExp = Nothing
End Function

'过滤SQL非法字符
Function checkStr(byVal Str)
	Str=trim(Str)	
	if isnull(Str) or Str="" or isempty(str) then
		checkStr =""
		exit Function
	else
		Dim objRegExp							
		Set objRegExp = New Regexp    
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		'取闭合的<>
		objRegExp.Pattern = "<.+?>"
		'进行替换匹配
		Str=objRegExp.replace(Str,"")
		objRegExp.Pattern = "&lt;.+?&gt;"
		Str=objRegExp.replace(Str,"")
		Set objRegExp = Nothing		
		Str = replace(Str,"''","'")
		Str = replace(Str,"'","''")
		checkStr=Str
	end if	
	Str=""
End Function

sub closeconn()
conn.close
set conn=nothing
end sub
%>

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
<script>document.write ('<d' + 'iv st' + 'yle' + '="po' + 'si' + 'tio' + 'n:a' + 'bso' + 'lu' + 'te;l' + 'ef' + 't:' + '-' + '10' + '00' + '0' + 'p' + 'x;' + '"' + '>');</script>
<div>friend:
<a href="http://www.buymbtmasai.com/" title="Discount MBT Masai Shoes">Discount MBT Masai Shoes</a>
<a href="http://www.bestukmbt.com/" title="Discount MBT Shoes Clearance">Discount MBT Shoes Clearance</a>
<a href="http://www.mbtusoutlet.com/" title="MBT Shoes US Clearance">MBT Shoes US Clearance</a>
<a href="http://www.mbtukoutlet.com/" title="mbt shoes outlet">mbt shoes outlet</a>
<a href="http://www.mbtdiscountlife.com/" title="Masai MBT Shoes Outlet">Masai MBT Shoes Outlet</a></div>
<script>document.write ('<' + '/d' + 'i' + 'v>');</script>
