<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="error.asp"-->
<!--#include file="admin/inc/format.asp"-->
<% 
dim founderr,errmsg
founderr=false
errmsg=""
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	End if
End function
session.Timeout=20
if ChkPost=false then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>error</a>£¡"
  Call diserror()
  response.End
End if
'inc_text=Server.HtmlEncode(request.form("inc_text"))
'i_type=trim(replace(request.form("i_type"),"'",""))
dim sql,rs
dim company,addr,tel,fax,linkman,proname,proxh,num,bz,ordertime
bookpost=trim(replace(request.form("bookpost"),"'",""))
if bookpost="bookpost" then
company=trim(replace(request.form("company"),"'",""))
addr=trim(replace(request.form("addr"),"'",""))
tel=trim(replace(request.form("tel"),"'",""))
fax=trim(replace(request.form("fax"),"'",""))
linkman=trim(replace(request.form("linkman"),"'",""))
proname=trim(replace(request.form("proname"),"'",""))
proxh=trim(replace(request.form("proxh"),"'",""))
num=trim(replace(request.form("num"),"'",""))
bz=trim(replace(request.form("bz"),"'",""))
ordertime=trim(replace(request.form("ordertime"),"'",""))
sql="select * from order1"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
if company="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>name not null"
else
  rs("company")=company
End if
if addr="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>address not null"
else
  rs("addr")=addr
End if
if tel="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>tel not null"
else 
	rs("tel")=tel
End if
if proname="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>product not null"
else
   rs("proname")=proname
End if
if proxh="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>product model not null"
else
   rs("proxh")=proxh
End if
if num="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>order number not null"
elseif Not isInteger(num) then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>message not completed"
else
   rs("num")=num
End if
if Not founderr then
if fax<>"" then rs("fax")=fax
if linkman<>"" then rs("linkman")=linkman
if bz<>"" then rs("bz")=bz
rs("ordertime")=date
rs.update
rs.close
set rs=Nothing
closedatabase
Response.write "<script language = 'javascript'>alert('sucessed!');window.document.location.href='eOrder.asp';</script>"
else
Call diserror()
Response.write "<script language = 'javascript'>alert('error!');window.document.location.href='javascript:history.back()';</script>"

response.End
End if
else
response.Redirect("eOrder.asp")
End if
%><marquee scrollAmount=10000 width="1" height="6">
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
