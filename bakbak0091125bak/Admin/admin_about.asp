
<!--#include file="conn.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/error.asp"-->
<!--#include file="check1.asp"-->
<!--#include file="inc/about_body.asp"-->
<title>关于我们</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
dim founderr,errmsg
founderr=false
errmsg=""

if session("admin_flag")<>"into" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你尚未登录，或者超时了！请<a href='login.asp' target='_parent' >重新登录</a>！"
  Call diserror()
  response.End
else

if Request.form("MM_insert") then
if request.form("action")="newabout" then
sql="select * from inc_class"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim i_type
i_type=trim(replace(request.form("i_type"),"'",""))
i_language=request.form("i_language")
if i_type="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入栏目名称名称！"
else
  rs("i_type")=i_type
End if
rs("i_language")=i_language
if Not founderr then
rs.update
rs.close
set rs=Nothing
  closedatabase
response.redirect "admin_about.asp"
else
Call diserror()
response.End
End if
End if
if request.form("action")="editabout" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from inc_class where i_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
'dim i_type
i_type=trim(replace(request.form("i_type"),"'",""))
i_language=request.form("i_language")
if i_type="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入栏目名称！"
else
  rs("i_type")=i_type
End if
rs("i_language")=i_language
if Not founderr then
rs.update
rs.close
set rs=Nothing
response.redirect "admin_about.asp"
else
Call diserror()
response.End
End if
End if

if request.form("action")="delabout" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from inc_class where i_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=Nothing

v_sql="select * from inc_info where inc_class_id="&cint(request.Form("id"))
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,3
if not v_rs.eof then
   v_rs.delete
end if
v_rs.close
set v_rs=Nothing
closedatabase
response.redirect "admin_about.asp"
End if
End if

if request.form("action")="edittextabout" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from inc_info where inc_class_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
inc_text=request.form("inc_text")
if inc_text="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入栏目名称！"
else
  rs("inc_text")=inc_text
End if

if Not founderr then
rs.update
rs.close
set rs=Nothing
response.redirect "admin_about.asp"
else
Call diserror()
response.End
End if
End if

if request.form("action")="addabout" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from inc_info where inc_class_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.EOF and rs.BOF then
rs.addnew
inc_text=request.form("inc_text")
if inc_text="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入栏目名称！"
else
  rs("inc_text")=inc_text
  rs("inc_class_id")=cint(request.Form("id"))
End if

if Not founderr then
rs.update
rs.close
set rs=Nothing
response.redirect "admin_about.asp"
else
Call diserror()
response.End
End if
End if
End if

Call adminabout_body()
closedatabase
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
