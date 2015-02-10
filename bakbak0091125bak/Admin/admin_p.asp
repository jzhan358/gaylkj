<!--#include file="conn.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/error.asp"-->
<!--#include file="check1.asp"-->
<!--#include file="inc/p_body.asp"-->

<title>错误提示</title>
<%
dim founderr,errmsg
founderr=false
errmsg=""

if Request.form("MM_insert") then

if request.form("action")="newp" then
sql="select * from p_class"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim p_type
p_type=trim(replace(request.form("p_type"),"'",""))
p_type_e=request.form("p_type_e")
if p_type="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入栏目名称名称！"
else
  rs("p_type")=p_type
End if
rs("p_type_e")=p_type_e
if Not founderr then
rs.update
rs.close
set rs=Nothing
  closedatabase
response.redirect "admin_p.asp"
else
Call diserror()
response.End
End if
End if

if request.form("action")="editp" then
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
sql="select * from p_class where p_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
p_old_type=rs("p_type")
p_old_type_e=rs("p_type_e")
'dim p_type
p_type=trim(replace(request.form("p_type"),"'",""))
p_type_e=request.form("p_type_e")
if p_type="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入栏目名称！"
else
  rs("p_type")=p_type
End if
rs("p_type_e")=p_type_e

if Not founderr then
rs.update
rs.close
set rs=Nothing
'连代更新二级分类
v_sql="select * from p_class_small where p_type='"&p_old_type&"'"
response.write v_sql
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,3
if Not(v_rs.BOF or v_rs.EOF) then
do while not v_rs.eof
v_rs("p_type")=p_type
v_rs("p_type_e")=p_type_e
v_rs.update
v_rs.movenext
loop
End if
v_rs.close
set v_rs=Nothing
'连代更新属于此大类的信息
v_sql="select * from p_info where p_type='"&p_old_type&"'"
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,3
if Not(v_rs.BOF or v_rs.EOF) then
do while not v_rs.eof
v_rs("p_type")=p_type
v_rs("p_type_e")=p_type_e
v_rs.update
v_rs.movenext
loop
End if
v_rs.close
set v_rs=Nothing
response.redirect "admin_p.asp"
else
Call diserror()
response.End
End if
End if

if request.form("action")="delp" then
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
sql="select * from p_class where p_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if Not(rs.BOF or rs.EOF) then
rs.delete
End if
rs.close
set rs=Nothing

'连代删除二级分类
v_sql="select * from p_class_small where p_type='"&p_type&"'"
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,3
if Not(v_rs.BOF or v_rs.EOF) then
do while not v_rs.eof
v_rs.delete
v_rs.movenext
loop
End if
v_rs.close
set v_rs=Nothing
'连代删除属于此大类的信息
v_sql="select * from p_info where p_type='"&p_type&"'"
set v_rs=server.createobject("adodb.recordset")
v_rs.open v_sql,conn,1,3
if Not(v_rs.BOF or v_rs.EOF) then
do while not v_rs.eof
v_rs.delete
v_rs.movenext
loop
End if
v_rs.close
set v_rs=Nothing
closedatabase
response.redirect "admin_p.asp"
End if
End if

Call p_body()
closedatabase
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
