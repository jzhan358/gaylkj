
<!--#include file="conn.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/error.asp"-->
<!--#include file="check1.asp"-->
<!--#include file="inc/config_body.asp"-->

<title>公司基本信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
dim founderr,errmsg
founderr=false
errmsg=""
dim c_web,c_incname,c_linkman,c_addr,c_email,c_tel,c_fax
if Request.form("MM_insert") then

if request.form("action")="addconfig" then
sql="select * from config"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if  Not(rs.BOF or rs.EOF) then
   rs.close
   set rs=Nothing
   founderr=true
   errmsg=errmsg+"<br>"+"<li>已经有了基本信息，不能增加了，请使用修改功能！"
   Call diserror()
   response.End
elseif rs.BOF and rs.EOF then

rs.addnew
c_web=trim(replace(request.form("c_web"),"'",""))
c_incname=trim(replace(request.form("c_incname"),"'",""))
c_linkman=trim(replace(request.form("c_linkman"),"'",""))
c_addr=trim(replace(request.form("c_addr"),"'",""))
c_email=trim(replace(request.form("c_email"),"'",""))
c_tel=trim(replace(request.form("c_tel"),"'",""))
c_fax=trim(replace(request.form("c_fax"),"'",""))
'英文信息
e_web=trim(replace(request.form("e_web"),"'",""))
e_incname=trim(replace(request.form("e_incname"),"'",""))
e_linkman=trim(replace(request.form("e_linkman"),"'",""))
e_addr=trim(replace(request.form("e_addr"),"'",""))
e_email=trim(replace(request.form("e_email"),"'",""))
e_tel=trim(replace(request.form("e_tel"),"'",""))
e_fax=trim(replace(request.form("e_fax"),"'",""))

if c_web="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入网址！"
else
  rs("c_web")=c_web
End if

if c_incname="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司名称！"
else
  rs("c_incname")=c_incname
End if

if c_linkman="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系人！"
else
  rs("c_linkman")=c_linkman
End if
if c_addr="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司地址！"
else
  rs("c_addr")=c_addr
End if
if c_email="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入电子邮箱！"
else
  rs("c_email")=c_email
End if
if c_tel="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系电话！"
else
  rs("c_tel")=c_tel
End if
if c_fax="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入传真号码！"
else
  rs("c_fax")=c_fax
End if
rs("e_web")=e_web
rs("e_incname")=e_incname
rs("e_linkman")=e_linkman
rs("e_addr")=e_addr
rs("e_email")=e_email
rs("e_tel")=e_tel
rs("e_fax")=e_fax

if Not founderr then
rs.update
rs.close
set rs=Nothing
  closedatabase
response.redirect "admin_config.asp"
else
Call diserror()
response.End
End if
End if

End if

if request.form("action")="editconfig" then
sql="select * from config"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.BOF and rs.EOF  then
   rs.close
   set rs=Nothing
   founderr=true
   errmsg=errmsg+"<br>"+"<li>没有输入基本信息，不能修改了，请先输入基本信息！"
   Call diserror()
   response.End
elseif  Not(rs.BOF or rs.EOF) then
'dim c_web,c_incname,c_linkman,c_addr,c_email,c_tel,c_fax
c_web=trim(replace(request.form("c_web"),"'",""))
c_incname=trim(replace(request.form("c_incname"),"'",""))
c_linkman=trim(replace(request.form("c_linkman"),"'",""))
c_addr=trim(replace(request.form("c_addr"),"'",""))
c_email=trim(replace(request.form("c_email"),"'",""))
c_tel=trim(replace(request.form("c_tel"),"'",""))
c_fax=trim(replace(request.form("c_fax"),"'",""))
'英文信息
e_web=trim(replace(request.form("e_web"),"'",""))
e_incname=trim(replace(request.form("e_incname"),"'",""))
e_linkman=trim(replace(request.form("e_linkman"),"'",""))
e_addr=trim(replace(request.form("e_addr"),"'",""))
e_email=trim(replace(request.form("e_email"),"'",""))
e_tel=trim(replace(request.form("e_tel"),"'",""))
e_fax=trim(replace(request.form("e_fax"),"'",""))

if c_web="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入网址！"
else
  rs("c_web")=c_web
End if

if c_incname="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司名称！"
else
  rs("c_incname")=c_incname
End if

if c_linkman="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系人！"
else
  rs("c_linkman")=c_linkman
End if
if c_addr="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入公司地址！"
else
  rs("c_addr")=c_addr
End if
if c_email="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入电子邮箱！"
else
  rs("c_email")=c_email
End if
if c_tel="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入联系电话！"
else
  rs("c_tel")=c_tel
End if
if c_fax="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入传真号码！"
else
  rs("c_fax")=c_fax
End if
rs("e_web")=e_web
rs("e_incname")=e_incname
rs("e_linkman")=e_linkman
rs("e_addr")=e_addr
rs("e_email")=e_email
rs("e_tel")=e_tel
rs("e_fax")=e_fax

if Not founderr then

rs.update
rs.close
set rs=Nothing
Dim EVS_C_Sql,EVS_C_Rs
EVS_C_Sql="select top 1 c_copyright from config"
Set EVS_C_Rs=Conn.Execute(EVS_C_Sql)
IF  Instr(EVS_C_Rs(0),"FIGO")=0 Then
    Conn.Execute("Update config Set c_copyright='程序制作:<a href=http://www.delphifan.com/studio target=_blank>FIGO</a>'")
	response.end()
End IF
EVS_C_Rs.Close
Set EVS_C_Rs=Nothing
response.redirect "admin_config.asp"
else
Call diserror()
response.End
End if
End if
End if
End if

Call adminconfig_body()
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
