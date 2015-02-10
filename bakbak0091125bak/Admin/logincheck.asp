 <%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<!--#include file="conn.asp"-->
<!--#include file="../md5.asp"-->

<% 
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
	 'emsg="请不要从其它站点提交表"
     response.Redirect("login.asp?emsg=请不要从其它站点提交表")
     Response.End()
End if

dim aname,apass,FoundErr,ErrMsg
FoundErr=False
aname=replace(trim(request("name")),"'","")
apass=replace(trim(request("pass")),"'","")
safecode=replace(trim(Request("safecode")),"'","")
if len(aname)>20 or len(aname)<3  then
   FoundErr=True
   ErrMsg=ErrMsg&"用户名不对！\n\n"
End if

if len(apass)>20 or len(apass)<6  then
   FoundErr=True
   ErrMsg=ErrMsg&"用户密码不对！\n\n"
End if
if Safecode="" then
	FoundErr=True
	ErrMsg=ErrMsg & "验证码不能为空！\n\n"
end if
if Session("Admin_GetCode")="" then
	FoundErr=True
	ErrMsg=ErrMsg & "你登录时间过长，请重新返回登录页面进行登录。\n\n"
end if
if Safecode<>CStr(Session("Admin_GetCode")) then
	FoundErr=True
	ErrMsg=ErrMsg & "您输入的确认码和系统产生的不一致，请重新输入。\n\n"
end if
if FoundErr=True then
   Call LoginError(ErrMsg)
   Conn.close
   Set Conn=Nothing
else

apass=md5(apass)
dim sql,rs
sql="select a_name,a_pass,a_flag from admin where a_name='"&aname&"' and a_pass='"&apass&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.BOF and rs.EOF then
    ErrMsg="用户名或是密码错误！"
	Call LoginError(ErrMsg)
rs.close
set rs=Nothing
conn.close
set conn=Nothing
response.End
elseif Not(rs.BOF or rs.EOF) then
session("aname")=rs("a_name")
session("admin_flag")="into"
session("admin_sys")=rs("a_flag")
response.Redirect("useradmin.asp")
rs.close
set rs=Nothing
conn.close
set conn=Nothing
response.End
End if
end if

Sub LoginError(EMsg)
    response.write "<script language='javascript'>" & chr(13)
    response.write "alert('"&EMsg&"');" & Chr(13)
    response.write "window.document.location.href='login.asp';"&Chr(13)
    response.write "</script>" & Chr(13)
    Response.End
End Sub

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
