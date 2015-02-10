
<!--#include file="conn.asp"-->
<!--#include file="check1.asp"-->
<!--#include file="nocache.asp"-->
<!--#include file="../md5.asp"-->
<!--#include file="inc/error.asp"-->

<html>
<head>
<title>用户管理</title>
</head>
<style type="text/css">
a:link { color:#000000;text-decoration:none}
a:hover {color:#666666;}
a:visited {color:#000000;text-decoration:none}

td {FONT-SIZE: 9pt; FILTER: dropshadow(color=#FFFFFF,offx=1,offy=1); COLOR: #000000; FONT-FAMILY: "宋体"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}
</style>

<body>
<%
dim emsg,sql,sqlupdate
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
if session("admin_flag")="" or isNull(session("admin_flag")) or isEmpty(session("admin_flag")) or session("admin_flag")<>"into" then
emsg="你的登陆信息失效，请重新登陆！"
response.Redirect("login.asp?emsg="&emsg)
response.End()
elseif session("admin_flag")="into" then
  dim post,EVS_repass,EVS_rename,pass1
  dim passold,nameold
  post = Trim(Request.QueryString("post"))
  if post="edit" then 
     if ChkPost=false then
	 emsg="请不要从其它站点提交表"
     response.Redirect("login.asp?emsg="&emsg)
     Response.End()
     End if
     EVS_repass = Trim(Request.Form("pass"))
	 
	 pass1  = Trim(Request.Form("pass1"))
	
	 passold= Trim(Request.Form("passold"))
  

	 if len(EVS_repass)>20 or len(EVS_repass)<6 then
	    response.write "<script language='javascript'>" & chr(13)
        response.write "alert('密码错误!提示:长度应在6-20个字符这间！');" & Chr(13)
        response.write "window.document.location.href='admin.asp';"&Chr(13)
        response.write "</script>" & Chr(13)
        Response.End
	End if
    if  EVS_repass<>pass1 then
	    response.write "<script language='javascript'>" & chr(13)
        response.write "alert('两次的密码不一样，请牢记你的密码!');" & Chr(13)
        response.write "window.document.location.href='admin.asp';"&Chr(13)
        response.write "</script>" & Chr(13)
        Response.End
	End if
	EVS_repass = md5(EVS_repass)
	passold=md5(passold)
  sql="select a_name from admin where a_name='admin' and a_pass='"&passold&"'"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,conn,1,1
  if  rs.BOF and rs.EOF then
     emsg="你的登陆信息失效，请重新登陆！"
     response.Redirect("login.asp?emsg="&emsg)
     response.End()
	 rs.close
	 set rs=Nothing
	 conn.close
	 set conn=Nothing
  elseif Not(rs.BOF or rs.EOF) then
    sqlupdate="update admin set a_pass='"&EVS_repass&"' where a_name='admin' and a_pass='"&passold&"'"
	conn.Execute  sqlupdate
	response.write "<script language='javascript'>" & chr(13)
    response.write "alert('密码修改成功，请牢记你的密码!');" & Chr(13)
    response.write "window.document.location.href='admin.asp';"&Chr(13)
    response.write "</script>" & Chr(13)
    Response.End
	rs.close
	set rs=Nothing
	conn.close
	set conn=Nothing
    response.End()
  End if
 End if
 End if


%>
	<form action="admin.asp?post=edit" method="post" name="f2" target="_self">
        <table  width="780" align="center" cellspacing="1" bgcolor="#999999" >
          <tr bgcolor="#CCCCCC">
            <td height="30" colspan="2">
            <p align="center"><b><font color="#660000">修 改 用 户</font></b></td>   
          </tr>
          
		  <tr>
            <td width="19%" height="30" align="center" bordercolordark="#FFFFFF" bgcolor="#CDE7D1">旧&nbsp;密&nbsp;码</td>
            <td height="30" align="left" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
              &nbsp;
              <input name="passold" type="password" id="passold" size="20" maxlength="20">
              &nbsp;</td>
          </tr>

          <tr>
            <td width="19%" height="30" align="center" bordercolordark="#FFFFFF" bgcolor="#CDE7D1">用户密码</td>
            <td height="30" align="left" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
              &nbsp;
              <input name="pass" type="password" id="pass" size="20" maxlength="20">
              &nbsp;[字母开数字组成,长度不大于20个字符]</td>
          </tr>
		  <tr>
            <td width="19%" height="30" align="center" bordercolordark="#FFFFFF" bgcolor="#CDE7D1">确认密码</td>
            <td height="30" align="left" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
              &nbsp;
              <input name="pass1" type="password" id="pass1" size="20" maxlength="20"> 
            </td>
          </tr>
		  
          <tr bgcolor="#CCCCCC">
            <td height="30" bordercolordark="#FFFFFF" align="center" colspan="2"><input name="Submit" type="submit" value="确认修改">　
              <input type="reset" name="Submit" value="取消修改"> 
            　<font color="#FFFFFF">&nbsp;</font></td>
          </tr>
        </table>
		</form>
</body>
</html>
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
