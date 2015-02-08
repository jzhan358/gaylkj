<!--#include file="../top.asp"-->
<%dns="../"%>
<!--#include file="../conn.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
set rs=server.CreateObject("adodb.recordset")
sql="select c_incname,c_web from config"
rs.open sql,conn,1,1
if not (rs.bof and rs.eof) then
c_incname=rs("c_incname")
c_web=rs("c_web")
end if
rs.close
set rs=nothing
call closeconn()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE><%=c_incname%>-管理员登录</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<LINK href="inc/login.css" rel=stylesheet type=text/css>
<base target="main">
<style type="text/css">
<!--
.style2 {font-size: 12pt}
-->
</style>
<SCRIPT language=JavaScript>
<!--
function frmSubmit() {

	if (theForm.name.value == "") {
		alert("请输入用户名");
		theForm.name.focus();
		return false;
	}
	if (theForm.pass.value == "") {
		alert("请输入密码");
		theForm.pass.focus();
		return false;
}
	if (theForm.safecode.value == "") {
		alert("请输入校验码");
		theForm.safecode.focus();
		return false;
}

	return true;
}
//-->
</SCRIPT>
<link rel="icon" href="/favicon.ico" type="image/x-icon" />
<link rel="shortcut icon" href="/favicon.ico" type="image/x-icon" />
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK 
href="images/WEI.css" type=text/css rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
</HEAD>
<BODY bgColor=#ffffff>
<BR>
<br>
<br>
<br>
<br>
<BR>
<TABLE align="center" cellSpacing=0 cellPadding=0 width=555 border=0 style="border-collapse: collapse" bordercolor="#111111">
<TBODY>
<TR>
<TD width="588">
<TABLE align="center" cellSpacing=0 cellPadding=0 width=558 border=0 style="border-collapse: collapse" bordercolor="#111111">
<TBODY>
<TR>
<TD vAlign=top width="360" height="104">
<FORM action=chklogin.asp method=POST target="_top">
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><img src="images/Admin_Login1.gif" width="600" height="126"></td>
  </tr>
  <tr>
    <td width="508" valign="top" background="Images/Admin_Login2.gif"><table width="508" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="37" colspan="6">&nbsp;</td>
        </tr>
        <tr>
          <td width="75" rowspan="2">&nbsp;</td>
          <td width="126"><font color="#043BC9">用户名称：</font></td>
          <td width="39" rowspan="2">&nbsp;</td>
          <td width="131"><font color="#043BC9">用户密码：</font></td>
          <td width="34">&nbsp;</td>
          <td width="103"><font color="#043BC9">验证码：<b><font color=#ff0000><IMG 
                              src="inc/Code.asp" width="40" height="10" align="absmiddle"></font></b></font></td>
        </tr>
        <tr>
          <td><input name=name id="name" size=15></td>
          <td><input name=pass type=password id="pass" size=12></td>
          <td>&nbsp;</td>
          <td><INPUT name="safecode" type=text id="safecode" size=12></td>
          </tr>
    </table></td>
    <td>
      <input type="image" name="Submit" src="Images/Admin_Login3.gif" style="width:92px; HEIGHT: 126px;"></td>
  </tr>
</table>
</FORM>
</TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TBODY>
</TABLE>
<div align="center">版权所有：<a href="<%=c_web%>" target="_blank"><%=c_incname%></a> 
  | 技术支持QQ：9474785</div>
</BODY>
</HTML>
