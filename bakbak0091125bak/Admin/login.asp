<%
option explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'主要是使随机出现的图片数字随机
%>
<!--#include file="inc/config.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE><%=rs_config("c_incname")%>-管理员登录</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
<META http-equiv=Content-Type content="text/html; charset=gb2312">
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
<FORM action=logincheck.asp method=POST target="_top">
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
<div align="center">版权所有：<a href="<%=rs_config("c_web")%>" target="_blank"><%=rs_config("c_incname")%></a> 
  | 技术支持QQ：4215958</div>
</BODY>
</HTML>
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
