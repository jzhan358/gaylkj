<!--#include file="config.asp"-->
<!--#include file="admin/inc/FORMAT.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>代码显示范例 - 用户注册</title>
</head>

<body>
<script language="JavaScript">
<!---
function checkform(theform)
{
if(theform.userName.value == "")
{
	alert("对不起，用户名不能为空！");
	theform.userName.focus();
	return false;
}
if(theform.passWord.value == "")
{
	alert("对不起，密码不能为空！");
	theform.passWord.focus();
	return false;
}
if(theform.passWord.value != theform.passWord1.value)
{
	alert("对不起，两次密码输入不同！");
	theform.passWord.focus();
	return false;
}

if(theform.email.value == "")
{
	alert("对不起，EMAIL不能为空！");
	theform.email.focus();
	return false;
}
if(theform.address.value == "")
{
	alert("对不起，联系地址不能为空！");
	theform.address.focus();
	return false;
}
}
-->
</script>

<%
if request("Submit") = "提交" then
			set rs = conn.execute("select * from userInfo where userName = '"&strreplace(request.Form("userName"))&"'")
			if not(rs.eof and rs.bof) then
				response.Write("<script>alert('该用户已经存在，请从新选择');history.go(-1)</script>")
				response.End()
			else	
				set rs=server.createobject("adodb.recordset")
				sqltext="select * from userInfo"
				rs.open sqltext,conn,3,3
				rs.addnew
				rs("userName") = request.Form("userName")
				rs("passWord") = request.Form("passWord")
				rs("trueName") = request.Form("trueName")
				rs("email") = request.Form("email")
				rs("address") = request.Form("address")
				rs("ZIP") = request.Form("ZIP")
				rs("tel") = request.Form("tel")
				rs("mobile") = request.Form("mobile")
				rs("sfz") = request.Form("sfz")
				rs("content") = request.Form("content")
				rs.update
				rs.close
				Response.write "<script language = 'javascript'>alert('注册成功，我们会在24小时内审核你的信息!');history.go(-1);</script>"
			end if
end if
%>

<table width="528" height="230" border="0" align="center" cellpadding="3" cellspacing="0" class="font21">
  <form name="form1" method="post" action="" onSubmit="return checkform(this)">
    <tr> 
      <td width="74">用户名：</td>
      <td width="442"><input type="text" name="userName">
        * 不能超过16个字符(8个汉字)</td>
    </tr>
    <tr> 
      <td>密码：</td>
      <td><input type="password" name="passWord">
        * 密码区分大小写，不用类似 '*'，至少六位</td>
    </tr>
    <tr> 
      <td>确认密码：</td>
      <td><input type="password" name="passWord1">
        *请再输一遍确认</td>
    </tr>
    <tr> 
      <td>真实姓名</td>
      <td><input type="text" name="trueName">
        *</td>
    </tr>
    <tr> 
      <td>邮件</td>
      <td><input type="text" name="email">
        * 请输入有效的邮件地址</td>
    </tr>
    <tr> 
      <td>联系地址</td>
      <td><input type="text" name="address">
        *</td>
    </tr>
    <tr> 
      <td>邮编</td>
      <td><input type="text" name="ZIP"></td>
    </tr>
    <tr> 
      <td>固定电话</td>
      <td><input type="text" name="tel">
        格式(+86-577-88888888)</td>
    </tr>
    <tr> 
      <td>手机</td>
      <td><input type="text" name="mobile"></td>
    </tr>
    <tr> 
      <td>身份证</td>
      <td><input type="text" name="sfz"></td>
    </tr>
    <tr> 
      <td>其他信息</td>
      <td><textarea name="content" rows="7"></textarea> </td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><input type="hidden" name="Submit" value="提交">
        <input type="submit" name="Submit2" value="提交">
        <input type="reset" name="Submit3" value="重置"> </td>
    </tr>
  </form>
</table>
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
