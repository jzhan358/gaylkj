<%sub adminconfig_body()
%>

<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
A{TEXT-DECORATION: none}
A:link {COLOR: #666666; FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:visited {COLOR: #666666; FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:active {FONT-FAMILY: 宋体; TEXT-DECORATION: none}
A:hover {BORDER-BOTTOM: 1px dotted; BORDER-LEFT-WIDTH: 1px; BORDER-RIGHT-WIDTH: 1px; BORDER-TOP-WIDTH: 1px; COLOR: #ff6600; TEXT-DECORATION: none}
BODY {
FONT-SIZE: 12px;
COLOR: #666666;
FONT-FAMILY:  宋体;
background-color: #ffffff; 
background-image: url(images/show.gif);
SCROLLBAR-FACE-COLOR: #e8e7e7; 
SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
SCROLLBAR-SHADOW-COLOR: #ffffff; 
SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
SCROLLBAR-ARROW-COLOR: #ff6600; 
SCROLLBAR-TRACK-COLOR: #EFEFEF; 
SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
SCROLLBAR-BASE-COLOR: #000000
}
TABLE {BORDER-COLLAPSE: collapse; FONT-FAMILY: 宋体; FONT-SIZE: 9pt}
.button{height:18px;width:62px;background:#f6f6f9 url(images/ButtonBg.gif); border:solid 1px #5589AA;color: #000000 ;FONT-SIZE: 9pt}
.lanyu{border:solid 1px #5589AA;color: #000000 ; font-size: 12px;}
.font {  filter: DropShadow(Color=#cccccc, OffX=2, OffY=1, Positive=2); text-decoration: none; font-size: 9pt}
</style>
<LINK href="inc/djcss.css" type=text/css rel=StyleSheet>
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
</HEAD>
<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) background=inc/dj_bg.gif>

     <% 
       sql="select * from config"
       set rs=server.createobject("adodb.recordset")
       rs.open sql,conn,1,1
     %>
      <table width="98%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#C0C0C0" >
        <% if rs.BOF and rs.EOF then %>
		<tr>
          <td height="30" colspan="4">&nbsp;没有输入公司基本信息&nbsp;&nbsp; 【<a href="admin_config.asp?action=addconfig">输入基本信息</a>】</td>
        </tr>
		<% 
		elseif Not(rs.EOF or rs.BOF) then %>
		<tr bgcolor="#E8E8E8">
          <td height="30" colspan="4">&nbsp;公司基本信息查看</td>
        </tr>
        <tr>
          <td width="13%" height="30" align="right" valign="middle">公司名称：</td>
          <td colspan="3" align="left" valign="middle"><%= rs("c_incname") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">公司地址：</td>
          <td width="41%" height="30" align="left" valign="middle"><%= rs("c_addr") %></td>
          <td width="14%" align="right" valign="middle">公司网址：</td>
          <td width="32%" align="left" valign="middle"><%= rs("c_web") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">电话号码： </td>
          <td height="30" align="left" valign="middle"><%= rs("c_tel") %></td>
          <td align="right" valign="middle">电子邮箱：</td>
          <td align="left" valign="middle"><%= rs("c_email") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">传真号码：</td>
          <td height="30" align="left" valign="middle"><%= rs("c_fax") %></td>
          <td align="right" valign="middle">联 系 人：</td>
          <td align="left" valign="middle"><%= rs("c_linkman") %></td>
        </tr>
		<tr bgcolor="#E8E8E8">
          <td height="30" colspan="4">&nbsp;公司基本信息(英文)</td>
        </tr>
		<tr>
          <td width="13%" height="30" align="right" valign="middle">公司英文名称：</td>
          <td colspan="3" align="left" valign="middle"><%= rs("e_incname") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">公司英文地址：</td>
          <td width="41%" height="30" align="left" valign="middle"><%= rs("e_addr") %></td>
          <td width="14%" align="right" valign="middle">公司网址：</td>
          <td width="32%" align="left" valign="middle"><%= rs("e_web") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">电话号码： </td>
          <td height="30" align="left" valign="middle"><%= rs("e_tel") %></td>
          <td align="right" valign="middle">电子邮箱：</td>
          <td align="left" valign="middle"><%= rs("e_email") %></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">传真号码：</td>
          <td height="30" align="left" valign="middle"><%= rs("e_fax") %></td>
          <td align="right" valign="middle">联 系 人：</td>
          <td align="left" valign="middle"><%= rs("e_linkman") %></td>
        </tr>
		<tr align="center" valign="middle">
          <td height="30" colspan="4">【<a href="admin_config.asp?action=editconfig">修改信息</a>】</td>
        </tr>
		<% 
		End if
		rs.close
		set rs=Nothing
		 %>
      </table>
      <br>
<%
if request.QueryString("action")="editconfig" then
sql="select * from config"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td height="30" colspan="4" bgcolor="#E8E8E8">&nbsp;修改公司基本信息</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td width="14%" align="right" valign="middle">公司名称：</td> 
            <td height="30" colspan="3" align="left" valign="middle"><input name="c_incname" type="text" class="textarea" id="c_incname" value="<%=rs("c_incname")%>" size="50" maxlength="70"></td>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="middle">公司地址：</td>
            <td width="40%" height="30" align="left" valign="middle"><input name="c_addr" type="text" id="c_addr" value="<%=rs("c_addr")%>" size="40" maxlength="50"></td>
            <td width="14%" align="right" valign="middle">公司网址：</td>
            <td width="32%" height="30" align="left" valign="middle"><input name="c_web" type="text" id="c_web" value="<%=rs("c_web")%>" size="30" maxlength="60"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="middle">电话号码： </td>
            <td height="30" align="left" valign="middle"><input name="c_tel" type="text" id="c_tel" value="<%=rs("c_tel")%>" size="30" maxlength="50"></td>
            <td align="right" valign="middle">电子邮箱：</td>
            <td height="30" align="left" valign="middle"><input name="c_email" type="text" id="c_email" value="<%=rs("c_email")%>" size="30" maxlength="50"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="middle">传真号码：</td>
            <td height="30" align="left" valign="middle"><input name="c_fax" type="text" id="c_fax" value="<%=rs("c_fax")%>" size="30" maxlength="50"></td>
            <td align="right" valign="middle">联 系 人：</td>
            <td height="30" align="left" valign="middle"><input name="c_linkman" type="text" id="c_linkman" value="<%=rs("c_linkman")%>" size="30" maxlength="50"></td>
          </tr>
		<tr bgcolor="#E8E8E8">
          <td height="30" colspan="4">&nbsp;公司基本信息(英文)</td>
        </tr>
		<tr>
          <td width="13%" height="30" align="right" valign="middle">公司英文名称：</td>
          <td colspan="3" align="left" valign="middle"><input name="e_incname" type="text" class="textarea" id="e_incname" value="<%=rs("e_incname")%>" size="50" maxlength="70"></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">公司英文地址：</td>
          <td width="41%" height="30" align="left" valign="middle"><input name="e_addr" type="text" id="e_addr" value="<%=rs("e_addr")%>" size="40" maxlength="50"></td>
          <td width="14%" align="right" valign="middle">公司网址：</td>
          <td width="32%" align="left" valign="middle"><input name="e_web" type="text" id="e_web" value="<%=rs("e_web")%>" size="30" maxlength="60"></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">电话号码： </td>
          <td height="30" align="left" valign="middle"><input name="e_tel" type="text" id="e_tel" value="<%=rs("e_tel")%>" size="30" maxlength="50"></td>
          <td align="right" valign="middle">电子邮箱：</td>
          <td align="left" valign="middle"><input name="e_email" type="text" id="e_email" value="<%=rs("e_email")%>" size="30" maxlength="50"></td>
        </tr>
        <tr>
          <td height="30" align="right" valign="middle">传真号码：</td>
          <td height="30" align="left" valign="middle"><input name="e_fax" type="text" id="e_fax" value="<%=rs("e_fax")%>" size="30" maxlength="50"></td>
          <td align="right" valign="middle">联 系 人：</td>
          <td align="left" valign="middle"><input name="e_linkman" type="text" id="e_linkman" value="<%=rs("e_linkman")%>" size="30" maxlength="50"></td>
        </tr>
          <tr>
            <td height="30" colspan="4" align="center" bgcolor="#F5F5F5" class="chinese"><input type="submit" name="Submit" value="确定修改" class="button">
                <input type="reset" name="Reset" value="清空重填" class="button">
  [<a href="admin_config.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="action" value="editconfig">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%
	  rs.close
      set rs=Nothing
	  End if
      if request.QueryString("action")="addconfig" then
       %>
      <table align="center" width="98%" border="1" cellspacing="0" cellpadding="4" bordercolor="#C0C0C0" style="border-collapse: collapse">
        <form name="form1" method="post" action="">
          <tr> 
            <td height="30" colspan="4" bgcolor="#E8E8E8">&nbsp;输入公司基本信息</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td width="14%" align="right" valign="middle">公司名称：</td> 
            <td height="30" colspan="3" align="left" valign="middle"><input name="c_incname" type="text" class="textarea" id="c_incname" size="50" maxlength="70"></td>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="middle">公司地址：</td>
            <td width="40%" height="30" align="left" valign="middle"><input name="c_addr" type="text" id="c_addr" size="40" maxlength="50"></td>
            <td width="14%" align="right" valign="middle">公司网址：</td>
            <td width="32%" height="30" align="left" valign="middle"><input name="c_web" type="text" id="c_web" size="30" maxlength="60"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="middle">电话号码： </td>
            <td height="30" align="left" valign="middle"><input name="c_tel" type="text" id="c_tel" size="30" maxlength="50"></td>
            <td align="right" valign="middle">电子邮箱：</td>
            <td height="30" align="left" valign="middle"><input name="c_email" type="text" id="c_email" size="30" maxlength="50"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td align="right" valign="middle">传真号码：</td>
            <td height="30" align="left" valign="middle"><input name="c_fax" type="text" id="c_fax" size="30" maxlength="50"></td>
            <td align="right" valign="middle">联 系 人：</td>
            <td height="30" align="left" valign="middle"><input name="c_linkman" type="text" id="c_linkman" size="30" maxlength="50"></td>
          </tr>
          <tr>
            <td height="30" colspan="4" align="center" bgcolor="#F5F5F5" class="chinese"><input type="submit" name="Submit" value="确定输入" class="button">
                <input type="reset" name="Reset" value="清空重填" class="button">
  [<a href="admin_config.asp">返回</a>] </td>
          </tr>
		  <input type="hidden" name="action" value="addconfig">
          <input type="hidden" name="MM_insert" value="true">
        </form>
      </table>
	  <%End if%>
      <br>
    </td>
  </tr>
  <tr> 
    <td colspan="3" height="1" background="images/dotlineh.gif"></td>
  </tr>
</table>
<% End sub %>
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
