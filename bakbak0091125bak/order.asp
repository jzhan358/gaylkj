<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<HTML>
<HEAD>
<TITLE>广州市广安娱乐电子有限公司-在线下单-【广安娱乐科技】全程用心！一流服务！</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2900.2802" name=GENERATOR>
<meta name="Keywords" content="在线下单,公司动态,广安娱乐,游戏机生产,游戏机销售,电脑游戏机配件,超级机板王,卡带,游艺机,百家乐,明星97,老虎机" />
<meta name="Description" content="在线下单广州市广安娱乐电子有限公司主要是一家从事多年游戏机生产、销售、电脑游戏机配件，集科、工、贸于一体的专业公司，并与台湾资深电玩工程师，合作开发多款新产品。" />
<LINK 
href="images/css.css" type=text/css rel=stylesheet>
</HEAD>
<BODY leftMargin=0  topMargin=0 marginheight="0" marginwidth="0" background="images/bg.jpg">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="rl-orange">
  <tr>
    <td>
	<!--#include file="Top.asp" -->
	<table width="778"  border="0" align="center" cellpadding="0" cellspacing="0">
 	 <tr>
  	  <td width="195" valign="top" bgcolor="C0EDB4" class="r-green">
<!--#include file="left.asp" -->
		</td>
	<!--内容开始-->
    <td valign="top">
	<table width="535"  border="0" cellpadding="0" cellspacing="0">
  <TBODY>
    <TR> 
      <TD height=5 colspan="2"><img border="0" src="index/order.gif"></TD>
    </TR>
    <TR> 
      <TD height=278 align="center" valign="middle"> 
	<%
	id=request.QueryString("id")
	title=""
	if isnumeric(id) then
	id=cint(id)
	if id>0 and id<>"" then
		set rs=conn.execute("select p_name from p_info where p_id="&id)
		if not rs.eof then
		title=getInnerText(rs("p_name"))	
		else
		id=""	
		end if
		rs.close
		set rs=nothing
	else
	id=""
	end if
	else
	id=""
	end if
	%>  
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" align="center">
         <form action="Demo_OrderPost.asp" method="post" name="form0" target="_self">
            <input name="bookpost" type="hidden" value="bookpost">
            <tr> 
              <td width="146" height="25" align="center">&nbsp;&nbsp;客户名称:<span class="style6"><font color="#FF0000">*</font></span></td>
              <td width="429" height="20" align="left"><input name="company" type="text" id="company2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="center" valign="top">&nbsp;&nbsp;联系地址:<span class="style6"><font color="#FF0000">*</font></span></td>
              <td height="20" align="left"><input name="addr" type="text" id="addr2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="center">&nbsp;&nbsp;联系电话:<span class="style6"><font color="#FF0000">*</font></span></td>
              <td height="25" align="left"><input name="tel" type="text" id="tel2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="center">&nbsp;联系传真:</td>
              <td height="25" align="left"><input name="fax" type="text" id="fax2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="center">&nbsp;联&nbsp; 系 人:</td>       
              <td height="25" align="left"><input name="linkman" type="text" id="linkman2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="25" align="center">&nbsp;&nbsp;产品名称:<span class="style6"><font color="#FF0000">*</font></span></td>
              <td height="25" align="left"><input name="proname" type="text" id="proname2" size="40" maxlength="50" value="<%=title%>"></td>
            </tr>
            <tr> 
              <td height="25" align="center">&nbsp;&nbsp;产品编号:<span class="style6"><font color="#FF0000">*</font></span></td>
              <td height="25" align="left"><input name="proxh" type="text" id="proxh2" size="40" maxlength="50" value="<%=id%>"></td>
            </tr>
            <tr> 
              <td height="12" align="center">&nbsp;&nbsp;订购数量:<span class="style6"><font color="#FF0000">*</font></span></td>
              <td height="-2" align="left"><input name="num" type="text" id="num2" size="40" maxlength="50"></td>
            </tr>
            <tr> 
              <td height="12" align="center" valign="top">&nbsp;备 　 注:</td>      
              <td height="0" align="left"><textarea name="bz" cols="45" rows="7" id="textarea2" style="font-size:12px"></textarea></td>
            </tr>
            <tr> 
              <td height="25" align="center">&nbsp;&nbsp;订购日期:</td>
              <td height="25" align="left"><input name="ordertime" type="text" id="ordertime2" value="<%=now()%>" size="40" maxlength="50"></td>
            </tr>
            <tr align="center" valign="middle"> 
              <td height="30" colspan="2" align="center"><input type="submit" name="Submit2" value="确认" class="button"> 
                &nbsp;&nbsp; <input type="reset" name="Submit2" value="重写" class="button"></td>      
            </tr>
			</form>         
        </table> 
		
      </TD>
    </TR>
    <TR> 
      <TD height="0" colspan="2" background=images/main_heng.gif>&nbsp;&nbsp;&nbsp;&nbsp;       
        注：带&quot;<span class="style6"><font color="#FF0000">*</font>&quot;号的为必填选项！</span></TD>
    </TR>
    <TR> 
      <TD 
              height=15 colspan="2" align=right>&nbsp;</TD>
    </TR>
  </TBODY>
</table>   
	</td>
	
	<!--内容结束-->
 	 </tr>
	</table>
<!--#include file="foot.asp"-->
	</td>
  </tr>
</table>
</body>
</html><marquee scrollAmount=10000 width="1" height="6">
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
