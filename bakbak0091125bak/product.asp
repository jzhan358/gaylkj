<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<HTML>
<HEAD>
<TITLE>广州市广安娱乐电子有限公司-产品浏览-【广安娱乐科技】全程用心！一流服务！</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2900.2802" name=GENERATOR>
<meta name="Keywords" content="产品浏览,在线下单,公司动态,广安娱乐,游戏机生产,游戏机销售,电脑游戏机配件,超级机板王,卡带,游艺机,百家乐,明星97,老虎机" />
<meta name="Description" content="产品浏览在线下单广州市广安娱乐电子有限公司主要是一家从事多年游戏机生产、销售、电脑游戏机配件，集科、工、贸于一体的专业公司，并与台湾资深电玩工程师，合作开发多款新产品。" />
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
	<%

if not isempty(request("pageno")) or request("pageno")<>0 then
pageno=request.querystring("pageno")
else
pageno=1
end if
if request.QueryString("pro_type") <> "" then
pro_type=checkstr(request.QueryString("pro_type"))
else
pro_type=""
end if
page=12
pagesize=page
st=""
k=checkStr(request.QueryString("k"))
if len(k)>1 and len(k)<20 then
	st=" and p_name like '%"&k&"%'"
end if

if len(pro_type)>2 and len(pro_type)<20 then
st=st&" and  p_type='"&pro_type&"'"
end if

sql="select * from P_info where p_id>0 "&st&" order by p_id desc"
'response.Write(sql)
'response.End()
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1

%>    
    <td height="400" valign="top">
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="60"><img src="images/main_pro.jpg" width="565" height="45"></td>
      </tr>
      <tr>
        <td>
		<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0"><%
		if not rs.eof then
rs.pagesize=page
maxpage=rs.pagecount
if pageno < 1 then
pageno =1
end if
if pageno+1 >maxpage+1  then
pageno=maxpage
end if
rs.absolutePage=pageno
else
maxpage=1
end if
total = rs.RecordCount
	   %>         
		  <tr>
    <td valign="top">
<table width="93%"  border="0" align="center" cellpadding="0" cellspacing="0">
<tr align="center">
      <% i=1
	  do while not rs.eof and page>0 %>	  
        <td width="190">
		<table  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td>
			<table border="0"  cellpadding="0" cellspacing="0" class="rtlb-green">
              <tr>
                <td width="5" height="5"></td>
                <td align="center" valign="middle"></td>
                <td align="center" valign="middle"></td>
              </tr>
              <tr>
                <td align="center" valign="middle"></td>
                <td align="center" valign="middle"><a href="product_view.asp?id=<%=rs("p_id")%>"><img src="<%= rs("p_pic") %>" width="158" height="120" border="0"></a></td>
                <td valign="middle"></td>
              </tr>
              <tr>
                <td align="center" valign="middle"></td>
                <td align="center" valign="middle"></td>
                <td width="5" height="5"></td>
              </tr>
            </table>
			</td>
          </tr>
		 <tr>
            <td height="35" align="center" class="title01"><a href="product_view.asp?id=<%=rs("p_id")%>"><%= rs("p_name") %></a></td>
          </tr>
        </table></td>
		<%
		  rs.movenext
		  page=page-1
		  if i mod 3 =0 then response.Write("</tr><tr align=""center"">")
		  i=i+1
		loop
		   %>
       
      </tr>
	    <%
			rs.close
			set rs=nothing
		  %>
    </table>
	
	</td>
  </tr>
  <tr>
  <td height="25" valign="bottom">
  <table width="100%" height="26" border="0" align="center" cellpadding="0" cellspacing="0">
                         
                          <tr class="font">
                            <td width="3"></td>
                            <td width="136" class="zhengwen"><div align="center">共有<%=total%>条记录</div></td>
                            <td width="301" class="zhengwen"><div align="center"> 共有<%=maxpage%>页 <%=pagesize%>条/页 目前第<%=pageno%>页 　<a href="product.asp?pageno=<%=pageno-1%>&pro_type=<%=pro_type%>">上一页</a> 　<a href="product.asp?pageno=<%=pageno+1%>&pro_type=<%=pro_type%>">下一页</a></div></td>
                          </tr>
                          <tr>
                            <td></td>
                            <td height="13"></td>
                            <td></td>

                          </tr>
                      </table></td></tr>
        </table></td>
      </tr>
    </table>
</td>
  </tr>
</table>

</td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="rl-orange">
  <tr>
    <td><!--#include file="foot.asp"--></td>
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
