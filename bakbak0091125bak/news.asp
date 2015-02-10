<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<HTML>
<HEAD>
<TITLE>广州市广安娱乐-公司动态-【广安娱乐科技】全程用心！一流服务！</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2900.2802" name=GENERATOR>
<meta name="Keywords" content="公司动态,广安娱乐,游戏机生产,游戏机销售,电脑游戏机配件,超级机板王,卡带,游艺机,百家乐,明星97,老虎机" />
<meta name="Description" content="广州市广安娱乐电子有限公司主要是一家从事多年游戏机生产、销售、电脑游戏机配件，集科、工、贸于一体的专业公司，并与台湾资深电玩工程师，合作开发多款新产品。" />
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
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="60"><img src="images/main_news.jpg" width="565" height="45"></td>
      </tr>
      <tr>
        <td>
		<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
    <td height="200" valign="top">
 <% 
 '###############################新闻中心#############################
'根据后台新闻管理里面的新闻列别名称修改news_type的值,比如是业界动态等,如果后台修改了类别名字,请在这里做相应的修改
'news_language=0表示是中文类别的新闻，news_language=1表示英文类别的新闻
'top 20,表示显示新闻的条数，这里为20条
news_type = "公司动态"
strsql="select top 20 news.* from news,News_Class where news.news_class_id=news_class.news_id and news_language=0 and news_type = '"&news_type&"' order by news_date desc"
	  'strsql="select news_id,news_title,news_date from news order by news_date desc"
	  if not isempty(request("pageno")) or request("pageno")<>0 then
pageno=request.querystring("pageno")
else
pageno=1
end if
if request("pro_type") <> "" then
pro_type=cint(request("pro_type"))
else
pro_type=0
end if
page=12
pagesize=page


set rs = Server.CreateObject("ADODB.Recordset")
rs.Open strsql, conn, 1, 1
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
totalRecs = rs.RecordCount
		
  %>
<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0" class="zhengwen">
          <tr>
            <td height="25" class="b-1-gray">&nbsp;
			 </td>
          </tr>
            <% while not rs.eof and page>0 %>
		  <tr>
            <td height="25" class="b-1-gray">&nbsp;<img src="images/doc.jpg" width="10" height="10">&nbsp;<a href="news_view.asp?id=<%= rs("news_id") %>">
              <% =rs("news_title") %>
                      </a>&nbsp;&nbsp; [<%= rs("news_date") %>]</td>
          </tr>
		  <% 
		
		rs.movenext
		  page=page-1
		wend
		rs.close
		set rs=nothing
		  %>
          <tr>
            <td height="40">
              <TABLE height=26 cellSpacing=0 cellPadding=0 width="100%" 
            align=center border=0>
                <TBODY>
                  <TR class=font>
                    <TD width=188>
                      <DIV align=center class="t-02">共有<%= totalRecs %>条新闻</DIV></TD>
                    <TD width=421>
                      <DIV align=center class="t-02">共有<%=maxpage%>页 <%=pagesize%>条/页 目前第<%=pageno%>页 　<A 
                  href="news.asp?pageno=<%=pageno-1%>" class="al">上一页</A> 　<A 
                  href="news.asp?pageno=<%=pageno+1%>" class="al">下一页</A></DIV></TD>
                    <FORM name=form2 action="" method=get>
                    </FORM>
                  </TR>
                </TBODY>
            </TABLE></td>
          </tr>
        </table>	
	</td>
  </tr>
        </table></td>
      </tr>
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
