<% 
	  strsql="select * from P_class order by p_id desc"
	  set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open strsql, conn, 1, 1
		totalRecs = rs.RecordCount
	   %>

<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">

      <tr>
        <td height="28" align="center" class="green"> 
		
<!--��ʼ-->
<table border="0" bgcolor="#C0EDB4">
<tr><td nowrap="nowrap" valign="top" align="left" height="22">
<strong>�㰲���ֵ��ӿƼ���˾ </strong>
</td></tr>
<tr>
<td nowrap="nowrap">
<strong>��ϵ�绰:</strong>020��39977130
</td></tr>
<tr>
<td nowrap="nowrap">
<strong>�ֻ�����:</strong>13719381591<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;13660170912
</td></tr>
<tr>
<td nowrap="nowrap">
<strong>��&nbsp;ϵ&nbsp;��&nbsp;:</strong>������
</td></tr>
<tr>
<td nowrap="nowrap">
<strong>�������:</strong>020��39977130
</td></tr>
<tr>
<td nowrap="nowrap">
<strong>QQ:</strong><a href="http://wpa.qq.com/msgrd?V=1&Uin=276033090&Site=�㰲���ֵ��ӿƼ���˾&Menu=yes" title="����������ҽ�̸" target="_blank">276033090</a>   <a href="http://wpa.qq.com/msgrd?V=1&Uin=841492922&Site=�㰲���ֵ��ӿƼ���˾&Menu=yes" title="����������ҽ�̸" target="_blank">841492922</a>
</td></tr>
</table>
<!--����-->
</td>
      </tr>
	  <tr><td height="10">
<!-- SiteSearch Google -->   
		</td>
      </tr>
      <tr>
        <td><form action="product.asp" method="get"><table width="168"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><img border="0" src="/images/i4.gif" />��Ʒ������
              <input type="text" name="k" style="width:100px; border:1px #CCCCCC solid; margin:0px; padding:0px" />
              <input type="image" src="/images/go.gif" /></td>
          </tr>          
        </table>
        </form></td>
      </tr>
<tr><td height="10">
<!-- SiteSearch Google -->   
		</td>
      </tr>
      <tr>
        <td><table width="168"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="images/left_pro.jpg" width="168" height="33"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="rl-green">
			<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
   <% while not rs.eof %> 
  <tr>
    <td height="26" class="b-1-gray"> <img src="images/doc.jpg" width="10" height="10"> <span class="title05"><a href="product.asp?pro_type=<%=rs("p_type")%>"><font color="#333333"><%= rs("p_type") %></font></a><a href="product.asp?pro_type=<%=rs("P_type")%>"></a></span></td>
  </tr>
  		  <% rs.movenext
		  wend
		  rs.close		  
		   %>	
</table>
</td>
          </tr>
          <tr>
            <td><img src="images/left_pro_foot.jpg" width="168" height="14"></td>
          </tr>
        </table></td>
      </tr>  
      <tr>
        <td height="2"></td>
      </tr>
      <tr>
        <td align="center"><a href="gbook.asp"><img src="images/left_gbook.jpg" width="168" height="66" border="0"></a></td>
      </tr>
	     <tr>
        <td height="2"></td>
      </tr>
      <tr>
        <td align="center">
		<!--news start-->
		<table width="168"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><img src="images/left_news.jpg" width="168" height="33"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF" class="rl-green">
			<table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
			<%
			news_type = "��˾��̬"
	sql="select top 10 news.* from news,News_Class where news.news_class_id=news_class.news_id and news_language=0 and news_type = '"&news_type&"' order by news.news_id desc"
			rs.open sql,conn,1,1
			do while not rs.eof %> 
			<tr><td height="26" class="b-1-gray">  &nbsp;<img src="images/doc.jpg" width="10" height="10">&nbsp;<a href="news_view.asp?id=<%=rs("news_id")%>"> <%= left(rs("news_title"),10) %></a></td>
			</tr>
			<%
			rs.movenext
			loop
			rs.close
			set rs=nothing
			%>
			
			</table>
			</td>
			</tr>
			 <tr>
            <td><img src="images/left_pro_foot.jpg" width="168" height="14"></td>
          </tr>
			</table>
			<!--news end-->
			</td>
      </tr>
	    <tr>
        <td height="2"></td>
      </tr>
    </table>
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
