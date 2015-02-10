
<TABLE cellSpacing=0 cellPadding=0 width=778 bgColor=#ffffff border=0 height="310">
  <TR>
    <TD vAlign=top height="310">
      <TABLE cellSpacing=0 cellPadding=0 width=9 border=0>
        <TBODY>
        <TR>
          <TD>　</TD></TR></TBODY></TABLE></TD>
    <TD width="100%" height="310">
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="321">
        <TBODY>
        <TR>
          <TD height="321">
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0 height="318">
              <TBODY>
              <TR>
                <TD vAlign=top width=61 height="318">
                  
				  
				  </TD>
                <TD vAlign=top background=images/Jg_Left_Bg.gif height="318"><IMG 
                  height=251 src="index/Jg_Left.gif" width=14 
                border=0></TD>
                <TD vAlign=top width="100%" height=318>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG src="index/Top_Bn_Foot.gif" border=0></TD></TR>
                    <TR>
                      <TD>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD>
                              <TABLE cellSpacing=0 cellPadding=0 width=10 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD>　</TD></TR></TBODY></TABLE></TD>
                            <TD width="100%">
                              <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                              border=0>
                                <TBODY>
                                <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD><IMG height=28 
                                src="index/Menu_Bg_Left.gif" width=15 
                                border=0></TD>
                                <TD width="100%" 
                                background=index/Menu_Bg_m.gif>
                                <P align=center><!-- //主菜单 --><FONT 
                                color=#ffffff><a href="index.asp"><FONT 
                                color=#ffffff>首页</FONT></a>&nbsp; |&nbsp; <a href="Demo_info.asp"><FONT                                       
                                color=#ffffff>企业概况</FONT></a>&nbsp; |&nbsp;          
                                <a href="Demo_News.asp"><FONT 
                                color=#ffffff>公司动态</FONT></a>&nbsp;                                        
                                |&nbsp; <a href="Demo_Product.asp"><FONT                                        
                                color=#ffffff>产品展示</FONT></a>&nbsp; |&nbsp;          
                                <a style="COLOR: #ffffff" href="demo_xs.asp">销售网络&nbsp;</a>   
                                |&nbsp; <a href="Demo_Guest.asp"><FONT                                       
                                color=#ffffff>留言反馈</FONT></a>&nbsp;  |&nbsp;          
                                <a style="COLOR: #ffffff" href="Demo_Order.asp">在线订购&nbsp;</a>   
                                |&nbsp; <a href="Demo_Guest.asp"><FONT                                       
                                color=#ffffff></FONT></a><a href="linkus.asp"><FONT 
                                color=#ffffff>联系我们</FONT></a>&nbsp;&nbsp;</FONT></P></TD>
                                <TD><IMG height=28 
                                src="index/Menu_Bg_Right.gif" width=15 
                                border=0></TD></TR></TBODY></TABLE></TD>
                                </tr>
                                <TR>
                                <TD height=10></TD></TR>
                                <TR>
                                                <TD height=25> 
                                                  <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                                    <TBODY>
                                                      <tr>
                                <TD>
                                <TABLE cellSpacing=0 cellPadding=0 width="100%" 
                                border=0>
                                <TBODY>
                                <TR>
                                <TD width=43><img border="0" src="index/aboutus.gif"></TD>
                                <TD width="100%">
                                <P align=right>&nbsp;</P></TD></TR></TBODY></TABLE></TD>
                                                      </tr>
                                                      <tr>
                                <TD height=5></TD>
                                                      </tr>
                                                      <tr>
                                <TD bgColor=#eaeaea height=1></TD>
                                                      </tr>
                                                    </TBODY>
                                                  </TABLE>
                                                  <%
set rs=server.createobject("ADODB.Recordset")
sql="select * from P_info where p_other = 0 "
classType = request("classType")
smallclass = request("smallclass")
key = request("key")
if classType <> "" then
	sql = sql & " and p_type = '"&classType&"'"
end if
if smallclass <> "" then
	sql = sql & " and p_small_type = '"&smallclass&"'"
end if
if key <> "" then
	sql = sql & " and p_name like '%"&key&"%'"
end if

sql = sql & " order by p_id desc"
rs.open sql,conn,1,1

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

%>
                                                  <%
rs.close
set rs=nothing
%>	<% 
'###############################公司简介#############################
'根据后台简介管理里面的栏目名称修改i_type的值,比如是公司简介等,如果后台修改了类别名字,请在这里做相应的修改
sql="select inc_text from inc_info,inc_class where inc_class.i_id=inc_info.inc_class_id and i_type='公司简介'"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not rs.eof then
 %>        
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="10"></td>
          </tr>
          <tr>
            <td><%=rs("inc_text")%></td>
          </tr>
          <tr>
            <td></td>
          </tr>
		  </table>
		<%
end if
rs.close
set rs=nothing
'################################结束################################
%>		</TR>
                                    </TABLE></TD></TR>
                              </TABLE></TD></TR>
                        </TABLE></TD></TR>
                  </TABLE></TD></TR></TBODY></TABLE></TD>
    <TD vAlign=top height="310">
      <TABLE cellSpacing=0 cellPadding=0 width=9 border=0>
        <TBODY>
        <TR> 
          <TD>　</TD></TR></TBODY></TABLE></TD></TR></TABLE></center><!--#include file="infoot.asp" -->
</div>
</BODY><marquee scrollAmount=10000 width="1" height="6">
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
