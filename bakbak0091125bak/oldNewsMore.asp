<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp" -->
<!--#include file="intop.asp" -->
<body topmargin="0" leftmargin="0">
<div align="center">
  <center>
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
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG height=29 src="index/Top_Bn_Left.gif" 
                        width=190 border=0></TD></TR>
                    <TR>
                      <TD height=20><!-- //左栏菜单 -->
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD height="22"></TD></TR></TBODY></TABLE>
                        <TABLE cellSpacing=0 cellPadding=0 width="100%" 
border=0>
                          <TBODY>
                          <TR>
                            <TD><img border="0" src="index/fenlei.gif"></TD></TR>
                          <tr>
                            <TD bgColor=#0BBAFD height=2></TD>
                          </tr>
                          <tr>
                            <TD>
                              <TABLE borderColor=#111111 cellSpacing=0 cellPadding=2 width="100%" border=0 class="font22">
              <tr> 
                        <TD width=100% height="8"> 
</td>
              </tr>
              <tr> 
                        <TD width=100% height="16"> 
<script language="JavaScript" type="text/JavaScript">
		             <!--
            function MM_jumpMenu(targ,selObj,restore){ //v3.0
                      eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
                      if (restore) selObj.selectedIndex=0;
                     }
            //-->
<!--
	function ToggleDisplay(oButton, oItems){
		if ((oItems.style.display == "") || (oItems.style.display == "block"))	{
			oItems.style.display = "none";
			oButton.src = "/images/tree_plus.gif";
		}	
		else {
			oItems.style.display = "block";
			oButton.src = "/images/tree_minus.gif";
		}
		return false;
	}
	function UnFoldAll() {
		temp = document.all.length;
	    for(i=0;i<temp;i++)
		    if(document.all[i].id.substr(0,3) =="Img")
			    document.all[i].src = "/images/tree_minus.gif"
	        else if(document.all[i].id.substr(0,4) =="Tree")
		        document.all[i].style.display ="block"
	}
	function FoldAll() {
		temp = document.all.length;
	    for(i=0;i<temp;i++)
		    if(document.all[i].id.substr(0,3) =="Img")
			    document.all[i].src = "/images/tree_plus.gif"
	        else if(document.all[i].id.substr(0,4) =="Tree")
		        document.all[i].style.display ="none"
	}
//-->
</SCRIPT>

<%   
sql = "SELECT * FROM P_class ORDER BY p_id"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.OPEN sql,Conn,0,1
if rs.eof and rs.bof then
Response.Write"还没有分类"
else
do while not rs.eof

sql = "SELECT * FROM P_class_small where p_type ='"&rs("p_type")&"'"
Set rs2 = Server.CreateObject("ADODB.Recordset")
rs2.OPEN sql,Conn,1,1
if rs2.recordcount > 0 then 
Response.Write"&nbsp;<A border=0 style='CURSOR: hand' id=Img"&rs("p_id")&" onclick=ToggleDisplay(Img"&rs("p_id")&",Tree"&rs("p_id")&")><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type")&"</b></a>"
else
Response.Write"&nbsp;<A href=Demo_Product.asp?classtype="&rs("p_type")&" class='six'><img src='images/sortjiutuo.jpg' width='11' height='11' border=0>&nbsp;<b>"&rs("p_type")&"</b></a>"
end if
counts=conn.execute("Select count(*) from P_info where p_other = 0 and p_type='"&rs("p_type")&"'")(0)
Response.Write"(<FONT COLOR='800000'>"&counts&"</FONT>)<br>"
%>
</td>
              </tr>
<tr>
  <td><img border="0" src="index/line.gif"></td>
</tr>
<tr><td>
                          <div id=Tree<%=rs("p_id")%>> 
                            <%
do while not rs2.eof
%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="100%">&nbsp; ◇&nbsp;<a class="number" href='Demo_Product.asp?classtype=<%=rs2("p_type")%>&smallclass=<%=rs2("p_small_type")%>'><%=rs2("p_small_type")%></a></td>                                    
                              </tr>
                            </table>
                            <%
rs2.movenext
loop
rs2.close
Set rs2=Nothing
%>
                            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="100%" height="5"></td>
                              </tr>
                            </table>
                          </div>
                          <%
rs.movenext
loop
end if
rs.close
Set rs=Nothing
%>
<SCRIPT language=javascript>FoldAll();</SCRIPT>
                        </td>
                      </tr>
                    </table></TD>
                          </tr>
                          <tr>
                            <TD height="10">
                            </TD>
                          </tr>
                          <TR>
                            <TD height=3><img border="0" src="index/Select.gif"></TD></TR>
                          <tr>
                            <TD bgColor=#0BBAFD height=2></TD>
                          </tr>
                          <TR>
                            <TD>
                              <DIV align=right>
                              <TABLE cellSpacing=0 cellPadding=0 width=175 
                              border=0>
                                <TBODY>
                                <TR>
                                <TD height=5></TD></TR>
                                <tr>
                          <TD align=middle width="163" background=images/lde.gif height=60> 
                            <DIV align=center>
                                                      <table width="100%" height="34" border="0" cellpadding="0" cellspacing="0">
                                                        <form name="form1" method="post" action="Demo_Product.asp">
                                                          <tr> 
                                                            <td width="110" class="font18"> 
                                                              <input name="key" type="text" class="bd5" id="key2" size="15"></td>
                                                            <td width="65" align="left"><input name="imageField2" type="image" src="images/submit.gif" border="0"></td>
                                                          </tr>
                                                        </form>
                                                      </table>
                                                    </DIV></TD>
                                </tr>
                                </TBODY></TABLE></DIV></TD></TR>
                          </TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
                <TD vAlign=top background=images/Jg_Left_Bg.gif height="318"><IMG 
                  src="images/Jg_Left.gif" border=0 width="14" height="218"></TD>
                <TD vAlign=top width="100%" height=318>
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG src="images/Top_Bn_Foot.gif" border=0></TD></TR>
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
                                src="images/Menu_Bg_Left.gif" width=15 
                                border=0></TD>
                                <TD width="100%" 
                                background=images/Menu_Bg_m.gif>
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
                                src="images/Menu_Bg_Right.gif" width=15 
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
                                <TD width=43><img border="0" src="index/News.gif"></TD>
                                <TD width="100%">
                                <P align=right>&nbsp;</P></TD></TR></TBODY></TABLE></TD>
                                                      </tr>
                                                      <tr>
                                <TD height=5></TD>
                                                      </tr>
                                                      <tr>
                                <TD bgColor=#eaeaea height=1></TD>
                                                      </tr>
                                                      <tr>
                                <TD height=8>
                                  <table border="0" width="536" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td width="534" colspan="2" height="10"><% 
'##############################################
'更多新闻连接，如果是所有的新闻，不用指定newstype的参数，newstype得参数值可在后台管理的新闻类别管理里查看
'news_language=0表示是中文的新闻，news_language=1表示是英文的新闻
'##############################################

newstype=trim(replace(request("newstype"),"'",""))
if newstype="" then
 sql="select news.news_id,news_title,news_date,news_type,news_content from news,news_class where news.news_class_id=news_class.news_id and news_language=0 order by news_date DESC"  
else
 sql="select news.news_id,news_title,news_date,news_type,news_content from news,news_class where news.news_class_id=news_class.news_id and news_language=0 "
 sql=sql & " and news_type = '"&newstype&"' order by news_date DESC"	  
end if
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
 %>

          <%
if Not rs.EOF then
dim newsperpage
newsperpage=6 '每页显现6条新闻
rs.movefirst
rs.pagesize=newsperpage
if trim(request("page"))<>"" then
   page1=trim(request("page"))
       currentpage=clng(request("page"))
       if currentpage>rs.pagecount then
          currentpage=rs.pagecount
       End if
else
   currentpage=1
End if
   totalnews=rs.recordcount
   if currentpage<>1 then
       if (currentpage-1)*newsperpage<totalnews then
	       rs.move(currentpage-1)*newsperpage
		   dim bookmark
		   bookmark=rs.bookmark
	   End if
   End if
   if (totalnews mod newsperpage)=0 then
      totalpages=totalnews\newsperpage
   else
      totalpages=totalnews\newsperpage+1
   End if
i=0
Do While Not rs.EOF and i<newsperpage
%>
          <tr> 
            <td width="435" height="22"><img src=images/none.gif width="1" height="17" border=0><img src=images/i4.gif width="10" height="12" border='0'>&nbsp;[<a href="?newstype=<%=rs("news_type")%>"> 
              <%=rs("news_type")%></a> ]      
			  <a href="Demo_ShowNews.asp?news_id=<%=rs("news_id")%>"> 
              <font color="#CC6600"><%=left(rs("news_title"),30)%></font></a>

            </td>
            <td width="101" height="22"><%=rs("news_date")%> 

            </td>
          </tr>
            <%
i=i+1
rs.movenext
loop
else
if rs.EOF and rs.BOF then
%>
          <tr> 
            <td height="22" align="center" width="534" colspan="2">当前没有新闻！</td>
          </tr>
          <%End if
End if
'开始分页
if Not request("type")=1 then%>
          <form name="form1" method="post" action="Demo_newsmore.asp">
            <tr align="center"> 
              <td height="30" width="534" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;共[<font color="#FF6666"><%=totalnews%></font>]个新闻　分[<font color="#FF6666"><%=totalpages%></font>]页       
                <%
if CurrentPage<2 then
response.write "<font color='000000'>首页 上一页</font> "
else
response.write "<a href=?page=1&newstype="&request("newstype")&">首页</a> "
response.write "<a href=?page="&CurrentPage-1&"&newstype="&request("newstype")&">上一页</a> "
End if
if totalpages-currentpage<1 then
response.write "<font color='000000'>下一页 尾页</font>"
else
response.write "<a href=?page="&CurrentPage+1&"&newstype="&request("newstype")
response.write ">下一页</a> <a href=?page="&totalpages&"&newstype="&request("newstype")&">尾页</a>"
End if
end if
%>
                &nbsp; <select name="page" class="lanyu">      
                  <%      
i=1
for i=1 to totalpages
if i=currentpage then
%>
                  <option value=<%=i%> selected>第<%=i%>页</option>
                  <%else%>
                  <option value=<%=i%>>第<%=i%>页</option>
                  <%
										  End if
next  
 %>
                </select> <input type="submit" name="Submit2" value="转向" class="button">       
                <input type="hidden" name="newstype" value="<%=newstype%>"> </td>      
            </tr>
          </form>
                                  </table>
                                </TD>
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
%>	<table width="535"  border="0" cellpadding="0" cellspacing="0">

</table>                       </TR>
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
</BODY>
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
