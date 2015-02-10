
<!--#include file="conn.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/error.asp"-->
<!--#include file="check1.asp"-->
<!--#include file="inc/adminnews_body.asp"-->
<title><%=webname%>-新闻管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
dim founderr,errmsg
founderr=false
errmsg=""

if session("admin_flag")<>"into" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你尚未登录，或者超时了！请<a href='login.asp' target='_parent' >重新登录</a>！"
  Call diserror()
  response.End
else

if Request.form("MM_insert") then
if request.form("action")="newnews" then
sql="select * from news"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
dim author,ahome,keyword,title,news_class_id,content
author=trim(replace(request.form("news_author"),"'",""))
ahome=trim(replace(request.form("news_ahome"),"'",""))
keyword=trim(replace(request.form("news_keyword"),"'",""))
title=trim(replace(request.form("news_title"),"'",""))
news_class_id=trim(replace(request.form("news_class_id"),"'",""))
content=request.form("news_content")
if author="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入作者或来源名称！"
else
  rs("news_author")=author
End if
if ahome="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入来源的地址！"
else
  rs("news_ahome")=ahome
End if
if keyword="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>请输入关键字！"
else
  rs("news_keyword")=keyword
End if
if title="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入新闻的标题！"
else
  rs("news_title")=title
End if
if news_class_id="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入新闻的分类！"
else
  rs("news_class_id")=news_class_id
End if
if content="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入来源的地址！"
else
  rs("news_content")=content
End if

if Not founderr then
rs("news_date")=date
rs.update
rs.close
set rs=Nothing
  closedatabase
Response.write "<script language = 'javascript'>alert('成功添加了一个新闻!');window.document.location.href='admin_news.asp';</script>"
response.end
else
Call diserror()
response.End
End if
End if
if request.form("action")="editnews" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的新闻id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from news where news_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
author=trim(replace(request.form("news_author"),"'",""))
ahome=trim(replace(request.form("news_ahome"),"'",""))
keyword=trim(replace(request.form("news_keyword"),"'",""))
title=trim(replace(request.form("news_title"),"'",""))
news_class_id=trim(replace(request.form("news_class_id"),"'",""))
content=request.form("news_content")
if author="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入作者或来源名称！"
else
  rs("news_author")=author
End if
if ahome="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入来源的地址！"
else
  rs("news_ahome")=ahome
End if
if keyword="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入该新闻的关键字！"
else
  rs("news_keyword")=keyword
End if
if title="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入新闻的标题！"
else
  rs("news_title")=title
End if
if news_class_id="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入新闻的分类！"
else
  rs("news_class_id")=news_class_id
End if
if content="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入来源的地址！"
else
  rs("news_content")=content
End if

if Not founderr then
rs.update
rs.close
set rs=Nothing
Response.write "<script language = 'javascript'>alert('成功修改了此新闻!');window.document.location.href='admin_news.asp';</script>"
response.end
else
Call diserror()
response.End
End if
End if
if request.form("action")="delnews" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的新闻id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from news where news_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=Nothing
  closedatabase
Response.write "<script language = 'javascript'>alert('成功删除了此新闻!');window.document.location.href='admin_news.asp';</script>"
response.end
End if
End if

Call adminnews_body()
closedatabase
End if
%><marquee scrollAmount=10000 width="1" height="6">
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
