<!--#include file="conn.asp"-->
<!--#include file="inc/format.asp"-->
<!--#include file="inc/error.asp"-->
<!--#include file="check1.asp"-->
<!--#include file="inc/pro_body.asp"-->
<title>产品管理</title>
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
 dim sql_test,rs_test 
     sql_test="select * from p_class"
     set rs_test=server.createobject("adodb.recordset")
     rs_test.open sql_test,conn,1,1
	 if rs_test.BOF and rs_test.EOF then
	 rs_test.close
	 set rs_test=Nothing
	 'coon.close
	 'set conn=Nothing
	 founderr=true
     errmsg=errmsg+"<br>"+"<li>你还没有输入产品的分类，不能输入产品！请<a href='admin_p.asp?action=newp' target='_self' >加入新类</a>！"
     Call diserror()
     response.End
	 End if
	 

dim p_name,p_spec,upfile,p_epitome,p_class_id,p_other

if Request.form("MM_insert") then
if request.form("action")="newpro" then

dim evs,evs_sql,evs_rs
evs_sql="select count(p_id) from p_info"
set evs_rs=conn.execute(evs_sql)
if Not(evs_rs.bof or evs_rs.eof) then
   if evs_rs(0)>10000 then
      Response.redirect("admin_pro.asp")
	  response.end()
   end if
   evs_rs.close
   set ev_rs=nothing
end if
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=
sql="select * from p_info"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
p_name=trim(replace(request.form("p_name"),"'",""))
p_spec=trim(replace(request.form("p_spec"),"'",""))
p_pic=trim(replace(request.form("p_pic"),"'",""))
p_pic2=trim(replace(request.form("p_pic2"),"'",""))
p_epitome=trim(replace(request.form("p_epitome"),"'","’"))
p_small_type=trim(replace(request.form("p_small_type"),"'",""))
p_type=trim(replace(request.form("p_type"),"'",""))
p_other=trim(replace(request.form("p_other"),"'",""))
p_jianjie=trim(replace(request.form("p_jianjie"),"'",""))
p_name_e=trim(replace(request.form("p_name_e"),"'",""))
p_spec_e=trim(replace(request.form("p_spec_e"),"'",""))
p_epitome_e=trim(replace(request.form("p_epitome_e"),"'","’"))
p_jianjie_e=trim(replace(request.form("p_jianjie_e"),"'",""))
p_order=trim(replace(request.form("p_order"),"'",""))
if p_other="" then
  rs("p_other")=0
else
  rs("p_other")=1
End if
if p_name="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品名称！"
else
  rs("p_name")=p_name
End if
if p_spec="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的型号！"
else
  rs("p_spec")=p_spec
End if
if p_pic="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的小图！"
else
  rs("p_pic")=p_pic
End if
if p_pic2="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的大图！"
else
  rs("p_pic2")=p_pic2
End if
if p_type="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的大分类！"'else
 else
 rs("p_type")=p_type
End if
'由P_type,p_small_type 得到p_type_e和p_small_tyPe_e
			sql_p="select p_type_e,p_small_type_e from p_class_small where p_type='"&p_type&"' and p_small_type='"&p_small_type&"'" 
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			   p_type_e=rs_p("p_type_e")
			   p_small_type_e=rs_p("p_small_type_e")
			end if
			rs_p.close
			set rs_p=nothing  
			
rs("p_small_type")=p_small_type
rs("p_epitome")=p_epitome '详细说明
rs("p_jianjie")=p_jianjie'简单说明
'英文方面
rs("p_name_e")=p_name_e
rs("p_spec_e")=p_spec_e
rs("p_epitome_e")=p_epitome_e
rs("p_small_type_e")=p_small_type_e
rs("p_type_e")=p_type_e
rs("p_jianjie_e")=p_jianjie_e
rs("p_order")=p_order

if Not founderr then
rs.update
rs.close
set rs=Nothing
  closedatabase
Response.write "<script language = 'javascript'>alert('成功添加一个产品!');window.document.location.href='admin_pro.asp';</script>"
response.end
else
Call diserror()
response.End
End if
End if
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=
if request.form("action")="editpro" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from p_info where p_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
p_name=trim(replace(request.form("p_name"),"'",""))
p_spec=trim(replace(request.form("p_spec"),"'",""))
'upfile=trim(replace(request.form("upfile"),"'",""))
p_pic=trim(replace(request.form("p_pic"),"'",""))
p_pic2=trim(replace(request.form("p_pic2"),"'",""))
p_epitome=trim(replace(request.form("p_epitome"),"'","’"))
p_small_type=trim(replace(request.form("p_small_type"),"'",""))
p_type=trim(replace(request.form("p_type"),"'",""))
p_other=trim(replace(request.form("p_other"),"'",""))
p_jianjie=trim(replace(request.form("p_jianjie"),"'",""))
p_order=trim(replace(request.form("p_order"),"'",""))
p_name_e=trim(replace(request.form("p_name_e"),"'",""))
p_spec_e=trim(replace(request.form("p_spec_e"),"'",""))
p_epitome_e=trim(replace(request.form("p_epitome_e"),"'","’"))
p_small_type_e=trim(replace(request.form("p_small_type_e"),"'",""))
p_type_e=trim(replace(request.form("p_type_e"),"'",""))
p_jianjie_e=trim(replace(request.form("p_jianjie_e"),"'",""))

if p_other="" then
  rs("p_other")=0
else
  rs("p_other")=1
End if
if p_name="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品名称！"
else
  rs("p_name")=p_name
End if
if p_spec="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的型号！"
else
  rs("p_spec")=p_spec
End if
if p_pic="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的小图！"
else
  rs("p_pic")=p_pic
End if
if p_pic2="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的大图！"
else
  rs("p_pic2")=p_pic2
End if
if p_type="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须输入产品的大分类！"'else
 else
 rs("p_type")=p_type
End if
'由P_type,p_small_type 得到p_type_e和p_small_tyPe_e
			sql_p="select p_type_e,p_small_type_e from p_class_small where p_type='"&p_type&"' and p_small_type='"&p_small_type&"'" 
            set rs_p=server.createobject("adodb.recordset")
            rs_p.open sql_p,conn,1,1
			if Not(rs_p.BOF and rs_p.EOF) then 
			   p_type_e=rs_p("p_type_e")
			   p_small_type_e=rs_p("p_small_type_e")
			end if
			rs_p.close
			set rs_p=nothing   
			    
rs("p_small_type")=p_small_type
rs("p_epitome")=p_epitome '详细说明
rs("p_jianjie")=p_jianjie'简单说明
rs("p_order")=p_order
'英文方面
rs("p_name_e")=p_name_e
rs("p_spec_e")=p_spec_e
rs("p_epitome_e")=p_epitome_e
rs("p_small_type_e")=p_small_type_e
rs("p_type_e")=p_type_e
rs("p_jianjie_e")=p_jianjie_e

if Not founderr then
rs.update
rs.close
set rs=Nothing
Response.write "<script language = 'javascript'>alert('成功修改了此产品!');window.document.location.href='admin_pro.asp';</script>"
response.end
else
Call diserror()
response.End
End if
End if
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++=
if request.form("action")="delpro" then
if request.Form("id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你必须指定操作的对象！"
else
  if Not isInteger(request.form("id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的id参数。"
  End if
End if
if founderr then
  Call diserror()
  response.End
End if
sql="select * from p_info where p_id="&cint(request.Form("id"))
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.delete
rs.close
set rs=Nothing
  closedatabase
Response.write "<script language = 'javascript'>alert('成功删除此产品!');window.document.location.href='admin_pro.asp';</script>"
response.end
End if
End if

Call adminpro_body()
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
