<!--#include file="top.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim founderr,errmsg
founderr=false
i=0
errmsg=""
dim p_name,p_name_e,p_spec,p_spec_e,p_type,p_type_e,p_type_id,p_order,p_pic,p_pic2,p_jianjie,p_jianjie_e,id
dim action,rs
p_name=checkstr(request.Form("p_name"))
p_name_e=checkstr(request.Form("p_name_e"))
p_spec=checkstr(request.Form("p_spec"))
p_spec_e=checkstr(request.Form("p_spec_e"))
p_type_id=request.Form("p_type_id")
p_order=request.Form("p_order")
p_pic=checkstr(request.Form("p_pic"))
p_pic2=checkstr(request.Form("p_pic2"))
p_jianjie=request.Form("p_jianjie")
p_jianjie_e=request.Form("p_jianjie_e")
id=request.Form("id")
action=request.Form("action")

if action="new" or action="edit" then
	if not chkRange(p_name,2,50) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须输入产品的中文标题！"	
	End if
	if not chkRange(p_name_e,2,150) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须输入产品的英文标题！"	
	End if
	if not chkRange(p_spec,2,30) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须输入产品的中文型号！"	
	End if
	if not chkRange(p_spec_e,2,30) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你必须输入产品的英文型号！"	
	End if
	if not chkRange(p_pic,10,150) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你没有上传产品图片！"	
	End if	
	if not chkrequest(p_type_id) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你没有选择产品的分类！"	
	else
		sql="select p_type,p_type_e from p_class where p_id="&p_type_id
		set rs=server.CreateObject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			founderr=true
			i=i+1
			errmsg=errmsg&i&"),你没有正确选择产品的分类！"
		else
			p_type=rs("p_type")
			p_type_e=rs("p_type_e")
		end if
		rs.close
		set rs=nothing
	end if
	if chkNull(p_jianjie,10) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你没有输入产品的中文介绍！"	
	End if
	if chkNull(p_jianjie_e,10) then
		founderr=true
		i=i+1
		errmsg=errmsg&i&"),你没有输入产品的英文介绍！"	
	End if	
end if

if action="edit" or action="del" then
	if not chkrequest(id) then alert "error","",1
end if

if founderr then alert errmsg,"",1

if action="new" then
	sql="select * from p_info"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	rs("p_date")=now()
	rs("p_name")=p_name
	rs("p_name_e")=p_name_e
	rs("p_spec")=p_spec
	rs("p_spec_e")=p_spec_e
	rs("p_order")=p_order
	rs("p_pic")=p_pic
	rs("p_pic2")=p_pic2
	rs("p_type_id")=p_type_id
	rs("p_type")=p_type
	rs("p_type_e")=p_type_e
	rs("p_jianjie")=replace_text(p_jianjie)
	rs("p_jianjie_e")=replace_text(p_jianjie_e)
	rs.update	
	id=rs("p_id")
	if err then
		response.Write(err.description)
		response.End()
	end if
	rs.close
End if

if action="edit" then
	sql="select * from p_info where p_id="&id
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not (rs.bof and rs.eof) then
		rs("p_name")=p_name
		rs("p_name_e")=p_name_e
		rs("p_spec")=p_spec
		rs("p_spec_e")=p_spec_e
		rs("p_order")=p_order
		rs("p_pic")=p_pic
		rs("p_pic2")=p_pic2
		rs("p_type_id")=p_type_id
		rs("p_type")=p_type
		rs("p_type_e")=p_type_e
		rs("p_jianjie")=replace_text(p_jianjie)
		rs("p_jianjie_e")=replace_text(p_jianjie_e)
		rs.update	
	end if
	rs.close
End if
if (action="new" or action="edit") and cityviewmode=1 then
	
	tarr=savetohtml(getproduct_viewpage(id),"/products/"&id&".html")
	tarr2=savetohtml(getindexpage(),"/index.html")

	tarr=savetohtml(geteproduct_viewpage(id),"/eproducts/"&id&".html")
	tarr2=savetohtml(getdefaultpage(),"/default.html")
	call createSeoPage()
	if tarr(0)<>0 then alert tarr(1),"",1
end if
if action="del" then

sql="select p_pic,p_pic2,p_jianjie,p_jianjie_e from p_info where p_id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if Not(rs.BOF or rs.EOF) then
	if not chkNull(rs("p_pic"),10) then call delfile(rs("p_pic"))
	if not chkNull(rs("p_pic2"),10) then call delfile(rs("p_pic2"))
	if not chkNull(rs("p_jianjie"),10) then call delContentPic(rs("p_jianjie"))
	if not chkNull(rs("p_jianjie_e"),10) then call delContentPic(rs("p_jianjie_e"))
rs.delete
call delfile("/products/"&id&".html")
call delfile("/eproducts/"&id&".html")
if cityviewmode=1 then
	call savetohtml(getindexpage(),"/index.html")
	call savetohtml(getdefaultpage(),"/default.html")
end if
call createSeoPage()
End if
rs.close
End if
set rs=nothing
closeconn
response.Redirect("products.asp")
%>