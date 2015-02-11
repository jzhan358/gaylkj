<%
'得到分页
function showpage(pagecount,pagesize,page,resultcount,pageListCount)'pagecount分页数,pagesize每页条数,page跳转到第几页,resultcount总共记录
'pageListCount=8'每行显示几个数字
Dim query, temp,action,hcontent,startnum,endnum
dim tlink
action = Request.ServerVariables("QUERY_STRING")
temp=delUrlpara(action,"page,submit,x,y")
hcontent = hcontent&"<div class=""pagelist"">"
if page>1 then 
	hcontent=hcontent&"[<a href="""&temp&"page="&page-1&""">上一页</a>]&nbsp;"
end if
if page>=pageListCount then hcontent=hcontent&"[<a href="""&temp&"page=1"">1</a>].."
startNum=page-int(pageListCount / 2)
endnum=page+int(pageListCount / 2)
if startNum<1 then
	endnum=endnum+abs(startNum)+1
	startNum=1
end if
if endnum>pagecount then 
	startnum=startnum-(endnum-pagecount)
	if startnum<1 then startnum=1
	endnum=pagecount	
end if
dim i
if pageListCount mod 2=0 then endnum=endnum-1
for i=startNum to endnum
	'response.write i
	if cint(i)=cint(page) then
		hcontent = hcontent&"<span style=""color:#ff0000"">"&i&"</span>&nbsp;"
	else
		hcontent = hcontent&"[<a href=""" & temp & "page="&i&""">"&i&"</a>]&nbsp;"
	end if		
Next
if page<pagecount-pageListCount then hcontent=hcontent&"..[<a href="""&temp&"page="&pagecount&""">"&pagecount&"</a>]&nbsp;"
if page<pagecount then hcontent=hcontent&"[<a href="""&temp&"page="&page+1&""">下一页</a>]&nbsp;"
hcontent = hcontent&"&nbsp;<input type=""text"" value="""&page&""" name=""ipage"" id=""ipage"" size=""2"" style=""width:18px"" /><a href=""javascript:;"" onclick=""return goto()"">跳转</a>"&vbcrlf
hcontent=hcontent&"<script language=""javascript"">" & vbcrlf 
hcontent=hcontent&"function goto()" & vbcrlf 
hcontent=hcontent&"{" & vbcrlf 
hcontent=hcontent&"var url="""&temp&""";" & vbcrlf 
hcontent=hcontent&"url=url+""page=""+document.getElementById(""ipage"").value;" & vbcrlf 
hcontent=hcontent&"location.href=url;" & vbcrlf 
hcontent=hcontent&"return false;" & vbcrlf 
hcontent=hcontent&"}" & vbcrlf 
hcontent=hcontent&"</script>" & vbcrlf 
hcontent = hcontent&"</div>"
showpage = hcontent
hcontent=""
End function
%>
