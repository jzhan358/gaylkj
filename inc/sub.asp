<%
'snum条数 swhere附件条件 ssort排序 slist是否分页 slanguage中英文
function getnews(snum,swhere,slist,slanguage)
dim trs,i
dim tt,turl,tturl,tdate,tcontent
dim tpage
if slist=0 then
sql="select top "&snum&" news_id,news_title,news_date,news_count from news where news_id>0"&swhere&" order by news_id desc "
else
	tpage=request.QueryString("page")
	if chkrequest(tpage) then tpage=cint(tpage) else tpage=0
	if tpage>1 then
		sql="SELECT TOP "&snum&" news_id,news_title,news_date,news_count from news where (news_id <(SELECT MIN(news_id) FROM (SELECT TOP "&((tpage-1)*snum)&" news_id FROM news where news_id>0"&sWhere&" order by news_id desc) AS tblTMP)) and news_id>0"&sWhere&" order by news_id desc"
	else
	sql="select top "&snum&" news_id,news_title,news_date,news_count from news where news_id>0"&swhere&" order by news_id desc"
	end if
end if
'response.Write(sql)
'response.End()
set trs=server.CreateObject("adodb.recordset")
trs.open sql,conn,1,1
if not (trs.eof and trs.bof) then
	i=0
	do while not trs.eof and i<snum
		tid=trs("news_id")
		tt=trs("news_title")
		tdate=trs("news_date")
		'if datediff("d",tdate,now())<30 then hot="<img src="""&citymodepath&"/images/hot.gif"" border=""0"" />" else hot=""
		if i<4 then hot="<img src="""&citymodepath&"/images/hot.gif"" border=""0"" />" else hot=""
		if slanguage=0 then			
		turl="/news_view.asp?id="&tid 			
		tlen=12
		else 
		turl="/enews_view.asp?id="&tid
		tlen=30
		end if
		if cityviewmode=1 then turl="/news/"&tid&".html"
		if slist=1 then showdate="<span>[ "&tdate&" ]</span>" else showdate=""
		tcontent=tcontent&"<li>·<a href="""&turl&""" title="""&tt&""">"&left(tt,tlen)&"</a>"&showdate&"</li>" & vbcrlf 		
		trs.movenext
		i=i+1
	loop
else
	tcontent="No Content"
end if
trs.close
set trs=nothing
getnews=tcontent
tcontent=""
end function

'snum条数 swhere附件条件 ssort排序 slist 0首页推荐 1首页最新 2分页
function getProducts(snum,swhere,slist,slanguage)
dim trs,i
dim tt,turl,tpic,tdate,tcontent
dim tpage
if slist=0 then
sql="select top "&snum&" p_id,p_name,p_pic,p_name_e from p_info order by p_order desc"
else
	tpage=request.QueryString("page")
	if chkrequest(tpage) then tpage=cint(tpage) else tpage=0
	if tpage>1 then
		sql="SELECT TOP "&snum&" p_id,p_name,p_pic,p_name_e from p_info where (p_id <(SELECT MIN(p_id) FROM (SELECT TOP "&((tpage-1)*snum)&" p_id FROM p_info where p_id>0"&sWhere&" order by p_id desc) AS tblTMP)) and p_id>0"&sWhere&" order by p_id desc"
	else
	sql="select top "&snum&" p_id,p_name,p_pic,p_name_e from p_info where p_id>0"&swhere&" order by p_id desc"
	end if
end if
set trs=server.CreateObject("adodb.recordset")
trs.open sql,conn,1,1
if not (trs.eof and trs.bof) then
	i=0
	if slist=0 then tcontent="<li>" else tcontent=""
	do while not trs.eof and i<snum
		tid=trs("p_id")
		if slanguage=0 then	
			tt=trs("p_name")
			if cityviewmode=0 then 
				turl="/product_view.asp?id="&tid
			else
				turl="/products/"&tid&".html"
			end if
		else 
			tt=trs("p_name_e")
			if cityviewmode=0 then
				turl="/eproduct_view.asp?id="&tid
			else
				turl="/eproducts/"&tid&".html"
			end if
		end if
		tpic=trs("p_pic")
		if left(tpic,1)<>"/" then tpic="/"&tpic
		if slist=0 then
			if i mod 5=0 and i<>0 then tcontent=tcontent&"</li>"&vbcrlf&"<li>"
			tcontent=tcontent&"<a href="""&turl&""" title="""&tt&""" target=""_blank""><img src="""&tpic&""" alt="""&tt&""" title="""&tt&""" border=""0""><br />"&left(tt,10)&"</a>" & vbcrlf 
		
			
		else
			tcontent=tcontent&"<li><a href="""&turl&""" title="""&tt&""" target=""_blank""><img src="""&tpic&""" alt="""&tt&""" title="""&tt&""" border=""0""></a>" & vbcrlf 
			tcontent=tcontent&"<a href="""&turl&""" title="""&tt&""" target=""_blank"">"&left(tt,10)&"</a></li>" & vbcrlf 		
		end if
		trs.movenext
		i=i+1		
	loop
	if slist=0 then tcontent=tcontent&"</li>"&vbcrlf
else
	tcontent="No Content"
end if
trs.close
set trs=nothing
getProducts=tcontent
tcontent=""
end function


function getClass(snum,slanguage)
dim trs,i
dim tt,turl,tturl,tcontent
dim tmparr
sql="select top "&snum&" p_id,p_type,p_type_e from p_class order by p_id desc "
set trs=server.CreateObject("adodb.recordset")
trs.open sql,conn,1,1
if not (trs.eof and trs.bof) then
	i=0
	tcontent=""
	do while not trs.eof and i<snum
		tid=trs("p_id")
		if slanguage=0 then 
			tt=trs("p_type")
			turl="/products.asp?id="&tid 
		else 
			tt=trs("p_type_e")
			turl="/eproducts.asp?id="&tid
		end if
		sql="select count(p_id) from p_info where p_type_id="&trs("p_id")
		tmparr=getmdbvalue(sql)
		if tmparr(0)=1 then ttnum=tmparr(1) else ttnum=0
		erase tmparr
		tcontent=tcontent&"<li>·<a href="""&turl&""" title="""&tt&""">"&left(tt,14)&"</a>&nbsp;(<font color=""red"">"&ttnum&"</font>)</li>" & vbcrlf 	
		trs.movenext
		i=i+1
	loop
else
	tcontent="No Content"
end if
trs.close
set trs=nothing
getClass=tcontent
tcontent=""
end function


function getAbout(sid)
dim trs
sql="select top 1 inc_text from inc_info where inc_class_id="&sid
set trs=server.CreateObject("adodb.recordset")
trs.open sql,conn,1,1
if not (trs.bof and trs.eof) then
getAbout=replace_t(trs("inc_text"))
else
getAbout="No Content"
end if
trs.close
set trs=nothing
end function

function getHomeAbout(sid,slen)
dim trs
sql="select top 1 inc_text from inc_info where inc_class_id="&sid
set trs=server.CreateObject("adodb.recordset")
trs.open sql,conn,1,1
if not (trs.bof and trs.eof) then
getHomeAbout=leftt(trs("inc_text"),slen)
else
getHomeAbout="No Content"
end if
trs.close
set trs=nothing
end function
%>
