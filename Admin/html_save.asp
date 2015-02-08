<!--#include file="top.asp"-->
<%
Response.Buffer=TRUE
server.ScriptTimeout=99999
if cityviewmode=0 then
	response.Write("您的网站是动态浏览模式，不需要静态生成!")
	response.End()
end if
call createSeoPage()
call savetohtml(getIndexpage(),"/index.html")
call savetohtml(getDefaultpage(),"/default.html")
call savetohtml(getContactpage(),"/contact.html")
call savetohtml(geteContactpage(),"/econtact.html")
call savetohtml(getegbookpage(),"/egbook.html")
call savetohtml(getgbookpage(),"/gbook.html")
call savetohtml(geteintropage(),"/eintro.html")
call savetohtml(getintropage(),"/intro.html")

sql="select p_id from p_info"
tarr=getmdbvaluelist(sql)
if clng(tarr(0,0))>0 then
	for i=0 to ubound(tarr,2)
		errarr=savetohtml(getproduct_viewpage(tarr(0,i)),"/products/"&tarr(0,i)&".html")
		errarr1=savetohtml(geteproduct_viewpage(tarr(0,i)),"/eproducts/"&tarr(0,i)&".html")
	next
end if
erase tarr

sql="select news_id,news_language from news"
tarr=getmdbvaluelist(sql)
if clng(tarr(0,0))>0 then
	for i=0 to ubound(tarr,2)
		l=tarr(1,i)
		if l="1" then	
			call savetohtml(getenews_viewpage(tarr(0,i)),"/news/"&tarr(0,i)&".html")
		else
			call savetohtml(getnews_viewpage(tarr(0,i)),"/news/"&tarr(0,i)&".html")
		end if
	next
end if
erase tarr

call closeconn()
response.Write("恭喜，所有静态已经生成完成")
%>
