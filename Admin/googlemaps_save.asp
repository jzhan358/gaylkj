<!--#include file="top.asp"-->
<!--#include file="googlemaps_sub.asp"-->
<%
dim str,filenamelastDate

str = "<?xml version='1.0' encoding='UTF-8'?>" & vbcrlf
str = str & "<urlset xmlns='http://www.google.com/schemas/sitemap/0.84'>" & vbcrlf

'首页
str = str &getfilelink(WebUrlAddr&"/",ShowDatecreated("/index."&weburlext))& vbcrlf
str = str &getfilelink(WebUrlAddr&"/default."&weburlext,ShowDatecreated("/default."&weburlext))& vbcrlf

'
str = str &getfilelink(WebUrlAddr&"/contact."&weburlext,ShowDatecreated("/contact."&weburlext))& vbcrlf
str = str &getfilelink(WebUrlAddr&"/econtact."&weburlext,ShowDatecreated("/econtact."&weburlext))& vbcrlf

'
str = str &getfilelink(WebUrlAddr&"/intro."&weburlext,ShowDatecreated("/intro."&weburlext))& vbcrlf
str = str &getfilelink(WebUrlAddr&"/eintro."&weburlext,ShowDatecreated("/eintro."&weburlext))& vbcrlf

'
str = str &getfilelink(WebUrlAddr&"/order.asp",ShowDatecreated("/order.asp"))& vbcrlf
str = str &getfilelink(WebUrlAddr&"/eorder.asp",ShowDatecreated("/eorder.asp"))& vbcrlf

'
str = str &getfilelink(WebUrlAddr&"/gbook."&weburlext,ShowDatecreated("/gbook."&weburlext))& vbcrlf
str = str &getfilelink(WebUrlAddr&"/egbook."&weburlext,ShowDatecreated("/egbook."&weburlext))& vbcrlf

'
sql="select p_id,p_date from p_info"
tarr=getmdbvaluelist(sql)
if clng(tarr(0,0))>0 then
	for i=0 to ubound(tarr,2)
		if cityviewmode=0 then
			str = str &getfilelink(WebUrlAddr&"/product_view.asp?id="&tarr(0,i),formatdatetime(tarr(1,i),2))& vbcrlf
			str = str &getfilelink(WebUrlAddr&"/eproduct_view.asp?id="&tarr(0,i),formatdatetime(tarr(1,i),2))& vbcrlf
		else
			str = str &getfilelink(WebUrlAddr&"/products/"&tarr(0,i)&".html",ShowDatecreated("/products/"&tarr(0,i)&".html"))& vbcrlf
			str = str &getfilelink(WebUrlAddr&"/eproducts/"&tarr(0,i)&".html",ShowDatecreated("/eproducts/"&tarr(0,i)&".html"))& vbcrlf
		end if
	next
end if
erase tarr

'
sql="select new_id,news_date from news"
tarr=getmdbvaluelist(sql)
if clng(tarr(0,0))>0 then
	for i=0 to ubound(tarr,2)
		if cityviewmode=0 then
			str = str &getfilelink(WebUrlAddr&"/news_view.asp?id="&tarr(0,i),formatdatetime(tarr(1,i),2))& vbcrlf
			str = str &getfilelink(WebUrlAddr&"/enews_view.asp?id="&tarr(0,i),formatdatetime(tarr(1,i),2))& vbcrlf
		else
			str = str &getfilelink(WebUrlAddr&"/news/"&tarr(0,i)&".html",ShowDatecreated("/news/"&tarr(0,i)&".html"))& vbcrlf			
		end if
	next
end if
erase tarr

'
sql="select p_id from p_class"
tarr=getmdbvaluelist(sql)
if clng(tarr(0,0))>0 then
	for i=0 to ubound(tarr,2)	
		str = str &getfilelink(WebUrlAddr&"/products.asp?id="&tarr(0,i),ShowDatecreated("/products.asp"))& vbcrlf
	next
end if
erase tarr


str = str & "</urlset>" & vbcrlf
errid=savetofile(str,"/sitemap.xml")
call closeconn()
if errid=0 then
response.Write("恭喜，生成完成！")
else
response.Write("生成失败！保存地图文件失败")
end if
Response.End
%>
