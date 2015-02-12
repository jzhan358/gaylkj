<%
function getIndexpage()
dim thead,tbody
dim k,d,t
k=cityname&"首页,广州游戏机供应,广州麻将机供应,广州益智类游戏机供应,广州游艺机供应,广州水果机供应,广州苹果机供应,广州玛丽机供应,广州连线机供应,广州轮盘供应,广州777机供应,广州37机供应,广州弹珠机供应,广州百家乐供应,广州模拟机供应,广州亲子游戏机供应,广州娃娃机供应,广州摇摆机供应,广州框体机供应,广州3D动物供应,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版"
d=cityname&"是一家从事多年游戏机生产、销售、电脑游戏机配件，集科、工、贸于一体的专业公司，并与台湾资深电玩工程师，合作开发多款新产品。"
t=cityname&"首页  广州游戏机批发零售 广州麻将机批发零售 广州益智类游戏机批发零售 广州游艺机批发零售 游戏机厂家直销 专业诚信的供应各类游戏机批发零售 手机：18924115908"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/index.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$产品分类$",getClass(50,0))
tbody=replace(tbody,"$公司动态$",getnews(6," and news_class_id=1",0,0))
tbody=replace(tbody,"$技术资料$",getnews(8," and news_class_id=11",0,0))
tbody=replace(tbody,"$公司介绍$",leftt(getAbout(14),380))
tbody=replace(tbody,"$推荐产品$",getProducts(15,"",0,0))
tbody=replace(tbody,"$最新产品$",getProducts(10,"",1,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getIndexpage=tbody:tbody="":thead=""
end function

function getDefaultpage()
dim thead,tbody
dim k,d,t
k=ecityname&"Home,Guangzhou game supply, Guangzhou mahjong machine supply, the supply of Guangzhou, Puzzle games, Guangzhou Amusement supply, Canton State fruit machine supply"
d=ecityname&"Is a game for many years engaged in the production, marketing, computer game accessories, set Branch, industry and trade in one of the professional company and Taiwan experienced video engineers, to develop a variety of new products"
t=ecityname&"Home_Wholesale and retail console game in Guangzhou, Guangzhou, Guangzhou wholesale and retail purchase online console game professional integrity of the supply of factory direct wholesale and retail all kinds of game consoles.Mobile:18924115908"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/default.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$产品分类$",getClass(50,1))
tbody=replace(tbody,"$公司动态$",getnews(6," and news_class_id=7",0,1))
tbody=replace(tbody,"$技术资料$",getnews(8," and news_class_id=12",0,1))
tbody=replace(tbody,"$公司介绍$",left(getAbout(34),650))
tbody=replace(tbody,"$推荐产品$",getProducts(15,"",0,1))
tbody=replace(tbody,"$最新产品$",getProducts(10,"",1,1))
tbody=replace(tbody,"$foot$",getfoot(1))
getDefaultpage=tbody:tbody="":thead=""
end function

function getContactpage()
dim thead,tbody
dim k,d,t
k="联系我们,在线联系广州玛丽机供应商,在线联系广州连线机供应厂家,在线联系广州轮盘公司,在线联系广州麻将机供应厂家,广州益智类游戏机供应批发零售联系方式,广州游艺机供应批发零售联系方式,广州水果机供应批发零售联系方式,广州苹果机供应批发零售联系方式"
d="联系我们，省份/城市:广东省广州市 详细地址:中国广州市番禺迎宾路 邮政编码:511400 电话号码：+86-20-39977130 传真号码：+86-20-39977130 联系手机：+86-13660170912 ;+86-13719381591 联 系 人：周先生 总经理"
t="联系我们 广州玛丽机供应联系方式 广州连线机供应联系方式 广州轮盘供应联系方式 广州麻将机供应联系方式 广州益智类游戏机供应在线联系方式 广州游艺机在线联系方式 广州水果机在线联系方式 广州苹果机在线联系方式"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/contact.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$getAbout$",getAbout(25))
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getContactpage=tbody:tbody="":thead=""
end function

function geteContactpage()
dim thead,tbody
dim k,d,t
k="Contact us"
d="Contact us"
t="Contact us_How to Contact"&ecityname&" "&ecityname&" Is a professional gaming console, brokers ltd"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/econtact.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$getAbout$",getAbout(35))
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
geteContactpage=tbody:tbody="":thead=""
end function

function getegbookpage()
dim thead,tbody
dim k,d,t
k="GuestBook"
d="On-line message to "&ecityname
t="GuestBook_How to prosper the company a message_"&ecityname&" Is a professional gaming console, brokers ltd"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/egbook.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
getegbookpage=tbody:tbody="":thead=""
end function

function getgbookpage()
dim thead,tbody
dim k,d,t
k="网站留言,游戏机咨询,给"&cityname&"留言,广州游戏机在线咨询,广州麻将机在线咨询,广州益智类游戏机在线咨询,广州游艺机在线咨询,广州水果机在线咨询,广州水果机在线咨询,广州苹果机在线咨询,广州玛丽机在线咨询,广州连线机在线咨询,广州轮盘在线咨询,广州777机在线咨询,广州37机在线咨询"
d="在线给"&cityname&"发送留言信息，在线咨询"&cityname&"供应的游戏相关信息"
t="网站留言咨询_广州弹珠机咨询,广州百家乐咨询,广州模拟机咨询,广州亲子游戏机咨询,广州娃娃机咨询,广州摇摆机咨询,广州框体机咨询,广州3D动物咨询,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版在线咨询"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/gbook.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getgbookpage=tbody:tbody="":thead=""
end function

function geteintropage()
dim thead,tbody
dim k,d,t
k="About Us,about "&ecityname
d="about "&ecityname&" Details"
t="About Us_to Learn "&ecityname&"_"&ecityname&" Is a professional gaming console"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/eintro.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$getAbout$",getAbout(34))
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
geteintropage=tbody:tbody="":thead=""
end function

function getintropage()
dim thead,tbody
dim k,d,t
k="关于我们,关于"&cityname&",关于广州游戏机资料,关于广州麻将机供应,关于广州益智类游戏机供应,关于广州游艺机供应,关于广州水果机供应,关于广州苹果机供应,关于广州玛丽机供应,关于广州连线机供应,关于广州轮盘供应,关于广州777机供应,关于广州37机供应"
d=cityname&"是一家专业游戏机及游戏机配件生产和销售公司，我们有专业的销售团队，过硬的技术水平，完善的售后服务"
t="关于我们_"&cityname&"介绍 关于广州游戏机介绍资料,广州麻将机介绍资料,广州益智类游戏机介绍资料,广州游艺机介绍资料,广州水果机介绍资料,广州苹果机介绍资料,广州玛丽机介绍资料,广州连线机介绍资料"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/intro.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$getAbout$",getAbout(14))
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getintropage=tbody:tbody="":thead=""
end function

function getejishupage()
dim thead,tbody
dim k,d,t
dim pre,pnext
k="TECHNICAL,"&ecityname
d="about "&ecityname&" TECHNICAL Details"
t="TECHNICAL_view "&ecityname&" TECHNICAL Details"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/ejishu.html")
tbody=replace(tbody,"$head$",thead)
dim pagesize,pagecount,recordcount,page
PageSize=18  
'获得总数  
sql="select count(news_id) from news where news_class_id=12"
recordcount=conn.execute(sql)(0)
'总页数  
if cint(recordcount) = 0 then 
	pagecount=1
else
	pagecount=Abs(Int(recordcount/PageSize*(-1)))   
end if
'获得当前页码   
page=request.QueryString("page")
if not chkrequest(page) then page=1 else page=cint(page)
if page>pagecount then page=pagecount
tbody=replace(tbody,"$getnews$",getnews(pagesize," and news_class_id=12",1,1))
tbody=replace(tbody,"$pages$",showpage(pagecount,pagesize,page,recordcount,10))
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
getejishupage=tbody:tbody="":thead=""
end function


function getjishupage()
dim thead,tbody
dim k,d,t
dim pre,pnext
k="技术资料,游戏机使用说明,"&cityname&"游戏机使用说明,游戏机技术资料,麻将机技术资料,益智类游戏机技术资料,游艺机技术资料,水果机技术资料,苹果机技术资料,玛丽机技术资料,连线机技术资料,轮盘技术资料,777机技术资料,37机技术资料,弹珠机技术资料,百家乐技术资料,模拟机技术资料,亲子游戏机技术资料,娃娃机技术资料,摇摆机技术资料,框体机技术资料,3D动物技术资料,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版技术资料"
d=cityname&"游戏机相关的技术资料列表"
t="技术资料 游戏机使用说明,游戏机技术资料,麻将机使用说明,益智类游戏机使用说明,游艺机使用说明,水果机使用说明,苹果机使用说明,玛丽机使用说明,连线机使用说明,轮盘使用说明,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版使用说明"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/jishu.html")
tbody=replace(tbody,"$head$",thead)
dim pagesize,pagecount,recordcount,page
PageSize=18  
'获得总数  
sql="select count(news_id) from news where news_class_id=11"
recordcount=conn.execute(sql)(0)
'总页数  
if cint(recordcount) = 0 then 
	pagecount=1
else
	pagecount=Abs(Int(recordcount/PageSize*(-1)))   
end if
'获得当前页码   
page=request.QueryString("page")
if not chkrequest(page) then page=1 else page=cint(page)
if page>pagecount then page=pagecount
tbody=replace(tbody,"$getnews$",getnews(pagesize," and news_class_id=11",1,0))
tbody=replace(tbody,"$pages$",showpage(pagecount,pagesize,page,recordcount,10))
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getjishupage=tbody:tbody="":thead=""
end function

function getenewspage()
dim thead,tbody
dim k,d,t
dim pre,pnext
k="News,"&ecityname
d="view "&ecityname&" News Details"
t="News_about "&ecityname&" News Details"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/enews.html")
tbody=replace(tbody,"$head$",thead)
dim pagesize,pagecount,recordcount,page
PageSize=18  
'获得总数  
sql="select count(news_id) from news where news_class_id=7"
recordcount=conn.execute(sql)(0)
'总页数  
if cint(recordcount) = 0 then 
	pagecount=1
else
	pagecount=Abs(Int(recordcount/PageSize*(-1)))   
end if
'获得当前页码   
page=request.QueryString("page")
if not chkrequest(page) then page=1 else page=cint(page)
if page>pagecount then page=pagecount
tbody=replace(tbody,"$getnews$",getnews(pagesize," and news_class_id=7",1,1))
tbody=replace(tbody,"$pages$",showpage(pagecount,pagesize,page,recordcount,10))
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
getenewspage=tbody:tbody="":thead=""
end function

function getnewspage()
dim thead,tbody
dim k,d,t
dim pre,pnext
k="公司动态,游戏机资讯,"&cityname&"游戏机资讯,麻将机资讯,益智类游戏机最新资讯,游艺机最新资讯,水果机最新资讯,苹果机最新资讯,玛丽机最新资讯,连线机最新资讯,轮盘最新资讯,777机最新资讯,37机最新资讯,弹珠机最新资讯,百家乐最新资讯,模拟机最新资讯,亲子游戏机最新资讯,娃娃机最新资讯,摇摆机最新资讯,框体机最新资讯,3D动物最新资讯,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版最新资讯"
d=cityname&"公司动态，游戏机最新行业动态"
t="公司动态,游戏机资讯 游戏机行业资讯,游戏机行业资讯,麻将机行业资讯,益智类游戏机行业资讯,游艺机最新动态,水果机最新资讯动态,苹果机最新资讯动态,玛丽机最新资讯动态,连线机最新资讯动态,轮盘最新资讯动态,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版最新资讯动态"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/news.html")
tbody=replace(tbody,"$head$",thead)
dim pagesize,pagecount,recordcount,page
PageSize=18  
'获得总数  
sql="select count(news_id) from news where news_class_id=1"
recordcount=conn.execute(sql)(0)
'总页数  
if cint(recordcount) = 0 then 
	pagecount=1
else
	pagecount=Abs(Int(recordcount/PageSize*(-1)))   
end if
'获得当前页码   
page=request.QueryString("page")
if not chkrequest(page) then page=1 else page=cint(page)
if page>pagecount then page=pagecount
tbody=replace(tbody,"$getnews$",getnews(pagesize," and news_class_id=1",1,0))
tbody=replace(tbody,"$pages$",showpage(pagecount,pagesize,page,recordcount,10))
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getnewspage=tbody:tbody="":thead=""
end function

function getenews_viewpage(sid)
dim thead,tbody
dim k,d,t
dim tmparr,newsclass,rsNews
strsql="select * from news where news_id="&sid
set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.Open strsql, conn, 1, 1

if rsNews.bof and rsNews.eof then
	call closers(rsNews)
	alert "error","",1
end if
sql="select news_type from news_class where news_id="&rsNews("news_class_id")
tmparr=getmdbvalue(sql)
if tmparr(0)=1 then newsclass=tmparr(1)
k=rsNews("news_keyword")&",new details,"&rsNews("news_title")&","&newsclass&","&ecityname
d=leftt(rsNews("news_content"),100)
t=rsNews("news_title")&"_"&newsclass&"_"&ecityname&" Is a professional gaming console"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/enews_view.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$id$",sid)
tbody=replace(tbody,"$news_title$",rsNews("news_title"))
tbody=replace(tbody,"$news_date$",rsNews("news_date"))
tbody=replace(tbody,"$news_ahome$",rsNews("news_ahome"))
tbody=replace(tbody,"$news_author$",rsNews("news_author"))
tbody=replace(tbody,"$news_content$",replace_t(rsNews("news_content")))
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
getenews_viewpage=tbody:tbody="":thead=""
end function

function getnews_viewpage(sid)
dim thead,tbody
dim k,d,t
dim tmparr,newsclass,rsNews
strsql="select * from news where news_id="&sid
set rsNews = Server.CreateObject("ADODB.Recordset")
rsNews.Open strsql, conn, 1, 1

if rsNews.bof and rsNews.eof then
	call closers(rsNews)
	alert "error","",1
end if
sql="select news_type from news_class where news_id="&rsNews("news_class_id")
tmparr=getmdbvalue(sql)
if tmparr(0)=1 then newsclass=tmparr(1)
k=rsNews("news_keyword")&",信息详细,"&rsNews("news_title")&","&newsclass&","&cityname
d=leftt(rsNews("news_content"),100)
t=rsNews("news_title")&"_"&newsclass&"_"&cityname&" 是专业的游戏机生产厂家，销售公司, 快速联系手机:18924115908"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/news_view.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$id$",sid)
tbody=replace(tbody,"$news_title$",rsNews("news_title"))
tbody=replace(tbody,"$news_date$",rsNews("news_date"))
tbody=replace(tbody,"$news_ahome$",rsNews("news_ahome"))
tbody=replace(tbody,"$news_author$",rsNews("news_author"))
tbody=replace(tbody,"$news_content$",replace_t(rsNews("news_content")))
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getnews_viewpage=tbody:tbody="":thead=""
end function

function geteorderpage()
dim thead,tbody
dim k,d,t
dim id,title,hh
id=request.QueryString("id")
title=""
if chkrequest(id) then	
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select p_name_e,p_spec_e from p_info where p_id="&id,conn,1,1	
	if not (rs.eof and rs.bof) then
		title=rs("p_name_e")
		hh=rs("p_spec_e")	
	else 
		id=""
	end if
	rs.close
	set rs=nothing
end if
k=title&","&hh&",Order"
d=title&","&hh&",Order"
t="Order_online buy "&ecityname&" products"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/eorder.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$title$",title)
tbody=replace(tbody,"$hh$",hh)
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
geteorderpage=tbody:tbody="":thead=""
end function

function getorderpage()
dim thead,tbody
dim k,d,t
dim id,title,hh
id=request.QueryString("id")
title=""
if chkrequest(id) then	
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select p_name,p_spec from p_info where p_id="&id,conn,1,1	
	if not (rs.eof and rs.bof) then
		title=rs("p_name")
		hh=rs("p_spec")	
	else 
		id=""
	end if
	rs.close
	set rs=nothing
end if
k=title&","&hh&",在线订单,在线订购"&title&"产品"
d=cityname&"在线订购"&title&"产品"
t="在线订购 在线购买"&title&"游戏机产品 购买"&cityname&"游戏机产品，质量保证，信誉保证，在线快速下单 线下联系手机:18924115908"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/order.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$title$",title)
tbody=replace(tbody,"$hh$",hh)
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getorderpage=tbody:tbody="":thead=""
end function

function geteproduct_viewpage(sid)
dim thead,tbody
dim k,d,t
dim rsPro,pic
dim pre,nex,pret,nextt
strsql="select * from p_info where p_id="&sid
set rsPro = Server.CreateObject("ADODB.Recordset")
rsPro.Open strsql, conn, 1, 1
if rsPro.bof and rsPro.eof then
	call closers(rsPro)
	geteproduct_view="false"
	exit function
end if
k=rsPro("p_name_e")&","&rsPro("p_type_e")&",Products"
d=leftt(rsPro("p_jianjie_e"),100)
t=rsPro("p_name_e")&"_"&rsPro("p_type_e")&" products_"&ecityname&" products"
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/eproduct_view.html")
pic=rsPro("p_pic2")
if chknull(pic,5) then
	pic=rsPro("p_pic")
end if

if left(pic,1)<>"/" then pic="/"&pic
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$id$",sid)
tbody=replace(tbody,"$pic$",pic)
tbody=replace(tbody,"$p_name_e$",rsPro("p_name_e"))
tbody=replace(tbody,"$p_spec_e$",rsPro("p_spec_e"))
tbody=replace(tbody,"$p_jianjie_e$",replace_t(rsPro("P_jianjie_e")))
rsPro.close
rsPro.open "select top 1 p_id from p_info where p_id<"&sid&" order by p_id desc",conn,1,1
if not(rsPro.bof and rsPro.eof) then
pre=clng(rsPro("p_id"))
else
pre=0
end if
rsPro.close
rsPro.open "select top 1 p_id from p_info where p_id>"&sid&" order by p_id asc",conn,1,1
if not(rsPro.bof and rsPro.eof) then
nex=clng(rsPro("p_id"))
else
nex=0
end if
rsPro.close
set rsPro=nothing
if pre<>0 then 
	if cityviewmode=0 then
	pret="<a href=""/eproduct_view.asp?id="&pre&""">Previous Product</a>"
	else
	pret="<a href=""/eproducts/"&pre&".html"">Previous Product</a>"
	end if
end if
if nex<>0 then
	if cityviewmode=0 then
	nextt="<a href=""/eproduct_view.asp?id="&nex&""">Next Product</a>"
	else
	nextt="<a href=""/eproducts/"&nex&".html"">Next Product</a>"
	end if
end if
tbody=replace(tbody,"$pre$",pret)
tbody=replace(tbody,"$next$",nextt)
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
geteproduct_viewpage=tbody:tbody="":thead=""
end function

function getproduct_viewpage(sid)
dim thead,tbody
dim k,d,t
dim rsPro,pic
dim pre,nex,pret,nextt
strsql="select * from p_info where p_id="&sid
set rsPro = Server.CreateObject("ADODB.Recordset")
rsPro.Open strsql, conn, 1, 1
if rsPro.bof and rsPro.eof then
	call closers(rsPro)
	geteproduct_view="false"
	exit function
end if
k=rsPro("p_name")&","&rsPro("p_type")&","&cityname&"游戏机产品详细"
d=leftt(rsPro("p_jianjie"),100)
t=rsPro("p_name")&" "&rsPro("p_type")&"相关游戏机及配件 "&cityname&" 销售热线:18924115908"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/product_view.html")
pic=rsPro("p_pic2")
if chknull(pic,5) then
	pic=rsPro("p_pic")
end if

if left(pic,1)<>"/" then pic="/"&pic
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$id$",sid)
tbody=replace(tbody,"$pic$",pic)
tbody=replace(tbody,"$p_name$",rsPro("p_name"))
tbody=replace(tbody,"$p_spec$",rsPro("p_spec"))
tbody=replace(tbody,"$p_jianjie$",replace_t(rsPro("P_jianjie")))
rsPro.close
rsPro.open "select top 1 p_id from p_info where p_id<"&sid&" order by p_id desc",conn,1,1
if not(rsPro.bof and rsPro.eof) then
pre=clng(rsPro("p_id"))
else
pre=0
end if
rsPro.close
rsPro.open "select top 1 p_id from p_info where p_id>"&sid&" order by p_id asc",conn,1,1
if not(rsPro.bof and rsPro.eof) then
nex=clng(rsPro("p_id"))
else
nex=0
end if
rsPro.close
set rsPro=nothing
if pre<>0 then 
	if cityviewmode=0 then
	pret="<a href=""/product_view.asp?id="&pre&""">上一个产品</a>"
	else
	pret="<a href=""/products/"&pre&".html"">上一个产品</a>"
	end if
end if
if nex<>0 then
	if cityviewmode=0 then
	nextt="<a href=""/product_view.asp?id="&nex&""">下一个产品</a>"
	else
	nextt="<a href=""/products/"&nex&".html"">下一个产品</a>"
	end if
end if
tbody=replace(tbody,"$pre$",pret)
tbody=replace(tbody,"$next$",nextt)
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getproduct_viewpage=tbody:tbody="":thead=""
end function


function geteproductspage()
dim thead,tbody
dim k,d,t
dim id,st
dim tmparr,classname
id=request.QueryString("id")
st="":classname=""
if chkrequest(id) then 
	sql="select p_type_e from p_class where p_id="&id
	tmparr=getmdbvalue(sql)
	if tmparr(0)=1 then
	st=st&" and  p_type_id="&id
	classname=tmparr(1)
	end if
	erase tmparr
end if
k=checkStr(request.QueryString("k"))
if len(k)>1 and len(k)<20 then
	st=" and p_name like '%"&k&"%'"
end if
dim pagesize,pagecount,recordcount,page
PageSize=12   
'获得总数  
sql="select count(p_id) from p_info where p_id>0"&st
recordcount=conn.execute(sql)(0)
'总页数  
if cint(recordcount) = 0 then 
	pagecount=1
else
	pagecount=Abs(Int(recordcount/PageSize*(-1)))   
end if
'获得当前页码   
page=request.QueryString("page")
if not chkrequest(page) then page=1 else page=cint(page)
if page>pagecount then page=pagecount
k="products"
d="All products"
t="products_view all products of "&ecityname
thead=getcache("/ehead.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/eproducts.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$id$",id)
tbody=replace(tbody,"$getproducts$",getproducts(pagesize,st,1,1))
tbody=replace(tbody,"$pages$",showpage(pagecount,pagesize,page,recordcount,10))
tbody=replace(tbody,"$left$",getleft(1,0))
tbody=replace(tbody,"$foot$",getfoot(1))
geteproductspage=tbody:tbody="":thead=""
end function

function getproductspage()
dim thead,tbody
dim k,d,t
dim id,st
dim tmparr,classname
id=request.QueryString("id")
st="":classname=""
if chkrequest(id) then 
	sql="select p_type from p_class where p_id="&id
	tmparr=getmdbvalue(sql)
	if tmparr(0)=1 then
	st=st&" and  p_type_id="&id
	classname=tmparr(1)
	end if
	erase tmparr
end if
k=checkStr(request.QueryString("k"))
if len(k)>1 and len(k)<20 then
	st=" and p_name like '%"&k&"%'"
end if
dim pagesize,pagecount,recordcount,page
PageSize=12   
'获得总数  
sql="select count(p_id) from p_info where p_id>0"&st
recordcount=conn.execute(sql)(0)
'总页数  
if cint(recordcount) = 0 then 
	pagecount=1
else
	pagecount=Abs(Int(recordcount/PageSize*(-1)))   
end if
'获得当前页码   
page=request.QueryString("page")
if not chkrequest(page) then page=1 else page=cint(page)
if page>pagecount then page=pagecount
k=cityname&"游戏机产品展示,"&classname&"产品展示,麻将机产品展示,益智类游戏机产品展示,游艺机产品展示,广州水果机产品展示,广州苹果机产品展示,广州玛丽机产品展示,广州连线机产品展示,广州轮盘产品展示,广州777机产品展示,广州37机产品展示,广州弹珠机产品展示,广州百家乐产品展示,广州模拟机产品展示,广州亲子游戏机产品展示,广州娃娃机产品展示,广州摇摆机产品展示,广州框体机产品展示,广州3D动物产品展示,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版产品展示"
d=cityname&"所有游戏机产品展示,广州"&classname&"游戏机产品展示"
t=cityname&"产品展示_广州"&classname&"游戏机产品展示 广州游戏机产品展示,广州麻将机产品展示,广州益智类游戏机产品展示,广州游艺机产品展示,广州水果机产品展示,广州苹果机产品展示,广州玛丽机产品展示,广州连线机产品展示,广州轮盘产品展示,广州777机产品展示,捞金鱼,SNK卡带,SNK卡座,游戏电脑板,框体节目版等游戏机产品展示"
thead=getcache("/head.html")
thead=setkw(thead,k,d,t)
tbody=getcache("/products.html")
tbody=replace(tbody,"$head$",thead)
tbody=replace(tbody,"$id$",id)
tbody=replace(tbody,"$getproducts$",getproducts(pagesize,st,1,0))
tbody=replace(tbody,"$pages$",showpage(pagecount,pagesize,page,recordcount,10))
tbody=replace(tbody,"$left$",getleft(0,0))
tbody=replace(tbody,"$foot$",getfoot(0))
getproductspage=tbody:tbody="":thead=""
end function
%>
