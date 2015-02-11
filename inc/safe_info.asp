<%
' by selectersky
' 2009.11.20修改了changeSwebUrl函数 11.25修改了delUrlPara函数
' 2009.11.15加入getmdbvalue getmdbvaluelist setmdbvalue函数 2009.11.18重新书写了delUrlPara函数
' 2009.9.17 修改了IsValidEmail函数 邮箱用户名支持.
' 2009.8.4	加入changeSwebUrl函数 2009.8.13修改chkRequest函数 2009.8.28 增加chkSql子程序
' 2009.6.24 加入Chknumber函数,修改了getDirect()函数 如果cookies为空，刚返回到首页 2009.7.21修改alert函数
' 2009.6.9  修改alert函数，加入flag参数 修改所以函数的参数传值方式 2009.6.10加入chkpara函数 修改alert函数
' 2009.4.27 重新书写delUrlpara函数 2009.5.3加入getTest函数 2009.5.19加入getrnd函数
' 2009.4.3  修改showmsg函数跳转不正确的问题 2009.4.6 修改getlen的错误 2009.4.18 JmailSend有返回错误ID
' 2009.4.2	将makePassword函数更名为getRndNum 增加isExists函数
' 2009.3.30	修改了getDirect()函数 如果cookies为空，刚返回上一页的地址,修改了检查会员是否登录的函数chkLogin(),改为通用

'会员发布的各种信息过滤
Function Replace_Text(byVal fString)
If Not IsNull(fString) and not isempty(fString) and fstring<>"" Then
'fString = trim(fString)
dim regEx
set regEx=new regExp
regEx.ignoreCase=true
regEx.Global=True
'<script
regEx.pattern="<script.+?script>"
fString=regEx.replace(fString,"")
'过滤asp代码
regEx.pattern="<\x25.+?\x25>"
fString=regEx.replace(fString,"")
'过滤php代码
regEx.pattern="<\x3f.+?\x3f>"
fString=regEx.replace(fString,"")

'<ifream></ifream>
regEx.pattern="<ifream.+?ifream>"
fString=regEx.replace(fString,"")
'<frameset></frameset>
regEx.pattern="<frameset.+?frameset>"
fString=regEx.replace(fString,"")
set regEx=nothing
fString = replace(fString, chr(62), "&gt;") '>
fString = replace(fString, "<", "&lt;")
fString = replace(fString, chr(60), "&lt;") '<
fString = replace(fString, """", "&quot;")
fString = Replace(fString, CHR(34), "&quot;") '双引号过滤
fString = Replace(fString, "'", "''")	'单引号过滤
End If
Replace_Text = fString
fString=""
End Function
'转换
Function replace_T(byVal fstring)
If Not IsNull(fString) and fString<>"" and not isempty(Fstring) Then
'fString = trim(fString)
fString = Replace(fString, "&lt;", "<")
fString = Replace(fString, "&gt;", ">")
fString = Replace(fString, "&quot;", """")
fString = Replace(fString, "''", "'")
end if 
replace_T = fString
fString=""
end Function

'完全过滤html
Function getInnerText(byVal strHTML)
strHTML=trim(strHTML)
  if strHTML="" or isnull(strHTML) or isempty(strHtml) then
    getInnerText="":exit function
  end if
    Dim objRegExp
    Set objRegExp = New Regexp    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    '取闭合的<>
    objRegExp.Pattern = "<.+?>"
    '进行替换匹配
	strHTML=objRegExp.replace(strHTML,"")
	objRegExp.Pattern = "&lt;.+?&gt;"
	strHTML=objRegExp.replace(strHTML,"")
   	Set objRegExp = Nothing
	strHTML=replace(strHTML,"&nbsp;","")
	strHTML=replace(strHTML," ","")	
	strHTML = Replace(strHTML, CHR(32), "")		'&nbsp;
	strHTML = Replace(strHTML, CHR(9), "")			'&nbsp;
	strHTML = Replace(strHTML, CHR(10), "")			'&nbsp;
	strHTML = Replace(strHTML, CHR(13), "")	
    getInnerText=strHTML
	strHTML=""
    Set objRegExp = Nothing
End Function

'获取字符串长度 取得字符串长度（汉字为2)
function getLen(str)
			Dim Rep,lens,i
			Set rep=new regexp
			rep.Global=true
			rep.IgnoreCase=true
			rep.Pattern="[\u4E00-\u9FA5\uF900-\uFA2D]"
			For each i in rep.Execute(str)
				lens=lens+1
			Next
			Set Rep=Nothing
			lens=lens + len(str)
			getLen=lens			
End Function

'取完全过滤的html的字符串数
Function LeftT(byVal str,n)
  if str="" or isnull(str) or isempty(str) then
    LeftT="":exit function
  end if
str=getInnerText(str)
If len(str)<=n/2 Then
LeftT=str
Else
Dim TStr
Dim l,t,c
Dim i
l=len(str)
if l > 0 then 
	TStr=""
	t=0
	for i=1 to l
	c=asc(mid(str,i,1))
	If c<0 then c=c+65536
	If c>255 then
	t=t+2
	Else
	t=t+1
	End If
	If t>n Then exit for
	TStr=TStr&(mid(str,i,1))
	next
	LeftT = TStr
	End If
end if
str=""
End Function

'过滤SQL非法字符
Function checkStr(byVal Str)
	Str=trim(Str)	
	if isnull(Str) or Str="" or isempty(str) then
		checkStr =""
		exit Function
	else
		Dim objRegExp							
		Set objRegExp = New Regexp    
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		'取闭合的<>
		objRegExp.Pattern = "<.+?>"
		'进行替换匹配
		Str=objRegExp.replace(Str,"")
		objRegExp.Pattern = "&lt;.+?&gt;"
		Str=objRegExp.replace(Str,"")
		Set objRegExp = Nothing		
		Str = replace(Str,"''","'")
		Str = replace(Str,"'","''")
		checkStr=Str
	end if	
	Str=""
End Function

'检测传递的参数是否为大于0的数字型
Function Chkrequest(byVal Para)
Chkrequest=False
Para=trim(Para)
dim tempNum
if isnull(para) or para="" or not isnumeric(para) or isempty(para) then exit function
tempNum=cdbl(Para)
if tempNum>0 then
 Chkrequest=True
end if
Para="":tempNum=""
End Function

'检测传递的参数是否为大于0的数字型
Function Chknumber(byVal Para)
Chknumber=False
Para=trim(Para)
dim tempNum
If not IsNull(Para) and not Para="" and IsNumeric(Para) and not isempty(para) Then
	Chknumber=true
End If
Para="":tempNum=""
End Function

'检查字符串是满足长度
Function ChkNull(byVal tempPara,strlen)
tempPara=trim(tempPara)
if isnull(tempPara) or tempPara="" or len(tempPara)<strlen then
ChkNull=True
else
ChkNull=False
end if
tempPara=""
end Function

'检查字符串满足长度的范围 如果满足，返回True 否则，返回False
function chkRange(byVal tempPara,smin,sMax)
tempPara=trim(tempPara)
if isnull(tempPara) or tempPara="" or len(tempPara)<smin or len(tempPara)>sMax then
chkRange=False
else
chkRange=True
end if
tempPara=""
end function

'检查是否汉字 true是，false否
function chkChinese(para)
	if chkNull(para,1) then 
		chkChinese=false
		exit function
	end if
	Set RegExpObj=new RegExp   
	RegExpObj.Pattern="^[\u4e00-\u9fa5]+$"   
	chkChinese=RegExpObj.test(para) 
	Set RegExpObj=nothing		
end function

'检查用户名是否合法 字母开头，4-20位的数字，字母，下划线组合
Function isValidName(str)
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "^[A-Za-z][A-Za-z0-9_]{3,19}$"
    isValidName = re.Test(str)
    Set re = Nothing	
End Function 

'返回定向文件
function getDirect()
dim tempUrl
tempUrl=checkstr(request.Cookies("direct"))
if not chkRange(tempUrl,1,100) then
tempUrl="/"
end if
getDirect=tempUrl:tempUrl=""
end function

'********************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
'********************************************
function IsValidEmail(temail)	
	dim regExpobj
	set regExpobj=new regexp
	regExpobj.IgnoreCase=True
	regExpobj.Global=true
	regExpobj.Pattern="^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
	IsValidEmail=regExpobj.test(temail)
	set regExpobj=nothing	
end function

'自定义正则表达式 
function getTest(Pattern,testTxt)
dim expobj
set expobj=new regexp
expobj.IgnoreCase=true
expobj.Global=true
expobj.Pattern=Pattern
getTest=expobj.test(testTxt)
set expobj=nothing
end function

'关闭对象
sub closers(Para)	
	if typename(Para)<>"Nothing" and isobject(Para) then
		if Para.state<>0 then Para.close	
		set Para=nothing
	end if
end sub

'释放资源
sub closeconn()'
	call closers(rs)
	if typename(conn)<>"Nothing" and isobject(conn) then
		if conn.state<>0 then conn.close	
		set conn=nothing	
	end if	
end sub
'****************************************************
'函数名：Showmsg
'作  用：信息提示函数
'参  数：msg            消息提示信息
'       href            跳转页面 为空 返回上一页 其他 指定页面
'       smode        消息类型，成功提醒 1 ；错误提醒 0 
'****************************************************
Sub Showmsg(msg,href,smode)	
'图片名称：smile0.gif smile1.gif returnBack.gif returnHome.gif
dim tcontent
tcontent=tcontent&"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbcrlf 
tcontent=tcontent&"<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbcrlf 
tcontent=tcontent&"<head>" & vbcrlf 
tcontent=tcontent&"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" & vbcrlf
if smode<>0 then smode=1
if smode=0 then
tcontent=tcontent&"<title>出错啦</title>" & vbcrlf 
else
tcontent=tcontent&"<title>操作成功</title>" & vbcrlf 
end if
tcontent=tcontent&"<style>" & vbcrlf 
tcontent=tcontent&"body{margin:0; padding:0; border:0px none; background:#fff; color:#000; font-size:12px; font-family:'宋体',arial,'Lucida Grande','Lucida Sans Unicode','新宋体',verdana,sans-serif; line-height:180%}" & vbcrlf 
tcontent=tcontent&"a {color:#333333; text-decoration:none}" & vbcrlf 
tcontent=tcontent&"a:hover {color:#f50; text-decoration:underline}" & vbcrlf 
tcontent=tcontent&"h2 {font-size:14px}" & vbcrlf 
tcontent=tcontent&"#bg{background:#333; opacity: 0.5;-moz-opacity:0.5;filter:alpha(opacity=20);  width:100%; height:100%;position:absolute; top:0; left:0 }" & vbcrlf 
tcontent=tcontent&"#info{height:0px; width:0px;top:50%; left:50%;position:absolute; line-height:1.7 }" & vbcrlf 
tcontent=tcontent&"#center{background:#fff; border:1px solid #357bb9; width:600px; height:300px; position:absolute; margin:-150px -300px }" & vbcrlf 
tcontent=tcontent&"#center strong{display:block; padding:2px 5px; background:#eef8f9; color:#4d4d4d; font-weight:100; width:590px; height:22px }" & vbcrlf 
tcontent=tcontent&"#center p{padding:10px }" & vbcrlf 
tcontent=tcontent&"</style>" & vbcrlf 
tcontent=tcontent&"</head>" & vbcrlf 
tcontent=tcontent&"" & vbcrlf 
tcontent=tcontent&"<body>" & vbcrlf 
tcontent=tcontent&"" & vbcrlf 
tcontent=tcontent&"<span id=""errboxs"">" & vbcrlf 
tcontent=tcontent&"<div id=""bg""></div>" & vbcrlf 
tcontent=tcontent&"<div id=""info"">" & vbcrlf 
tcontent=tcontent&"<div id=""center"">" & vbcrlf 
tcontent=tcontent&"<strong><span style=""float:left"">温馨提示</span><span style=""float:right""><a href=""javascript:window.close(this)"">关闭窗口</a></span></strong>" & vbcrlf 
tcontent=tcontent&"<p>" & vbcrlf 
tcontent=tcontent&"<div style=""overflow:hidden; margin:0 auto; width:400px"">" & vbcrlf 
tcontent=tcontent&"	<span style=""float:left; width:150px; height:100px; background:url(/images/smile"&smode&".gif) no-repeat""></span>" & vbcrlf 
if smode=0 then
tcontent=tcontent&"	<span style=""float:left;width:250px""><h2 style=""color:#f50"">对不起，您的操作有误！可能原因：</h2>" & vbcrlf
else
tcontent=tcontent&"	<span style=""float:left;padding-top:10px;width:250px""><h2 style=""color:#f50"">恭喜你，操作成功！</h2>" & vbcrlf 
end if
tcontent=tcontent&"	<ul>" & vbcrlf 
if instr(msg,"li")=0 then msg="<li>"&msg&"</li>"&vbcrlf
tcontent=tcontent& msg & vbcrlf 
'msg的格式：<li>您尚未登录，请先进行<a href=""/reg/login.html"" class=""yellow"">登录</a>！</li>
tcontent=tcontent&"</ul>" & vbcrlf 
tcontent=tcontent&"</span></div>" & vbcrlf 
tcontent=tcontent&"<div style=""clear:both;margin:30px auto; width:200px"">" & vbcrlf 
tcontent=tcontent&"<div style=""float:left;margin-right:10px;width:85px;height:21px;background:url(/images/returnBack.gif);text-align:center"">" & vbcrlf 
if href="" then href="history.back();" else href="location.href='"&href&"';"
tcontent=tcontent&"	<a href=""javascript:"&href&""" class=""menuwhite"">我要返回</a>" & vbcrlf 
tcontent=tcontent&"</div>" & vbcrlf 
tcontent=tcontent&"<div style=""float:left;width:85px;height:21px;background:url(/images/returnHome.gif);text-align:center""><a href=""/"" class=""menuwhite"">回到首页</a></div>" & vbcrlf 
tcontent=tcontent&"</div>" & vbcrlf 
tcontent=tcontent&"</p>" & vbcrlf 
tcontent=tcontent&"</div>" & vbcrlf 
tcontent=tcontent&"</div>" & vbcrlf 
tcontent=tcontent&"</span>" & vbcrlf 
tcontent=tcontent&"</body>" & vbcrlf 
tcontent=tcontent&"<script>" & vbcrlf 
tcontent=tcontent&" window.scrollTo(0,0);" & vbcrlf 
tcontent=tcontent&" var bo = document.getElementsByTagName('body')[0];" & vbcrlf 
tcontent=tcontent&" var ht = document.getElementsByTagName('html')[0];" & vbcrlf 
tcontent=tcontent&" var boht = document.getElementById('errboxs');  " & vbcrlf 
tcontent=tcontent&" bo.style.height='auto';" & vbcrlf 
tcontent=tcontent&" bo.style.overflow='auto';" & vbcrlf 
tcontent=tcontent&" ht.style.height='auto'; " & vbcrlf 
tcontent=tcontent&"  bo.style.height='100%';" & vbcrlf 
tcontent=tcontent&"  bo.style.overflow='hidden';" & vbcrlf 
tcontent=tcontent&"  ht.style.height='100%';  " & vbcrlf 
tcontent=tcontent&"</script>" & vbcrlf 
tcontent=tcontent&"</html>" & vbcrlf 
response.Write(tcontent)
call closeconn()
response.End()
tcontent="":msg=""
End Sub

'输出提示信息 
'msg 提示的信息内容 
'href 1)为空时返回上一页 2)为0时不返回,即程序继续执行 3)为1为关闭当前窗口 4)为2时，父窗口返回链接 msg信息格式：信息提示内容|打开链接(注意格式，否则会出错) 5)其它值为返回到指定的网址
'flag 是否关闭数据库对象
sub alert(msg,href,flag)
dim tmparr
select case href
case ""
href="history.back();"
case "0"
href=""
case "1"
href="window.close(this);"
case "2"
tmparr=split(msg,"|")
msg=tmparr(0)
href="parent.location.href='"&tmparr(1)&"'"
erase tmparr
case else
href="location.href='"&href&"';"
end select
response.Write("<script language=""javascript"">alert('"&msg&"');"&href&"</script>")
if flag=1 then
	call closeconn()
	response.End()
end if
msg="":flag=""
end sub

'检查是否登录 cookiesname:保存cookies的变量名
function chkLogin(cookiesname)
	chkLogin=request.Cookies(cookiesname).HasKeys
end function

'================================================= 
  '函数名：JmailSend 
  '作 用：用Jmail发送邮件 
  '参 数：MailTo收件人地址 Subject邮件标题 Body内容 
  '返回值：flase 发送失败　true 发送成
function JmailSend(MailTo,Subject,Body,FromName)
	on error resume next 
	dim JmailMsg,From,Smtp,Username,Password
	' Body 邮件内容    
	' MailTo 收件人Email 
	' From 发件人Email 
	' FromName 发件人姓名 
	' Smtp smtp服务器 
	' Username 邮箱用户名 
	' Password 邮箱密码 
	From="root@xsp2.com"
	'FromName="小商品资源网"
	Smtp="mail@xsp2.com"
	Username="root@xsp2.com"
	Password="wujinchen"
	
	set JmailMsg=server.createobject("jmail.message") 
	if err then
		JmailSend=1
		exit function
	end if
	JmailMsg.charset="uft-8" 
	JmailMsg.ContentType = "text/html" 
	JmailMsg.ISOEncodeHeaders = False '是否进行ISO编码，默认为True
	JmailMsg.AddHeader "Originating-IP", "115.28.24.81"
	JmailMsg.from=From
	JmailMsg.AddRecipient MailTo 
	JmailMsg.MailDomain=Smtp      
	JmailMsg.MailServerUserName =Username
	JmailMsg.MailServerPassWord =Password     
	JmailMsg.fromname=FromName     
	JmailMsg.subject=Subject 
	JmailMsg.body=Body       
	if JmailMsg.send("mail.xsp2.com")   then
		JmailSend=0
	else
		JmailSend=2
	end if
	JmailMsg.close 
	set JmailMsg=nothing    
	Body="":Subject=""
end function 

'********************************************
'函数名：getRndNum
'作  用：产生指定长度的随机数字
'参  数：maxLen ----随机数的长度
'返回值：产生的随机数
'********************************************
function   getRndNum(maxLen)   
  Dim   strNewPass   
  Dim   whatsNext,   upper,   lower,   intCounter   
  Randomize      
  For   intCounter   =   1   To   maxLen   
  whatsNext   =   Int((1   -   0   +   1)   *   Rnd   +   0) 
  upper   =   57   
  lower   =   48   
  strNewPass   =   strNewPass   &   Chr(Int((upper   -   lower   +   1)   *   Rnd   +   lower))   
  Next   
  getRndNum   =   strNewPass    
  strNewPass=""   
end   function

'获取指定长度随机数字字母组合 不会生成 0 1 o O l这些字符
function getRnd(num)
dim char_array(62)
char_array(0)="2"
char_array(1) = "2"
char_array(2) = "2"
char_array(3) = "3"
char_array(4) = "4"
char_array(5) = "5"
char_array(6) = "6"
char_array(7) = "7"
char_array(8) = "8"
char_array(9) = "9"
char_array(10) = "a"
char_array(11) = "b"
char_array(12) = "c"
char_array(13) = "d"
char_array(14) = "e"
char_array(15) = "f"
char_array(16) = "g"
char_array(17) = "h"
char_array(18) = "i"
char_array(19) = "j"
char_array(20) = "k"
char_array(21) = "L"
char_array(22) = "m"
char_array(23) = "n"
char_array(24) = "f"
char_array(25) = "p"
char_array(26) = "q"
char_array(27) = "r"
char_array(28) = "s"
char_array(29) = "t"
char_array(30) = "u"
char_array(31) = "v"
char_array(32) = "w"
char_array(33) = "x"
char_array(34) = "y"
char_array(35) = "z"
char_array(36) = "s"
char_array(37) = "A"
char_array(38) = "B"
char_array(39) = "C"
char_array(40) = "D"
char_array(41) = "E"
char_array(42) = "F"
char_array(43) = "G"
char_array(44) = "H"
char_array(45) = "I"
char_array(46) = "J"
char_array(47) = "K"
char_array(48) = "L"
char_array(49) = "M"
char_array(50) = "N"
char_array(51) = "D"
char_array(52) = "P"
char_array(53) = "Q"
char_array(54) = "R"
char_array(55) = "S"
char_array(56) = "T"
char_array(57) = "U"
char_array(58) = "V"
char_array(59) = "W"
char_array(60) = "X"
char_array(61) = "Y"
char_array(62) = "Z"
dim i,tmpnum,rndnum
getRnd=""
for i=1 to num
Randomize
rndnum=cint(rnd()*62)
tmpnum=char_array(rndnum)
getRnd=getRnd&tmpnum
next
erase char_array
end function


'得到价格 如果价格为0，刚返回面议 para价格 sdot小数位数
function getPrice(byval para,sdot)
para=trim(para)
if isnull(para) or para="" or len(para)<1 or not isnumeric(para) then
	getPrice=0
	exit function
end if
para=cdbl(para)
if para=0 then
	getPrice=0
	exit function
end if
para=formatnumber(para,sdot)
if left(cstr(para),1)="." then para="0"+cstr(para) else para=cstr(para)
getPrice=para
para=""
end function

'删除地址中的参数
'surl 地址
'spara 要查找的参数 多个参数以,号分隔
'返回格式： 如果长度大于2 返回?a=2& 的形式 否则 返回 ?
function delUrlpara(byVal surl,spara)
if surl="" or len(surl)<1 or isnull(surl) or isempty(surl) then
delUrlpara="?"
exit function
end if
surl=lcase(surl)
spara=lcase(spara)
dim paraArr,j
dim urlarr,tmparr,t
paraArr=split(spara,",")

for j=0 to ubound(paraArr)
	'如果参数可能存在
	if instr(surl,paraArr(j)&"=")>0 then
		'分离所有参数
		urlarr=split(surl,"&")	
		t=""
		for i=0 to ubound(urlarr)
			'参数及值可能有效（也许参数无效但是有=号）
			if not chkNull(urlarr(i),1) and instr(urlarr(i),"=")>0 then
				tmparr=split(urlarr(i),"=")
				'如果参数不是spara,即要删除的参数 并且是有效参数(字母开头 4-20位)，则重组参数
				if tmparr(0)<>paraArr(j) and not chknull(tmparr(0),1)  then
					if t="" then t=tmparr(0)&"="&tmparr(1) else t=t&"&"&tmparr(0)&"="&tmparr(1)					
				end if
				erase tmparr
			end if
		next
		erase urlarr
		surl=t
	end if	
next

erase paraArr
if chkNull(surl,1) then surl="?" else surl="?"&surl&"&"
delUrlpara=surl
end function

'获取远程的IP地址
Function GetIP()
Dim Temp
Temp = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Temp ="" or isnull(Temp) or isEmpty(Temp) Then Temp = Request.ServerVariables("REMOTE_ADDR")
If Instr(Temp,"'")>0 Then Temp="0.0.0.0"
GetIP = Temp
Temp=""
End Function

'****************************************************
	'函数名：ChkPost
	'作  用：禁止站外提交表单
	'返回值：true站内提交，flase站外提交
'****************************************************
Function ChkPost()
		Dim url1,url2
		chkpost=true
		url1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		url2=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Mid(url1,8,Len(url2))<>url2 Then
			 chkpost=false
			 exit function
		End If
		url1="":url2=""
End function

'---------------------数据处理---------------------------
'****************************************************
'函数名：getDbValue
'作  用：取一条记录
'参  数：Sql           传入参数 Sql语句
'       ArrayCount    传入参数 需要的字段的个数
'返回值:getDbValue      输出参数 为一个数组，数组的第一项为是否取到数据，如果有数据则为大于0的数字，否则为0。参数一次类推为字段
'****************************************************
Function getDbValue(Sql,ArrayCount)
	ReDim NewArray(ArrayCount)
	Dim Recordcounts,Rs,i,Comm
	Set Comm = Server.CreateObject("ADODB.Command")
	With Comm
		Comm.ActiveConnection = Conn
		Comm.CommandType = 4
		Comm.CommandText = "ColumnGet"
		Comm.Prepared  = true
		Comm.Parameters.Append .CreateParameter("Return",2,4)
		Comm.Parameters.Append .CreateParameter("@Sql",200,1,1000,Sql)
		Set Rs = .Execute	
	End With
	Rs.Close
	Recordcounts=Comm(0)
	NewArray(0)=Recordcounts
	If Recordcounts=0 Then getDbValue=NewArray:Exit Function
	Set Comm=Nothing
	Rs.Open
	For i = 0 to cint(ArrayCount-1)
		NewArray(i+1)=Rs(i)
	Next
	Rs.Close
	getDbValue=NewArray
	Erase NewArray
	Sql =""
End Function


'****************************************************
'函数名：getDBValueList
'作  用：按指定sql语句取记录集
'参  数：Sql           传入参数 Sql语句
'返回值:getDBValueList      输出参数 为一个二维数组，数组的第一项为是否取到数据，如果有数据则为大于0的数字，否则为0。参数一次类推为字段
'****************************************************
Function getDBValueList(Sql)

	Dim Comm,Rs
	Dim Recordcounts
	Dim ArrayRecord
	Set Comm=Server.CreateObject("ADODB.Command")
	With Comm
	.ActiveConnection = Conn
	.CommandType = 4
	.CommandText = "Columnget"
	.Prepared = true
	.Parameters.Append .CreateParameter("@Return",3,4)
	.Parameters.Append .CreateParameter("@Sql",200,1,2000,Sql)'参数1为定义的变量 2为数据类型 3为数据宽度 4为值
	Set Rs=.Execute()
	End With
	Rs.Close 
	RecordCounts=Comm(0)
	Rs.Open	
	If RecordCounts=0 Then 
		Redim ArrayRecord(0,0):ArrayRecord(0,0)=0
	Else
		ArrayRecord=Rs.GetRows()
	end if
	Rs.Close
	getDBValueList=ArrayRecord
	erase ArrayRecord
	Sql =""
End Function

'设置数据数据库 插入数据 修改数据 删除数据
'返回值：0为不成功 或者 没有影响行数 大于0的值，是当前操作影响的行数 如果是插入操作，刚返回插入后的ID号
Function setDBValue(Sql)
	'On Error Resume Next
	Dim Comm,idnum
	dim tmparr
	Set Comm=Server.CreateObject("ADODB.Command")
	With Comm
	.ActiveConnection = Conn
	.CommandType = 4
	.CommandText = "ColumnSet"
	.Prepared = true
	.Parameters.Append .CreateParameter("@Return",3,4)
	.Parameters.Append .CreateParameter("@Sql",200,1,2000,Sql)
	.Execute()
	End With
	if err then	
		setDBValuse=0
		err.clear		
	else
		setDBValue = Comm(0)	
	end if
	Set Comm = Nothing	
	Sql=""
End Function

'判断文件是否存在
Function isExists(byVal FileName) 
On Error Resume Next
filename=server.MapPath(filename)
if err then
	isExists=false
	err.clear
	filename=""
	exit function
end if	
dim delobj
set delobj=server.CreateObject("Scripting.FileSystemObject")
if delobj.FileExists(filename) then
	isExists=true
else
	isExists=false
end if
set delobj=nothing
filename=""
End Function 

'检查是否符合 1,3 （数字,数字的格式）
function chkPara(byVal para)
	if para="" or isnull(para) or isempty(para) then
		chkPara=false
		exit function
	end if
	dim regEx
	set regEx=new regexp
	regEx.global=true
	regEx.pattern="^\d+(,?\d+)*$"
	chkPara=regEx.test(para)
	set regEx=nothing
end function

'获取网站网址
Function GetWebUrl()
Dim WebUrl,Temp_Lodo_Str
WebUrl=Request.ServerVariables("URL")
WebUrl=Left(WebUrl,InStrRev(WebUrl,"/")-1)
If InStrRev(WebUrl,"/")>0 Then WebUrl=Left(WebUrl,InStrRev(WebUrl,"/")-1)
if Trim(Request.ServerVariables("SERVER_PORT"))<>"80" Then WebUrl=":"&Trim(Request.ServerVariables("SERVER_PORT"))&WebUrl
WebUrl=Request.ServerVariables("SERVER_NAME")&WebUrl&"/"
Temp_Lodo_Str=split(Request.ServerVariables("SERVER_NAME"),".")
if Ubound(Temp_Lodo_Str)=1 Then
if Trim(Lcase(Temp_Lodo_Str(0)))<>"www" and (Not IsNumeric(Trim(Lcase(Temp_Lodo_Str(0))))) and Instr(Lcase(WebUrl),"localhost")<=0 Then WebUrl="www."&WebUrl
End If
GetWebUrl=WebUrl
End Function

'自动将短域名更换为长域名 并实现301定向
function changeSwebUrl()
	dim url,pageUrl
	dim f,fs
	dim i,finded
	
	'要检测的短域名 如果在左边不存在下面数组里面的值，则重定向到http://www.域名
	f=array("www","www1")
	url=lcase(Request.ServerVariables("SERVER_NAME"))
	if url="localhost" then exit function '本地浏览，退出函数	
	finded=false
	'检查是否存在域名中
	for i=0 to ubound(f)
		fs=f(i)&"."
		if left(url,len(fs))=fs then
			finded=true
			exit for
		end if
	next	
	'如果没有则定向
	if not finded then
		Response.Status="301 Moved Permanently" 
		Response.AddHeader "Location","http://www."&url&"/" 
		Response.End
	end if		
end function

sub chkSql()
	dim sql_leach,sql_leach_0,Sql_DATA 
	sql_leach = "and,exec,insert,select,delete,update,count,chr,mid,master,truncate,char,declare" 
	sql_leach_0 = split(sql_leach,",") 
	
	
	If Request.QueryString<>"" Then 
	For Each SQL_Get In Request.QueryString 
	For SQL_Data=0 To Ubound(sql_leach_0) 
	if instr(Request.QueryString(SQL_Get),sql_leach_0(Sql_DATA))>0 Then 
	Response.Write "参数包含敏感字符" 
	Response.end 
	end if 
	next 
	Next 
	End If 
	
	If Request.Form<>"" Then 
	For Each Sql_Post In Request.Form 
	For SQL_Data=0 To Ubound(sql_leach_0) 
	if instr(Request.Form(Sql_Post),sql_leach_0(Sql_DATA))>0 Then 
	Response.Write "参数包含敏感字符" 
	Response.end 
	end if 
	next 
	next 
	end if 
end sub

'获取MDB数据库的一行记录 返回一维数组 
'arr(0) 为-1是查询失败 arr(1)存储错误代码 arr(2)存储查询的sql语句
'arr(0) 为0时查询成功，数据库里面没有记录 arr(2)存储查询的sql语句
'arr(0) 为1时查询成功并返回相关记录
function getmdbvalue(tsql)
on error resume next
dim trs,ti
redim tmparr(2)
tmparr(0)=0
tmparr(1)=""
tmparr(2)=""
set trs=server.CreateObject("adodb.recordset")
trs.open tsql,conn,1,1
if err then
	tmparr(0)=-1
	tmparr(1)=err.number&":"&err.description&" 可能Sql查询语句错误"	
	tmparr(2)=tsql
else
	if trs.bof and trs.eof then
		tmparr(2)=tsql
		getmdbvalue=tmparr
	else	
		tmparr(0)=1
		for ti=1 to trs.fields.count
			redim preserve tmparr(ti)
			tmparr(ti)=trs(ti-1)
		next
	end if
	trs.close
end if
set trs=nothing
getmdbvalue=tmparr:erase tmparr
end function

'获取MDB数据库的列表记录 返回二维数组 
'arr(0,0) 为-1是查询失败 arr(0,1)存储错误代码 arr(0,2)存储查询的sql语句
'arr(0,0) 为0时查询成功，数据库里面没有记录 arr(2)存储查询的sql语句
'arr(0,0) 为1是查询成功并返回相关记录
function getmdbvaluelist(tsql)
on error resume next
dim trs,ti
dim tmparr(0,2)
tmparr(0,0)=0
tmparr(0,1)=""
tmparr(0,2)=""
set trs=server.CreateObject("adodb.recordset")
trs.open tsql,conn,1,1
if err then	
	tmparr(0,0)=-1
	tmparr(0,1)=err.number&":"&err.description&" 可能Sql查询语句错误"
	tmparr(0,2)=tsql
	getmdbvaluelist=tmparr		
else
	if trs.bof and trs.eof then
		tmparr(0,0)=0
		tmparr(0,1)=""
		tmparr(0,2)=tsql
		getmdbvaluelist=tmparr
	else	
		getmdbvaluelist=trs.getrows()
	end if
	trs.close
end if
set trs=nothing
erase tmparr
end function

function setmdbvalue(tsql)
conn.execute(tsql)
end function

%>
