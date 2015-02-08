<%
Rem 过滤SQL非法字符
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	End if
	checkStr=replace(str,"'","''")
End function

Rem 判断发言是否来自外部
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	End if
End function

Rem 过滤表单字符
function HTMLcode(fString)
if Not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
End if
End function

Rem 过滤HTML代码
function HTMLEncode(fString)
if Not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    HTMLEncode = fString
End if
End function
'#######################检测是否为字母或数字函数##################
function checkletter(str)
   str=LCase(str)
   dim i
   For i = 1 To Len(str)
    if  Asc(Mid(str, i, 1)) < 97  or   Asc(Mid(str, i, 1)) > 122  then
	    if  Asc(Mid(str, i, 1)) < 48  or   Asc(Mid(str, i, 1)) > 57 then
	         checkletter=false
			 exit function
	    else 
		     checkletter=true
	    End if
	else
	    checkletter=true
	End if
   Next
End function
'#########################################################
sub checkdate(strtime)
 if isdate(strtime)=false then
  response.write "<script language='javascript'>" & chr(13)
    response.write "alert('日期输入错误!格式如：2004-10-10');" & Chr(13)
    response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
    response.write "</script>" & Chr(13)
    Response.End
  End if 
End sub
'######################邮箱检测##############################

function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
End if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   End if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and Not IsNumeric(c) then
       IsValidEmail = false
       exit function
     End if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   End if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
End if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
End if
if InStr(email, "..") > 0 then
   IsValidEmail = false
End if

End function

'########################随机数#############################
function rndnum()
dim cs,cs0,cnum,i,cs1
cs="0,1,2,3,4,5,6,7,8,9,a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z"
cnum=20
cs0=split(cs,",")
randomize
for i=1 to cnum
randomize
cs1=cs1&cs0(int(61*rnd))'cs0为数组
next
rndnum=cs1
End function
'#########################发邮件####################################
%>