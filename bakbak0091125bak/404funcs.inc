<%
'404α��̬ר�ú���
'���ߣ�Anligeԭ��

class Html404
 private regUrl,queryS,url
 Public Property Get Pertrol()
   Pertrol = searchStr(2)
 End Property
 
 Public Property Get Domain()
   Domain = searchStr(4)
 End Property
 
 Public Property Get Port()
   Port = searchStr(6)
 End Property
 
 Public Property Get Path()
   Path = searchStr(9)
 End Property
 
 Public Property Get Strings()
   Strings = searchStr(11)
 End Property
 
 Public Property Get fname()
	dim bAry
	if instr(searchStr(9),"/")>0 then
		bAry=split(searchStr(9),"/")
		fname=bAry(ubound(bAry))
	else
		fname=searchStr(9)
	end if
 End Property
 
 Private Sub Class_Initialize()
	regUrl="((http|ftp|https)\:\/\/)(([^\s\:\?\/]+)(\:([\d]+))?(\/(([^\s\?]+)?(\?([^\s]+))?)?)?)?"
	queryS=request.QueryString
	url=split(queryS,";")(1) 
 End Sub

 Private Sub Class_Terminate()
 
 End Sub
 
 Private function searchStr(index)
	dim objRegExp
    set objRegExp=New RegExp
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = regUrl
    searchStr = objRegExp.Replace(url, "$" & index)
    Set objRegExp = Nothing
 end function
 
'������:isAPI
'����:ƥ�����򣬷���ƥ���滻���url
'����:source,Ҫ�������ִ���regFrom,Ҫƥ�������regTo,Ҫ�滻ƥ�����ݵ��ִ�
Public Function isAPI(byval source,ByVal regFrom,byval regTo)
    Dim objRegExp, Match, Matches
    set objRegExp=New RegExp
    Dim str: str = source
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = regFrom
    str = objRegExp.Replace(str, regTo)
    Set objRegExp = Nothing
    isAPI = str
End Function

'������:getString
'����:��ȡ����
'����:index����������������һ���ļ���cat_11_5.html,��getString(1)=11,getString(2)=5;getString(0)=��Ȼ����cat�ˣ�����ô�����
function getString(index)
	dim strs
	dim str:str=split(fname,".")(0)
	if instr(str,"_")>0 then
		strs=split(str,"_")(index)
	else
		strs=""
	end if
	getString=strs
end function

'������:vCount
'����:��ȡ������������������_���ָ���
'����:�ޣ������Զ����
function vCount()
	dim c
	dim str:str=split(fname,".")(0)
	if instr(str,"_")>0 then
		c=ubound(split(str,"_"))
	else
		c=0
	end if
	vCount=c
end function

 Public function Redir(url,Desc)
  dim tempStr:tempStr=""
  tempStr=tempStr&"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"&vbcrlf
  tempStr=tempStr&"<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""zh-cn"" lang=""zh-cn"">"&vbcrlf
  tempStr=tempStr&" <head>"&vbcrlf
  tempStr=tempStr&"     <meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"&vbcrlf
  tempStr=tempStr&"     <meta name=""verify-v1"" content=""NZP+k2TE7UPnlA32XcbHowT2LZZRyiOus7icxcZDsdA="" />"&vbcrlf
  tempStr=tempStr&"     <meta name=""AUTHOR"" content=""Anlige,zhanghuiguoanlige@126.com"" />"&vbcrlf
  tempStr=tempStr&"     <meta name=""Copyright"" content=""copyright 2008 www.ii-home.cn"" />"&vbcrlf
  tempStr=tempStr&"     <meta http-equiv=""Refresh"" content=""3;URL=" & url & """ />"&vbcrlf
  tempStr=tempStr&"     <title>���Ժ�,ת����...</title>"&vbcrlf
  tempStr=tempStr&" </head>"&vbcrlf
  tempStr=tempStr&" <body>" & Desc & ":3���ת��" & url & "...</body></html>"&vbcrlf
  response.write tempStr
 end function
 
end class
%>