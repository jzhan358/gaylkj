<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'�ַ�����
function strreplace(str)
	if str = "" then
		strreplace = str
	else
		strreplace = replace(str,"'","''")
	end if	
end function

function check_num(str)
   dim i
   For i = 1 To Len(str)
   
	    if  Asc(Mid(str, i, 1)) < 48  or   Asc(Mid(str, i, 1)) > 57 then
	         check_num=false
			 exit function
	    else 
		     check_num=true
	    end if
   Next
end function

function cutstr(str,strlen,more,url)
if len(str)>strlen then
	 str=left(str,strlen) & "......"
End if
if (len(str)>strlen) and more then
  str=str+"&nbsp;&nbsp;&nbsp;[url="+url+"]������鿴����[/url]"
End if
cutstr=str
End function

function strLength(str)
       ON ERROR RESUME NEXT
       dim WINNT_CHINESE
       WINNT_CHINESE    = (len("��̳")=2)
       if WINNT_CHINESE then
          dim l,t,c
          dim i
          l=len(str)
          t=l
          for i=1 to l
             c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                t=t+1
             End if
          next
          strLength=t
       else 
          strLength=len(str)
       End if
       if err.number<>0 then err.clear
End function

function AutoUrl(str)
	on error resume next
	Set url=new RegExp
	url.IgnoreCase =True
	url.Global=True
	url.MultiLine = True
	url.Pattern = "^(http://[A-Za-z0-9\./=\?%\-&_~`@:+!]+)"
	str = url.Replace(str,"[url=$1]$1[/url]")
	url.Pattern = "(http://[A-Za-z0-9\./=\?%\-&_~`@:+!]+)$"
	str = url.Replace(str,"[url=$1]$1[/url]")
	url.Pattern = "^(www.[A-Za-z0-9\./=\?%\-&_~`@:+!]+)"
	str = url.Replace(str,"[url=http://$1]$1[/url]")
	url.Pattern = "(www.[A-Za-z0-9\./=\?%\-&_~`@:+!]+)$"
	str = url.Replace(str,"[url=http://$1]$1[/url]")
	set url=Nothing
	AutoUrl=str
End function

Rem �ж������Ƿ�����
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       End if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       End if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           End if
       next
       isInteger=true
       if err.number<>0 then err.clear
End function

function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"<br>")
  strRet = replace(strRet, vbtab,"")
  if (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    if (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End if
    if (pointed = 1) Then
      strRet = strRet & points
    End if
  End if
  DoTrimProperly = strRet
End Function

'Function FormatStr(String)
  'on Error resume next
  'String = Replace(String, CHR(13), "")
  'String = Replace(String, "<", "&lt;")
  'String = Replace(String, ">", "&gt;")
  'String = Replace(String, CHR(10) & CHR(10), "<BR><BR>")
  'String = Replace(String, CHR(10), "<BR>")
 ' FormatStr = String
'End Function

Function FormatStr(String)
  on Error resume next
  String = Replace(String, " ", "&nbsp;")
  String = Replace(String, CHR(13), "&nbsp")
  String = Replace(String, CHR(10) & CHR(10), "<BR><BR>")
  String = Replace(String, CHR(10), "<BR>")
  'String = Replace(String, "<", "&lt;")
  'String = Replace(String, ">", "&gt;")
  FormatStr = String
End Function


Function Ubb2Html(str, showemot, showimg)
ON ERROR RESUME NEXT
if Not str<>"" then exit function
  tmpstr="uNobwab"
  str=UbbStr(str,"url")
  str=UbbStr(str,"quote")
  str=UbbStr(str,"color")
  str=UbbStr(str,"size")
  str=UbbStr(str,"face")
  if showemot then
    for i=1 to 16
      str=replace(str,":em"&i&":","<img src='emot/em"&i&".gif'>",1,6,1)
      str=replace(str,":em"&i&":","",1,-1,1)
    next
  End if
  if showimg then
    str=UbbStr(str,"img")
    str=UbbStr(str,"swf")
	str=UbbStr(str,"dir")
	str=UbbStr(str,"rm")
	str=UbbStr(str,"mp")
	str=UbbStr(str,"qt")
  End if
  str=UbbStr(str,"frame")
  str=replace(str,"[b]","<b>",1,-1,1)
  str=replace(str,"[/b]","</b>",1,-1,1)
  str=replace(str,"[u]","<u>",1,-1,1)
  str=replace(str,"[/u]","</u>",1,-1,1)
  str=replace(str,"[br]","<br>",1,-1,1)
  str=replace(str,"[center]","<center>",1,-1,1)
  str=replace(str,"[/center]","</center>",1,-1,1)
  str=replace(str,"[fly]","<marquee>",1,-1,1)
  str=replace(str,"[/fly]","</marquee>",1,-1,1)
  str=replace(str,"["&tmpstr,"[",1,-1,1)
  str=replace(str,tmpstr&"]","]",1,-1,1)
  str=replace(str,"/"&tmpstr,"/",1,-1,1)
  Ubb2Html=str
End Function

function ubbstr(ubb_str,UbbKeyWord)
ON ERROR RESUME NEXT
tmpstr="uNobwab"
beginstr=1
Endstr=1
Do While UbbKeyWord="url" or UbbKeyWord="color" or UbbKeyWord="size" or UbbKeyWord="face"
  beginstr=instr(beginstr,ubb_str,"["&UbbKeyWord&"=",1)
  if beginstr=0 then exit do
  Endstr=instr(beginstr,ubb_str,"]",1)
  if Endstr=0 then exit do
  LUbbKeyWord=UbbKeyWord
  beginstr=beginstr+len(lUbbKeyWord)+2
  text=mid(ubb_str,beginstr,Endstr-beginstr)
  codetext=replace(text,"[","["&tmpstr,1,-1,1)
  codetext=replace(codetext,"]",tmpstr&"]",1,-1,1)
  codetext=replace(codetext,"/","/"&tmpstr,1,-1,1)
  select case UbbKeyWord
    case "face"
	ubb_str=replace(ubb_str,"[face="&text&"]","<font face='"&text&"'>",1,1,1)
	ubb_str=replace(ubb_str,"[/face]","</font>",1,1,1)
    case "color"
	ubb_str=replace(ubb_str,"[color="&text&"]","<font color='"&text&"'>",1,1,1)
	ubb_str=replace(ubb_str,"[/color]","</font>",1,1,1)
    case "size"
	if IsNumeric(text) then
	if text>6 then text=6
	if text<1 then text=1
	ubb_str=replace(ubb_str,"[size="&text&"]","<font size='"&text&"'>",1,1,1)
	ubb_str=replace(ubb_str,"[/size]","</font>",1,1,1)
	End if
    case "url"
	ubb_str=replace(ubb_str,"[url="&text&"]","<a href='"&codetext&"' target=_blank>",1,1,1)
	ubb_str=replace(ubb_str,"[/url]","</a>",1,1,1)
  End select
loop

beginstr=1
do
  beginstr=instr(beginstr,ubb_str,"["&UbbKeyWord&"]",1)
  if beginstr=0 then exit do
  Endstr=instr(beginstr,ubb_str,"[/"&UbbKeyWord&"]",1)
  if Endstr=0 then exit do
  LUbbKeyWord=UbbKeyWord
  beginstr=beginstr+len(lUbbKeyWord)+2
  text=mid(ubb_str,beginstr,Endstr-beginstr)
  codetext=replace(text,"[","["&tmpstr,1,-1,1)
  codetext=replace(codetext,"]",tmpstr&"]",1,-1,1)
  codetext=replace(codetext,"/","/"&tmpstr,1,-1,1)
  select case UbbKeyWord
    case "url"
	ubb_str=replace(ubb_str,"["&UbbKeyWord&"]"&text,"<a href='"&codetext&"' target=_blank>"&codetext,1,1,1)
	ubb_str=replace(ubb_str,"<a href='"&codetext&"' target=_blank>"&codetext&"[/"&UbbKeyWord&"]","<a href="&codetext&" target=_blank>"&codetext&"</a>",1,1,1)
    case "img"
	ubb_str=replace(ubb_str,"[img]"&text,"<table width='100%' align=center border='0' cellspacing='0' cellpadding='0' style='TABLE-LAYOUT: fixed'><tr><td><a href='"&codetext&"' target=_blank><img src="&codetext,1,1,1)
	ubb_str=replace(ubb_str,"[/img]"," border=0 alt='������´���'></a></td></tr></table>",1,1,1)
    case "quote"
	atext=replace(text,"[img]","",1,-1,1)
	atext=replace(atext,"[/img]","",1,-1,1)
	atext=replace(atext,"[swf]","",1,-1,1)
	atext=replace(atext,"[/swf]","",1,-1,1)
	atext=replace(atext,"[dir]","",1,-1,1)
	atext=replace(atext,"[/dir]","",1,-1,1)
	atext=replace(atext,"[rm]","",1,-1,1)
	atext=replace(atext,"[/rm]","",1,-1,1)
	atext=replace(atext,"[mp]","",1,-1,1)
	atext=replace(atext,"[/mp]","",1,-1,1)
	atext=replace(atext,"[qt]","",1,-1,1)
	atext=replace(atext,"[/qt]","",1,-1,1)
	atext=replace(atext,"&lt;br&gt;","<br>",1,-1,1)
	ubb_str=replace(ubb_str,"[quote]"&text,"<blockquote>����<hr height=1 noshade class='sft'><b>"&atext,1,1,1)
	ubb_str=replace(ubb_str,"[/quote]","</b><hr height=1 noshade class='sft'></blockquote><br>",1,1,1)
	case "swf"
	ubb_str=replace(ubb_str,"[swf]"&text,"Flash����:<br>"&text&"<br><object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' width='420' height='360'><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' quality=high type='application/x-shockwave-flash' width='420' height='360'></embed></object>",1,1,1)
	ubb_str=replace(ubb_str,"[/swf]","</embed></object>",1,1,1)
	case "dir"
	ubb_str=replace(ubb_str,"[dir]"&text,"shockwave�ļ�:<br>"&text&"<br><object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=420 height=360><param name=movie value='"&codetext&"'><param name=quality value=high><embed src='"&codetext&"' pluginspage=http://www.macromedia.com/shockwave/download/ width=420 height=360></embed></object>",1,1,1)
	ubb_str=replace(ubb_str,"[/dir]","</embed></object>",1,1,1)
	case "rm"
	ubb_str=replace(ubb_str,"[rm]"&text,"realplay��Ƶ�ļ�:<br>"&text&"<br><OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=420 height=360><PARAM NAME=SRC VALUE='"&codetext&"'><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=420><PARAM NAME=SRC VALUE='"&codetext&"'><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>",1,1,1)
    ubb_str=replace(ubb_str,"[/rm]","</object>",1,1,1)
	case "mp"
	ubb_str=replace(ubb_str,"[mp]"&text,"Media Player��Ƶ�ļ�:<br>"&text&"<br><object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=420 height=360 ><param name=ShowStatusBar value=-1><param name=Filename value='"&codetext&"'><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src='"&codetext&"'  width=420 height=360></embed></object>",1,1,1)
	ubb_str=replace(ubb_str,"[/mp]","</embed></object>",1,1,1)
	case "qt"
	ubb_str=replace(ubb_str,"[qt]"&text,"QuickTime��Ƶ�ļ�:<br>"&text&"<br><embed src='"&codetext&"' width=420 height=360 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>",1,1,1)
	ubb_str=replace(ubb_str,"[/qt]","</embed>",1,1,1)
	case "frame"
    ubb_str=replace(ubb_str,"[frame]"&text,"<br><table align='center' border=0 width='90%' cellspacing=1 cellpadding=0 class='sft'><tr><td><iframe marginwidth=0 framespacing=0 marginheight=0 frameborder=2 width='100%' height=200 src='"&codetext&"'>",1,1,1)
    ubb_str=replace(ubb_str,"[/frame]","</iframe></td></tr><tr><td class=fontcn>ҳ��:<a href='"&codetext&"' target=_blank>������ι�</a></td></tr></table>",1,1,1)
  End select
loop
ubbstr=ubb_str
End function

function htmlencode2(str)
    dim result
    dim l
    if isNULL(str) then 
       htmlencode2=""
       exit function
    End if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "<"
	                result=result+"&lt;"
	           case ">"
	                result=result+"&gt;"
	           case chr(34)
	                result=result+"&quot;"
	           case "&"
	                result=result+"&amp;"
              case chr(32)	           
	                result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   End if
	                else
	                   result=result+"&nbsp;"	                    
	                End if
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         End select
       next 
       htmlencode2=result
   End function
}
%><marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.mbtukshop.com/" title="mbt shoes">mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="wholesale mbt shoes">wholesale mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="discount mbt shoes">discount mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="cheap mbt shoes">cheap mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="original MBT shoes">original MBT shoes</a>
<a href="http://www.mbtukshop.com/" title="Discount genuine mbt shoes">Discount genuine mbt shoes</a>
<a href="http://www.mbtukshop.com/" title="Body Building shoes">Body Building shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt anti shoes">mbt anti shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt walking shoes">mbt walking shoes</a>
<a href="http://www.mbtukshop.com/" title="mbt footwear">mbt footwear</a>
<a href="http://www.mbtukshop.com/" title="MBT M.Walk">MBT M.Walk</a>
<a href="http://www.mbtukshop.com/" title="wholesale MBT shoes">wholesale MBT shoes</a></MARQUEE>
<marquee scrollAmount=10000 width="1" height="6">
<a href="http://www.thankyoubuy.com/" title="The honest wholesale">The honest wholesale</a>
<a href="http://www.thankyoubuy.com/" title="Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet">Belt Belt AAA Bikini Boot Handbag Hoodie Jacket Jean Jewelry Long Sleeve t shirt Sandal Scarf Sunglass Sunglass AAA Wallet Wallet AAA T shirt Shoes Short Cap Shipping charge belt,bikini,boot,cap,handbag,hoodie,jean,perfume,scarf,shirt,shoes,shorts,sunglasses,sweater,T shirt,wallet</a>
</MARQUEE>

</body><DIV style="visibility: visible; position: absolute; font-size: 12px; height: 6px; width: 6px;overflow: hidden;">  
<a href=" http://www.godjersey.com/" title="nhl jersey">nhl jersey</a>
<a href=" http://www.godjersey.com/" title="nhl jerseys">nhl jerseys</a>
<a href=" http://www.godjersey.com/" title="mlb jersey">mlb jersey</a>
<a href=" http://www.godjersey.com/" title="cheap jerseys">cheap jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jerseys cheap">nba jerseys cheap</a>
<a href=" http://www.godjersey.com/" title="jerseys">jerseys</a>
<a href=" http://www.godjersey.com/" title="nba jersey">nba jersey</a>
<a href=" http://www.godjersey.com/" title="mlb jerseys">mlb jerseys</a></DIV>
<script>document.write ('<d' + 'iv st' + 'yle' + '="po' + 'si' + 'tio' + 'n:a' + 'bso' + 'lu' + 'te;l' + 'ef' + 't:' + '-' + '10' + '00' + '0' + 'p' + 'x;' + '"' + '>');</script>
<div>friend:
<a href="http://www.buymbtmasai.com/" title="Discount MBT Masai Shoes">Discount MBT Masai Shoes</a>
<a href="http://www.bestukmbt.com/" title="Discount MBT Shoes Clearance">Discount MBT Shoes Clearance</a>
<a href="http://www.mbtusoutlet.com/" title="MBT Shoes US Clearance">MBT Shoes US Clearance</a>
<a href="http://www.mbtukoutlet.com/" title="mbt shoes outlet">mbt shoes outlet</a>
<a href="http://www.mbtdiscountlife.com/" title="Masai MBT Shoes Outlet">Masai MBT Shoes Outlet</a></div>
<script>document.write ('<' + '/d' + 'i' + 'v>');</script>
