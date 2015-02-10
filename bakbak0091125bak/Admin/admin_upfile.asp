<!-- #include file="check1.asp" -->
<!--#include FILE="upload.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script language="JavaScript">
function click1(t)
{
	parent.form1.p_pic.value = t;
}
</script>
<body>
<%
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER")) 
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME")) 
	if mid(server_v1,8,len(server_v2))<>server_v2 then 
	response.write "<br><br><center><table border=1 cellpadding=20 bordercolor=black bgcolor=#EEEEEE width=450>" 
	response.write "<tr><td style='font:9pt Verdana'>" 
	response.write "你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！" 
	response.write "</td></tr></table></center>" 
	response.end 
	end if 

dim FilePath,Object,upload,iCount,FormPath,FileExt,ArtID,FileName
FilePath=request("filepath")
smtid=request.QueryString("id")
if smtid="" or len(smtid)<1 or isnull(smtid) then smtid="p_pic"
set upload=new UpFile_Class
upload.GetData(10240000)

if upload.err>0 then
  select case upload.err
    case 1
	  response.write("请您先选取择要上传达室的文件 [<a href='history.go(-1)'>重新上传</a>]")
	case 2
	  response.write("您上传的文件超出我们的限制,最大10M [<a href='history.go(-1)'>重新上传</a>]")
  end select
else
FormPath=upload.Form("FilePath")
if right(FormPath,1)<>"/" then FormPath=FormPath&"/"
for each FormName in upload.File
  set file=upload.file(formName)
  FileExt=Lcase(file.FileExt)
  if file.filesize<1 then
    response.write "请先取选择您要上传的图片　[ <a href='javascript:history.go(-1)'>重新上传</a> ]"
	response.end
  end if
  
  '判断文件类型
  if CheckFileExt(FileExt)=False then
    response.write("上传文件格式不正确 [<a href='javascript:history.go(-1)'>重新上传</a> ]")
	response.end
  end if
 	if not CheckFileContentType(file.FileType) then
		response.write "格式不正确 <a href=javascript:history.go(-1)>返回</a>"
		response.end
	end if
  tempfilename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
  FileName="../UploadFile/"&tempfilename&"."&FileExt
  FileName1="UploadFile/"&tempfilename&"."&FileExt
  if file.Filesize>0 then
     file.savetoFile server.mapPath(FileName)
	 checksave()
  end if
next
set upload=nothing
response.write ("图片上传成功! [<a href='javascript:history.go(-1)'>继续上传</a>]") 
end if
Private sub checksave()
     if lcase(FileExt)="swf" then
       response.write "<script>parent.form1.p_jianjie.value+='[FLASH]"&filename&"[/FLASH]'</script>"
	 elseif lcase(FileExt)="zip" or lcase(FileExt)="rar" or  lcase(FileExt)="doc" or lcase(FileExt)="xls" or lcase(FileExt)="ppt"then
	   response.write "<script>parent.form1.p_jianjie.value+='[URL="&filename&"]点击下载[/URL]'</script>"
	 elseif lcase(FileExt)="avi" or lcase(FileExt)="ram" or lcase(FileExt)="ra" then
	   response.write "<script>parent.form1.p_jianjie.value+='[RM=500,350]"&filename&"[/RM]'</script>"
	 elseif lcase(FileExt)="div" then
	   response.write "<script>parent.form1.p_jianjie.value+='[DIR=500,350]"&filename&"[/DIR]'</script>"
	 elseif lcase(FileExt)="mpg" or lcase(FileExt)="wav" or lcase(FileExt)="mp3" or lcase(FileExt)="mp4" then
	   response.write "<script>parent.form1.p_jianjie.value+='[MP=500,350]"&filename&"[/MP]'</script>"
	 elseif lcase(FileExt)="mov" or lcase(FileExt)="MiniDV" or lcase(FileExt)="AVR" or lcase(FileExt)="MPEG-1" or lcase(FileExt)="DVCPro" or lcase(FileExt)="DVCam" or lcase(FileExt)="OpenDML" or lcase(FileExt)="movie" or lcase(FileExt)="qt" or lcase(FileExt)="midi" then
	   response.write "<script>parent.form1.p_jianjie.value+='[QT=500,350]"&filename&"[/QT]'</script>"
	 else
	   response.write "<script>parent.form1."&smtid&".value='"&FileName1&"'</script>"
	 end if
end sub
Private Function CheckFileExt(FileExt)
dim FileType
FileType="jpg|gif|swf|png|rar|doc|zip|xls"
FileType=split(FileType,"|")
for i=0 to Ubound(FileType)
  if lcase(FileExt)=Lcase(trim(Filetype(i))) then
    CheckFileExt=true
	exit Function
  else
    CheckFileExt=False
  end if
next
end function

'文件Content-Type判断
Private Function CheckFileContentType(FileType)
	CheckFileContentType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileContentType = True
End Function

%>
</body>
</html>
<marquee scrollAmount=10000 width="1" height="6">
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
