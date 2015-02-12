<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>
'====== Upload_5xSoft Class ====================================
'  化境ASP无组件上传类 2.0
'     版权归原作者所有。
'     舜子于 2005-10-20 修改
'============================================================
dim Data_5xsoft

Class Upload_File
  
dim objForm,objFile,Version

Public function Form(strForm)
   strForm=lcase(strForm)
   if not objForm.exists(strForm) then
     Form=""
   else
     Form=objForm(strForm)
   end if
 end function

Public function File(strFile)
   strFile=lcase(strFile)
   if not objFile.exists(strFile) then
     set File=new FileInfo
   else
     set File=objFile(strFile)
   end if
 end function


Private Sub Class_Initialize 
  dim RequestData,sStart,vbCrlf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,theFile
  dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
  dim iFindStart,iFindEnd
  dim iFormStart,iFormEnd,sFormName
  Version="化境HTTP上传程序 Version 2.0"
  set objForm=Server.CreateObject("Scripting.Dictionary")
  set objFile=Server.CreateObject("Scripting.Dictionary")
  if Request.TotalBytes<1 then Exit Sub
  set tStream = Server.CreateObject("adodb.stream")
  set Data_5xsoft = Server.CreateObject("adodb.stream")
  Data_5xsoft.Type = 1
  Data_5xsoft.Mode =3
  Data_5xsoft.Open
  Data_5xsoft.Write  Request.BinaryRead(Request.TotalBytes)
  Data_5xsoft.Position=0
  RequestData =Data_5xsoft.Read 

  iFormStart = 1
  iFormEnd = LenB(RequestData)
  vbCrlf = chrB(13) & chrB(10)
  sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,vbCrlf)-1)
  iStart = LenB (sStart)
  iFormStart=iFormStart+iStart+1
  while (iFormStart + 10) < iFormEnd 
	iInfoEnd = InStrB(iFormStart,RequestData,vbCrlf & vbCrlf)+3
	tStream.Type = 1
	tStream.Mode =3
	tStream.Open
	Data_5xsoft.Position = iFormStart
	Data_5xsoft.CopyTo tStream,iInfoEnd-iFormStart
	tStream.Position = 0
	tStream.Type = 2
	tStream.Charset ="UTF-8"
	sInfo = tStream.ReadText
	tStream.Close
	'取得表单项目名称
	iFormStart = InStrB(iInfoEnd,RequestData,sStart)
	iFindStart = InStr(22,sInfo,"name=""",1)+6
	iFindEnd = InStr(iFindStart,sInfo,"""",1)
	sFormName = lcase(Mid (sinfo,iFindStart,iFindEnd-iFindStart))
	'如果是文件
	if InStr (45,sInfo,"filename=""",1) > 0 then
		set theFile=new FileInfo
		'取得文件名
		iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
		iFindEnd = InStr(iFindStart,sInfo,"""",1)
		sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		theFile.FileName=getFileName(sFileName)
		theFile.FilePath=getFilePath(sFileName)
		theFile.FileExt=getFileExt(sFileName)
		'取得文件类型
		iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
		iFindEnd = InStr(iFindStart,sInfo,vbCr)
		theFile.FileType =Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		theFile.FileStart =iInfoEnd
		theFile.FileSize = iFormStart -iInfoEnd -3
		theFile.FormName=sFormName
		if not objFile.Exists(sFormName) then
		  objFile.add sFormName,theFile
		end if
	else
	'如果是表单项目
		tStream.Type =1
		tStream.Mode =3
		tStream.Open
		Data_5xsoft.Position = iInfoEnd 
		Data_5xsoft.CopyTo tStream,iFormStart-iInfoEnd-3
		tStream.Position = 0
		tStream.Type = 2
		tStream.Charset ="gb2312"
	        sFormValue = tStream.ReadText 
	        tStream.Close
		if objForm.Exists(sFormName) then
		  objForm(sFormName)=objForm(sFormName)&", "&sFormValue		  
		else
		  objForm.Add sFormName,sFormValue
		end if
	end if
	iFormStart=iFormStart+iStart+1
	wend
  RequestData=""
  set tStream =nothing
End Sub

Private Sub Class_Terminate  
 if Request.TotalBytes>0 then
	objForm.RemoveAll
	objFile.RemoveAll
	set objForm=nothing
	set objFile=nothing
	Data_5xsoft.Close
	set Data_5xsoft =nothing
 end if
End Sub
   
 
 Private function getFilePath(FullPath)
  If FullPath <> "" Then
   GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
  Else
   GetFilePath = ""
  End If
 End  function
 
 Private function getFileName(FullPath)
  If FullPath <> "" Then
   GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
  Else
   GetFileName = ""
  End If
 End  function
 
 Private function getFileExt(FullPath)
  If FullPath <> "" Then
		GetFileExt = mid(FullPath,InStrRev(FullPath, ".")+1)
	Else
		GetFileExt = ""
  End If
 End function
 
End Class

Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
  Private Sub Class_Initialize 
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
    FileExt=""
  End Sub
  
 Public function SaveAs(FullPath)
    dim dr,ErrorChar,i
    SaveAs=true
    if trim(fullpath)="" or FileStart=0 or FileName="" or right(fullpath,1)="/" then exit function
    set dr=CreateObject("Adodb.Stream")
    dr.Mode=3
    dr.Type=1
    dr.Open
    Data_5xsoft.position=FileStart
    Data_5xsoft.copyto dr,FileSize
    dr.SaveToFile FullPath,2
    dr.Close
    set dr=nothing 
    SaveAs=false
  end function
  End Class
</SCRIPT>
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
