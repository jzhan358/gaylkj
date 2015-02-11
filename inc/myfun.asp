<%

'获取模版文件
function getcache(spath)
dim tpath
tpath=citymodepath&spath
getcache=ReadFromUTF(tpath,"utf-8")
getcache=setCommon(getcache)
end function

'设置公共变量
function setCommon(sbody)
sbody=replace(sbody,"$cityname$",cityname)
sbody=replace(sbody,"$cityurl$",cityurl)
sbody=replace(sbody,"$ecityname$",ecityname)
sbody=replace(sbody,"$weburlext$",weburlExt)
setCommon=replace(sbody,"$skinpath$",citymodepath)
end function

'设置页面的关键字
function setKW(sbody,sk,sd,st)
sbody=replace(sbody,"$pagekeywords$",sk)
sbody=replace(sbody,"$pagedescription$",sd)
sbody=replace(sbody,"$pagetitle$",st)
setKW=sbody
end function

'获取左边 slanguage语言 0中文 1英文 shome是否首页 0不是 1是
function getleft(slanguage,shome)
dim snum,sfilename
snum=50
sfilename="/left.html"
if shome=1 then snum=10
if slanguage=1 then sfilename="/eleft.html"
getleft=getcache(sfilename)
getleft=replace(getleft,"$getclass$",getclass(snum,slanguage))
end function

'获取脚部 slanguage语言 0中文 1英文
function getfoot(slanguage)
dim sfilename
sfilename="/foot.html"
if slanguage=1 then sfilename="/efoot.html"
getfoot=getcache(sfilename)
end function

'保存为静态文件
function savetohtml(byval sbody,sfile)
savetohtml=SaveToFileauto(sbody,sfile,"utf-8")
end function

function createSeoPage()
dim home
home=getIndexpage()
call savetohtml(home,"/google.html")
call savetohtml(home,"/baidu.html")
call savetohtml(home,"/yahoo.html")
call savetohtml(home,"/soso.html")
call savetohtml(home,"/163.html")
home=""
end function
%>