<%
	session.codepage = 65001  
	Response.CharSet = "utf-8"
	dim conn
	dim connstr
	dim db
	dim cityname,cityurl,ecityname,cityviewmode,citymodepath
	dim WeburlExt
	citymodepath="/skin/default"
	cityviewmode=1 '网站浏览模式 0动态 1静态
	if chkNull(dns,1) then dns=""
	db=dns&"data/1.mdb"
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	conn.Open connstr
	
	sub closedatabase()
	  conn.close
      set conn=Nothing
    End sub
	
	set rsconfig=server.CreateObject("adodb.recordset")
	rsconfig.open "select top 1 c_incname,c_web,e_incname from config",conn,1,1
	if not (rsconfig.eof and rsconfig.bof) then	
		cityname=rsconfig("c_incname")
		cityurl=rsconfig("c_web")
		ecityname=rsconfig("e_incname")		
	end if	
	if cityviewmode=1 then weburlExt="html" else weburlExt="asp"
	rsconfig.close
	set rsconfig=nothing	
%>

