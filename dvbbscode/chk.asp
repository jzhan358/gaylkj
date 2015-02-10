<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>

<body>

<%

	If Request("codestr")="" Then
			Response.write "showerr.asp?action=OtherErr&ErrCodes=<li>请返回输入确认码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			response.end
		Elseif Session("getcode")="9999" then
			Session("getcode")=""
		Elseif Session("getcode")="" then
			Response.write "showerr.asp?action=OtherErr&ErrCodes=<li>请不要重复提交，如需重新登录请返回登录页面。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			response.end
		ElseIf Cstr(Session("getcode"))<>Lcase(Cstr(Trim(Request("codestr")))) Then
			Response.write "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的确认码和系统产生的不一致，请重新输入。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
			response.end
		End If
		Session("getcode")=""
	
response.write "输对了"


%>
</body>
</html>
