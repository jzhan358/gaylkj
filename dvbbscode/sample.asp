<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>验证码调用范例文件</title>
</head>
<body>
<script>

/*显示认证码 o start1*/
function get_Code() {
	var Dv_CodeFile = "/dvbbscode/Dv_GetCode.asp";
	if(document.getElementById("imgid"))
		document.getElementById("imgid").innerHTML = '<img src="'+Dv_CodeFile+'?t='+Math.random()+'" alt="点击刷新验证码" style="cursor:pointer;border:0;vertical-align:middle;" onclick="this.src=\''+Dv_CodeFile+'?t=\'+Math.random()" />'
}
/*o end*/
</script>
<form action="chk.asp" method="post">
附加码
<input type="text" name="codestr"  onfocus="get_Code();return true;" />
&nbsp;
<div id="imgid"></div>
<input  class="button" type="submit" name="submit" value="提交"/>
</form>
</body>
</html>
