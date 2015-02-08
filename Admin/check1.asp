<% 
Session.timeout=60
if isempty(session("aname")) or len(session("aname"))=0 or isempty(session("admin_flag")) or len(session("admin_flag"))=1 or session("admin_flag")<>"into" then  
  founderr=true
  errmsg=errmsg+"<br>"+"<li>你尚未登录，或者超时了！请<a href='login.asp' target='_parent' >重新登录</a>！"
  Call diserror()
  response.End
End if

%>