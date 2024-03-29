﻿<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<div id="Tbody">
    <div style="text-align:center;"><br/>
<%
'=====================================
'  评论处理页面
'    更新时间: 2006-1-12
'=====================================
If Not ChkPost() Then
    response.Write ("非法操作!!")
ElseIf Request.Form("action") = "post" Then
    '评论发表代码
    Dim PostBComm
    PostBComm = postcomm
%>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=PostBComm(0)%></div>
      <div id="MsgBody">
		 <div class="<%=PostBComm(2)%>"></div>
         <div class="MessageText"><%=PostBComm(1)%></div>
	  </div>
	</div>
  </div>
<%
ElseIf Request.QueryString("action") = "del" Then
    Dim DelBComm
    DelBComm = delcomm
%>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=DelBComm(0)%></div>
      <div id="MsgBody">
		 <div class="<%=DelBComm(2)%>"></div>
         <div class="MessageText"><%=DelBComm(1)%></div>
	  </div>
	</div>
  </div>
<%
Else
    response.Write ("非法操作!!")
End If

'============================ 删除评论函数 =================================================

Function delcomm
    Dim post_commID, blog_Comm, blog_CommAuthor, logid
    Dim ReInfo
    ReInfo = Array("错误信息", "", "MessageIcon")
    post_commID = CLng(CheckStr(request.QueryString("commID")))
    Set blog_Comm = Conn.Execute("select top 1 comm_ID,blog_ID,comm_Author from blog_Comment where comm_ID="&post_commID)
    If blog_Comm.EOF Or blog_Comm.bof Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>不存在此评论,或该评论已经被删除!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        delcomm = ReInfo
        Exit Function
    End If
    blog_CommAuthor = blog_Comm("comm_Author")
    If stat_Admin = True Or (stat_CommentDel = True And memName = blog_CommAuthor) Then
        ReInfo(0) = "评论删除成功"
        ReInfo(1) = "<b>评论已经被删除成功!</b><br/><a href=""default.asp?id="&blog_Comm("blog_ID")&""">单击返回</a>"
        ReInfo(2) = "MessageIcon"
        logid = Conn.Execute("select blog_ID from blog_Comment where comm_ID="&post_commID)(0)
        Conn.Execute("update blog_Content set log_CommNums=log_CommNums-1 where log_ID="&blog_Comm("blog_ID"))
        Conn.Execute("update blog_Member set mem_PostComms=mem_PostComms-1 where mem_Name='"&blog_CommAuthor&"'")
        Conn.Execute("DELETE * FROM blog_Comment WHERE comm_ID="&post_commID)
        Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums-1")
        PostArticle logid, False
        getInfo(2)
        NewComment(2)
        delcomm = ReInfo
        Session(CookieName&"_LastDo") = "DelComment"
    Else
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>你没有权限删除评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        delcomm = ReInfo
    End If
    call newEtag
End Function

'====================== 评论发表函数 ===========================================================

Function postcomm
    Dim username, post_logID, post_From, post_FromURL, post_disImg, post_DisSM, post_DisURL, post_DisKEY, post_DisUBB, post_Message, validate
    Dim password
    Dim ReInfo, LastMSG, FlowControl
    ReInfo = Array("错误信息", "", "MessageIcon")
    username = Trim(CheckStr(request.Form("username")))
    password = Trim(CheckStr(request.Form("password")))
    post_logID = CLng(CheckStr(request.Form("logID")))
    validate = Trim(request.Form("validate"))
    post_Message = CheckStr(request.Form("Message"))
    FlowControl = False

    If (memName=empty or blog_validate=true) and (cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) or IsEmpty(Session("GetCode"))) Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">返回重新输入</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Session("GetCode") = Empty
        Exit Function
    End If

    Set LastMSG = conn.Execute("select top 1 comm_Content from blog_Comment order by comm_ID desc")
    If LastMSG.EOF Then
        FlowControl = False
    Else
        If Trim(LastMSG("comm_Content")) = Trim(post_Message) Then FlowControl = True
    End If

    If Left(post_Message,1)= Chr(32) then
            ReInfo(0)="评论发表错误信息"
            ReInfo(1)="<b>评论内容中不允许首字空格</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2)="WarningIcon"
            postcomm = ReInfo
	    Exit Function 
	  End If
  
    If Left(post_Message,1)= Chr(13) then
            ReInfo(0)="评论发表错误信息"
            ReInfo(1)="<b>评论内容中不允许首字换行</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2)="WarningIcon"
            postcomm = ReInfo
	    Exit Function 
	  End If
	  
    If stat_Admin = False Then
        '高级过滤规则
        If regFilterSpam(post_Message, "reg.xml") Then
            ReInfo(0) = "评论发表错误信息"
            ReInfo(1) = "<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2) = "WarningIcon"
            postcomm = ReInfo
            Exit Function
        End If

        '基本过滤规则
        If filterSpam(post_Message, "spam.xml") Then
            ReInfo(0) = "评论发表错误信息"
            ReInfo(1) = "<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
            ReInfo(2) = "WarningIcon"
            postcomm = ReInfo
            Exit Function
        End If
    End If

    If FlowControl Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>禁止恶意灌水！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        postcomm = ReInfo
        Exit Function
    End If

   If DateDiff("s", Request.Cookies(CookieName)("memLastPost"), Now())<blog_commTimerout Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>发言太快,请 "&blog_commTimerout&" 秒后再发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        postcomm = ReInfo
        Exit Function
    End If
    
    If Len(username)<1 Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>请输入你的昵称.</b><br/><a href=""javascript:history.go(-1);"">返回重新输入</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If

    If IsValidUserName(username) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If

    Dim checkMem
    If memName = Empty Then'匿名评论
        If Len(password)>0 Then
            Dim loginUser
            loginUser = login(Request.Form("username"), Request.Form("password"))
            If Not request.Cookies(CookieName)("memName") = username Then
                ReInfo(0) = "评论发表错误信息"
                ReInfo(1) = "<b>登录失败，请检查用户名和密码</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
                ReInfo(2) = "WarningIcon"
                postcomm = ReInfo
                Exit Function
            End If
        Else
            Set checkMem = Conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
            If Not checkMem.EOF Then
                ReInfo(0) = "评论发表错误信息"
                ReInfo(1) = "<b>该用户已经存在，无法发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
                ReInfo(2) = "WarningIcon"
                postcomm = ReInfo
                Exit Function
            End If
        End If
    Else
 			If Not request.Cookies(CookieName)("memName") = username Then
                ReInfo(0) = "评论发表错误信息"
                ReInfo(1) = "<b>请输入正确的用户名</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
                ReInfo(2) = "WarningIcon"
                postcomm = ReInfo
                Exit Function
            End If
    End If
    
    If Not stat_CommentAdd Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>你没有权限发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If
    If Conn.Execute("select log_DisComment from blog_Content where log_ID="&post_logID)(0) Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>该日志不允许发表任何评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        postcomm = ReInfo
        Exit Function
    End If
    post_DisSM = request.Form("log_DisSM")
    post_DisURL = request.Form("log_DisURL")
    post_DisKEY = request.Form("log_DisKey")

    If Len(post_Message)<1 Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>不允许发表空评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If
    If Len(post_Message)>blog_commLength Then
        ReInfo(0) = "评论发表错误信息"
        ReInfo(1) = "<b>评论超过最大字数限制</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        postcomm = ReInfo
        Exit Function
    End If
    'UBB 特别属性
    post_disImg = 1
    post_DisUBB = 0
    If post_DisSM = 1 Then post_DisSM = 1 Else post_DisSM = 0
    If post_DisURL = 1 Then post_DisURL = 0 Else post_DisURL = 1
    If post_DisKEY = 1 Then post_DisKEY = 0 Else post_DisKEY = 1
    '插入数据
    Dim AddComm,re
    
    '去掉回复标签，不允许直接输入
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "\[reply=(.[^\]]*)\](.*?)\[\/reply\]"
    post_Message = re.Replace(post_Message, "$2")

    AddComm = Array(Array("blog_ID", post_logID), Array("comm_Content", post_Message), Array("comm_Author", username), Array("comm_DisSM", post_DisSM), Array("comm_DisUBB", post_DisUBB), Array("comm_DisIMG", post_disImg), Array("comm_AutoURL", post_DisURL), Array("comm_PostIP", getIP), Array("comm_AutoKEY", post_DisKEY))
    DBQuest "blog_Comment", AddComm, "insert"
    'Conn.ExeCute("INSERT INTO blog_Comment(blog_ID,comm_Content,comm_Author,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY) VALUES ("&post_logID&",'"&post_Message&"','"&username&"',"&post_DisSM&","&post_DisUBB&","&post_disImg&","&post_DisURL&",'"&getIP()&"',"&post_DisKEY&")")
    Conn.Execute("update blog_Content set log_CommNums=log_CommNums+1 where log_ID="&post_logID)
    Conn.Execute("update blog_Info set blog_CommNums=blog_CommNums+1")
    Response.Cookies(CookieName)("memLastpost") = Now()
    getInfo(2)
    NewComment(2)
    If memName<>Empty Then
        conn.Execute("update blog_Member set mem_PostComms=mem_PostComms+1 where mem_Name='"&memName&"'")
    End If
    SQLQueryNums = SQLQueryNums + 3
    ReInfo(0) = "评论发表成功"
    ReInfo(1) = "<b>你成功地对该日志发表了评论</b><br/><a href=""default.asp?id="&post_logID&""">单击返回该日志</a>"
    ReInfo(2) = "MessageIcon"
    Session("GetCode") = Empty
    Session(CookieName&"_LastDo") = "AddComment"
    postcomm = ReInfo
    PostArticle post_logID, False
    call newEtag
End Function
%>
  <br/></div> 
 </div>
<!--#include file="footer.asp" -->
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
