<%
'=================================================
' moduleSetting Class for PJBlog2
' Author: PuterJam
' UpdateDate: 2005-7-31
'=================================================

Class ModSet
    Private ModSetArray
    Private ModName
    Private state

    Private Sub Class_Initialize()

    End Sub

    Private Sub Class_Terminate()

    End Sub

    '=================================================
    ' 打开模块Open(ModName)
    '=================================================

    Public Function Open(LoadName)
        ModName = LoadName
        If Not IsArray(Application(CookieName&"_Mod_"&ModName))Then
            state = -18902
            ReLoad()
        Else
            ModSetArray = Application(CookieName&"_Mod_"&ModName)
            state = 0
        End If
    End Function

    '=================================================
    ' 从数据库里重新读取模块到缓存ReLoad()
    '=================================================

    Public Function ReLoad()
        If ModName = "" Then
            state = -18901
            Exit Function
        End If
        Dim ModDB, KeyLen, i, GetPlugPath
        i = 0
        KeyLen = conn.Execute("select count(*) from blog_ModSetting where set_ModName='"&ModName&"'")(0)
        Set ModDB = conn.Execute("select * from blog_ModSetting where set_ModName='"&ModName&"'")
        ReDim ModSetArray(KeyLen, 1)
        Do Until ModDB.EOF
            ModSetArray(i, 0) = ModDB("set_KeyName")
            ModSetArray(i, 1) = ModDB("set_KeyValue")
            i = i + 1
            ModDB.movenext
        Loop
        ModSetArray(KeyLen, 0) = "PlugingPath"
        Set GetPlugPath = conn.Execute("select InstallFolder from blog_module where name='"&ModName&"'")
        If GetPlugPath.EOF Then
            state = -18903
            Exit Function
        Else
            ModSetArray(KeyLen, 1) = GetPlugPath(0)
        End If
        Application.Lock
        Application(CookieName&"_Mod_"&ModName) = ModSetArray
        Application.UnLock
        state = 0
    End Function

    '=================================================
    ' 读取字段名称getKeyValue(KeyName)
    '=================================================

    Public Function getKeyValue(KeyName)
        Dim KeysLen, i
        getKeyValue = ""
        KeysLen = UBound(ModSetArray, 1)
        For i = 0 To KeysLen
            If ModSetArray(i, 0) = KeyName Then
                getKeyValue = ModSetArray(i, 1)
                Exit Function
            End If
        Next
    End Function

    '=================================================
    ' 获得出错信息ReLoad()
    '=================================================

    Public Function PasreError
        PasreError = state
        ' -18901 没有打开模块
        ' -18902 缓存里没有任何信息
        ' -18903 没有安装插件
    End Function

    '=================================================
    ' 获得插件所在路径
    '=================================================

    Public Function GetPath
        Dim KeysLen, i
        GetPath = ""
        KeysLen = UBound(ModSetArray, 1)
        GetPath = ModSetArray(KeysLen, 1)
    End Function

    '=================================================
    ' 清除插件占用的 Application 地址
    '=================================================

    Public Function RemoveApplication
        Application.Lock
        Application.Contents.Remove(CookieName&"_Mod_"&ModName)
        Application.UnLock
    End Function

End Class
%>
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
