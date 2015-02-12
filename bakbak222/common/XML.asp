<%
'=================================================
' XML Class for PJBlog2
' Author: PuterJam
' UpdateDate: 2006-1-19
'=================================================

Class PXML
    Public XmlPath
    Private errorcode
    Public XMLDocument

    Private Sub Class_Initialize()
        errorcode = -1
    End Sub

    Private Sub Class_Terminate()

    End Sub

    '------------------------------------------------
    '函数名字：Open()
    'Open=0，XMLDocument就是一个成功装载XML文档的对象了。
    '------------------------------------------------

    Public Function Open()
        Dim strSourceFile, strError, xmlDom
        xmlDom = getXMLDOM()
        If xmlDom = False Then
            errorcode = -18239123
            Exit Function
        End If

        Set XMLDocument = Server.CreateObject(xmlDom)
        XMLDocument.async = False
        strSourceFile = Server.MapPath(XmlPath)
        XMLDocument.load(strSourceFile)
        errorcode = XMLDocument.parseerror.errorcode
    End Function

    '------------------------------------------------
    '函数名字：OpenXML()
    'Open=0，XMLDocument就是一个成功装载XML文档的对象了。
    '------------------------------------------------

    Public Function OpenXML(xmlStr)
        Dim strSourceFile, strError, xmlDom
        xmlDom = getXMLDOM()
        If xmlDom = False Then
            errorcode = -18239123
            Exit Function
        End If

        Set XMLDocument = Server.CreateObject(getXMLDOM())
        XMLDocument.async = False
        XMLDocument.load(xmlStr)
        errorcode = XMLDocument.parseerror.errorcode
    End Function

    '------------------------------------------------
    '函数名字：getError()
    '------------------------------------------------

    Public Function getError()
        getError = errorcode
    End Function

    '------------------------------------------------
    '函数名字：CloseXml()
    '------------------------------------------------

    Public Function CloseXml()
        If IsObject(XMLDocument) Then
            Set XMLDocument = Nothing
        End If
    End Function


    '------------------------------------------------
    'SelectXmlNodeText(elementname)
    '获得当个 elementname 元素
    '------------------------------------------------

    Public Function SelectXmlNodeText(elementname)
        Dim xmlItems
        selectXmlNodeText = ""
        Set xmlItems = XMLDocument.getElementsByTagName(elementname)
        If xmlItems.Length <> 0 Then selectXmlNodeText = xmlItems.Item(0).text
    End Function

    '------------------------------------------------
    'SelectXmlNode(elementname,itemID)
    '获得当个 elementname 元素
    '------------------------------------------------

    Public Function SelectXmlNode(elementname, itemID)
        Set SelectXmlNode = XMLDocument.getElementsByTagName(elementname).Item(itemID)
    End Function

    '------------------------------------------------
    'GetXmlNodeLength(elementname)
    '获得当个 elementname 元素的Length值
    '------------------------------------------------

    Public Function GetXmlNodeLength(elementname)
        GetXmlNodeLength = XMLDocument.getElementsByTagName(elementname).Length
    End Function

    '------------------------------------------------
    'GetAttributes(elementname,nodeName,ID)
    '获得当个 elementname 元素的attributes值
    '------------------------------------------------

    Public Function GetAttributes(elementname, nodeName, itemID)
        Dim XmlAttributes, i
        Set XmlAttributes = XMLDocument.getElementsByTagName(elementname).Item(itemID).Attributes
        For i = 0 To XmlAttributes.Length -1
            If XmlAttributes(i).Name = nodeName Then
                GetAttributes = XmlAttributes(i).Value
                Exit Function
            End If
        Next
        GetAttributes = 0
    End Function

    '------------------------------------------------
    'SelectXmlNodeItemText(elementname,ID)
    '获得当个某 elementname 元素的Length值
    '------------------------------------------------

    Public Function SelectXmlNodeItemText(elementname, ID)
        Dim xmlItems
        SelectXmlNodeItemText = ""
        Set xmlItems = XMLDocument.getElementsByTagName(elementname)
        If xmlItems.Length <> 0 Then SelectXmlNodeItemText = xmlItems.Item(ID).text
    End Function

    '------------------------------------------------
    'WriteXmlNodeItemText(elementname,ID)
    '写入当个某 elementname 元素的text值
    '------------------------------------------------

    Public Function WriteXmlNodeItemText(elementname, ID, Str)
        WriteXmlNodeItemText = 0
        Dim temp, temp1
        Set temp = XMLDocument.getElementsByTagName(elementname).Item(ID)
        temp.childNodes(0).text = Str
        XMLDocument.save Server.MapPath(XmlPath)
    End Function

    '------------------------------------------------
    'IsXmlNode(elementname)
    '检测是否存在 elementname 元素
    'True代表存在,False代表不存在
    '------------------------------------------------

    Public Function IsXmlNode(elementname)
        Dim Temp
        IsXmlNode = True
        Set Temp = XMLDocument.getElementsByTagName(elementname)

        If Temp.Length = 0 Then
            IsXmlNode = False
        End If
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
