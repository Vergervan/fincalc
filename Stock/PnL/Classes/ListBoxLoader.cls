VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ClearMode
                Remove = 0
                Highlight = 1
End Enum

Public Enum CalculateType
                TotalSummarize = 0
                SelectedSummarize = 1
End Enum

Private hiddenmode As ClearMode
Private xmlDoc As DOMDocument60
Private path As String
Private exceptList As IXMLDOMNodeList
Private ownConfigList As IXMLDOMNodeList
Public pnls As Dictionary
Private infoCon As InfoContainer
Private totalinf As TotalInfo
Public Property Get Doc() As DOMDocument60
                Set Doc = xmlDoc
End Property
Public Sub SetMode(ByRef newmode As ClearMode)
                hiddenmode = newmode
                RefreshList
End Sub

Public Function GetPnlInfo() As TotalInfo
            'TODO Add List name into InfoContainer and remove it from Stock and Deal
            Dim keystr As Variant, expel As IXMLDOMNode, res As InfoContainer, item As InfoContainer, ch As Object
            Dim regex As RegExp
            Set regex = New RegExp
            totalinf.totalCost = 0
            totalinf.totalPnl = 0
            For Each keystr In pnls.Keys
                            regex.Pattern = keystr
                            For Each ch In exceptList
                                        If regex.Test(ch.Text) = True Then
                                                    Debug.Print "Continue " & ch.Text
                                                    GoTo Continue
                                        End If
                            Next
                            Set item = pnls.item(keystr)
                            Set res = BuildInfo(FindDeal(keystr, Buy), FindDeal(keystr, Sell), FindStock(keystr))
                            If res Is Nothing Then Debug.Print "Returned nothing"
                            CalculateSum res, TotalSummarize
                            'TODO Fix values' bug
                            With totalinf
                                        .totalCost = .totalCost + item.PortfolioCost
                                        .totalPnl = .totalPnl + item.CurrentPnl
                            End With
                            Set GetPnlInfo = totalinf
Continue:
            Next
End Function

Public Function BuildInfo(buyDeal As deal, sellDeal As deal, stock As stock) As InfoContainer
            If buyDeal Is Nothing Or sellDeal Is Nothing Or stock Is Nothing Then
                        Set BuildInfo = Nothing
                        Exit Function
            End If
            Set BuildInfo = New InfoContainer
            With BuildInfo
                        Set .dealBuy = buyDeal
                        Set .dealSell = sellDeal
                        Set .stockObj = stock
            End With
End Function

Private Sub Class_Initialize()
            Set totalinf = New TotalInfo
            hiddenmode = Highlight
            FillPnls
            path = "stocksconfig.xml"
            LoadXML
            
            FillTotalInfo
End Sub

Private Sub FillPnls()
            Set pnls = New Dictionary
            Dim ch As Worksheet, newInf As InfoContainer
            For Each ch In ActiveWorkbook.Worksheets
                        Set newInf = New InfoContainer
                        pnls.Add ch.name, newInf
            Next
End Sub

Private Sub FillTotalInfo()
            Dim costNode As Object, totpnlNode As Object
            Set costNode = xmlDoc.getElementsByTagName("TotalCostCell")
            If costNode.Length > 0 Then
                        totalinf.totalCostAddr = costNode(0).Text
            End If
            Set totpnlNode = xmlDoc.getElementsByTagName("TotalPnlCell")
            If totpnlNode.Length > 0 Then
                        totalinf.totalPnlAddr = totpnlNode(0).Text
            End If
End Sub

Private Sub LoadXML()
            Set xmlDoc = New DOMDocument60
            xmlDoc.async = False
            
            If Not xmlDoc.Load(path) Then
                        InitializeConfig
                        Exit Sub
            End If
            RefreshAll
End Sub

Public Sub CalculateSum(ByRef infocon1 As InfoContainer, sumtype As CalculateType)
            Dim sum1 As Variant, sum2 As Variant, lf As Range, curPnl As Variant, buySum As Variant, curCost As Range, buyAmount As Variant
            '        Sell                           Buy                         Lot
            With infocon1
                    If infocon1 Is Nothing Then
                                    MsgBox "Ошибка нахождения сделки или информации по акции" & vbNewLine & _
                                                    "Возможно не установен конфиг для нахождения необходимых ячеек"
                                    Exit Sub
                    End If
                    Dim exEl As IXMLDOMNode, exName As String, regex As RegExp
                    Set regex = New RegExp
                    If Not exceptList Is Nothing Then
                            For Each exEl In exceptList
                                            regex.Pattern = exEl.Text
                                            If regex.Test(.dealBuy.listName) Or regex.Test(.dealSell.listName) Then
                                                        If sumtype = SelectedSummarize Then
                                                                    MsgBox "Выбранный лист находится в исключениях"
                                                        Else: Debug.Print .dealBuy.listName & " находится в исключениях"
                                                        End If
                                                        
                                                        Exit Sub
                                            End If
                            Next
                    End If
                    Dim c1 As Range, a1 As Range
                    Dim c2 As Range, a2 As Range
                    
                    Set lf = Range(.dealBuy.lotAddr)
                    Set curCost = Range(.stockObj.currentCostAddr)
                        
                    Set c1 = Range(.dealBuy.costAddr)
                    Set a1 = Range(.dealBuy.amountAddr)
                    Set c2 = Range(.dealSell.costAddr)
                    Set a2 = Range(.dealSell.amountAddr)
                    
                    Dim sellAmount As Variant
                    Do While Not c2.Text = ""
                            sum2 = sum2 + c2.value * a2.value
                            sellAmount = sellAmount + a2.value
                            Set c2 = c2.Offset(1)
                            Set a2 = a2.Offset(1)
                    Loop
                    
                    Dim diffA As Variant
                    diffA = sellAmount
                    Do While Not c1.Text = ""
                            If Not diffA <= 0 Then
                                If diffA > a1.value Then
                                            sum1 = sum1 + c1.value * a1.value
                                Else:    sum1 = sum1 + c1.value * diffA
                                End If
                                diffA = diffA - a1.value
                            End If
                                buySum = buySum + c1.value * a1.value
                                buyAmount = buyAmount + a1.value
                                Set c1 = c1.Offset(1)
                                Set a1 = a1.Offset(1)
                    Loop
            End With
            Dim inf As InfoContainer
            Set inf = pnls.item(infocon1.dealBuy.listName)
            With inf
                        .SetCurrentPnl Round((curCost * buyAmount) - buySum, 1)
                        .SetSellsPnl Round((sum2 - sum1) * lf.value, 1)
                        .SetPortfolioCost (curCost * buyAmount)
            End With
            Dim stpnl As Range
            Set stpnl = Range(infocon1.stockObj.stockPnlAddr)
            stpnl.value = inf.CurrentPnl
End Sub

Public Function FindStock(name As Variant) As stock
            Set FindStock = New stock
            FindStock.listName = name
            Dim stockNode As Object, config As Object, ch As IXMLDOMNode
            
            Set config = xmlDoc.getElementsByTagName("UniqueConfigs")
            '/// TODO UNIQUE CONFIGS PARSE
            If config.Length > 0 Then
                       Set config = config.item(0)
                       Set stockNode = config.SelectNodes("Stock")
                       If stockNode.Length > 0 Then Set stockNode = stockNode.item(0)
            End If
            '///
            
            Set config = xmlDoc.getElementsByTagName("GlobalConfig")
            If config.Length > 0 Then
                    Set config = config.item(0)
                    Set stockNode = config.SelectNodes("Stock")
                    If stockNode.Length > 0 Then
                                    FillStock FindStock, stockNode.item(0)
                                    Exit Function
                    End If
            End If
            
            Set FindStock = Nothing
End Function

Public Sub FillStock(ByRef stockObj As stock, stockNode As Object)
            Dim nodeHandle As Object
            Set nodeHandle = stockNode.SelectNodes("CurrentCostCell")
            If nodeHandle.Length > 0 Then
                        If stockObj.currentCostAddr = "" Then stockObj.currentCostAddr = "'" & stockObj.listName & "'" & "!" & nodeHandle.item(0).Text
            End If
            Set nodeHandle = stockNode.SelectNodes("StockPnlCell")
            If nodeHandle.Length > 0 Then
                        If stockObj.stockPnlAddr = "" Then stockObj.stockPnlAddr = "'" & stockObj.listName & "'" & "!" & nodeHandle.item(0).Text
            End If
End Sub

Public Function FindDeal(name As Variant, dealt As DealType) As deal
            Set FindDeal = New deal
            FindDeal.listName = name
            Dim dealName As String, config As Object, dealNode As Object, ch As IXMLDOMNode
            If dealt = Buy Then
                     dealName = "Buy"
            Else: dealName = "Sell"
            End If
            Set config = xmlDoc.getElementsByTagName("UniqueConfigs")
            If config.Length > 0 Then
                    Set config = config.item(0)
                    Set dealNode = config.SelectNodes(dealName)
                    If dealNode.Length > 0 Then
                                Set dealNode = dealNode.item(0)
                    End If
                    Set dealNode = dealNode.SelectNodes("Item")
                    If dealNode.Length > 0 Then
                        For Each ch In dealNode
                                        Dim nameAttr As Object
                                        If ch.Attributes.getNamedItem("Name").Text = name Then
                                                    Debug.Print "Found " & name
                                                    
                                                    FillDeal FindDeal, ch
                                                    Exit For
                                        End If
                    Next
                    End If
            End If
            Set config = xmlDoc.getElementsByTagName("GlobalConfig")
            If config.Length > 0 Then
                    Set config = config.item(0)
                    Set dealNode = config.SelectNodes(dealName)
                    If dealNode.Length > 0 Then
                                    FillDeal FindDeal, dealNode.item(0)
                                    
                                    Exit Function
                    End If
            End If
            
            
            Set FindDeal = Nothing
End Function

Private Sub FillDeal(ByRef dealObj As deal, cellsNode As Object)
            Dim nodeHandle As Object
            Set nodeHandle = cellsNode.SelectNodes("DateCell")
            If nodeHandle.Length > 0 Then
                        If dealObj.dateAddr = "" Then dealObj.dateAddr = "'" & dealObj.listName & "'" & "!" & nodeHandle.item(0).Text
            End If
            Set nodeHandle = cellsNode.SelectNodes("CostCell")
            If nodeHandle.Length > 0 Then
                        If dealObj.costAddr = "" Then dealObj.costAddr = "'" & dealObj.listName & "'" & "!" & nodeHandle.item(0).Text
            End If
            Set nodeHandle = cellsNode.SelectNodes("AmountCell")
            If nodeHandle.Length > 0 Then
                        If dealObj.amountAddr = "" Then dealObj.amountAddr = "'" & dealObj.listName & "'" & "!" & nodeHandle.item(0).Text
            End If
            Set nodeHandle = cellsNode.SelectNodes("LotCell")
            If nodeHandle.Length > 0 Then
                        If dealObj.lotAddr = "" Then dealObj.lotAddr = "'" & dealObj.listName & "'" & "!" & nodeHandle.item(0).Text
            End If
End Sub

Public Sub ClearConfig()
            Dim ch As IXMLDOMNode
            For Each ch In xmlDoc.ChildNodes
                            xmlDoc.RemoveChild ch
            Next
            xmlDoc.Save path
            InitializeConfig
            RefreshAll
End Sub

Private Sub InitializeConfig()
            Dim settingsEl As IXMLDOMNode
            Set settingsEl = xmlDoc.createNode(NODE_ELEMENT, "Settings", "")
            xmlDoc.appendChild settingsEl
            xmlDoc.Save path
End Sub

Public Sub LoadBook(ByRef curListBox As Object) 'Загрузчик листов из текущей книги
            curListBox.Clear
            Dim list As Worksheet
            For Each list In Application.Worksheets
                    curListBox.AddItem list.name 'Добавление наименования листа в ListBox
            Next
End Sub

Public Sub MakeOwnConfig(dealObj As deal)
            If dealObj.listName = "" Then Exit Sub
            Dim ownConfigsEl As Object, settingsEl As IXMLDOMNode
            Set settingsEl = xmlDoc.getElementsByTagName("Settings")(0)
            Set ownConfigsEl = xmlDoc.getElementsByTagName("UniqueConfigs")
            If ownConfigsEl.Length > 0 Then
                    Set ownConfigsEl = ownConfigsEl.item(0)
            Else:
                    Set ownConfigsEl = xmlDoc.createNode(NODE_ELEMENT, "UniqueConfigs", "")
                    settingsEl.appendChild ownConfigsEl
            End If
            Dim item As IXMLDOMNode
            Dim dealNode As Object, itemEl As IXMLDOMNode
            Dim dealName As String
            If dealObj.deal = Buy Then
                     dealName = "Buy"
            Else: dealName = "Sell"
            End If
            Set dealNode = ownConfigsEl.getElementsByTagName(dealName)
            If dealNode.Length > 0 Then
                    Set dealNode = dealNode.item(0)
            Else:
                    Set dealNode = xmlDoc.createNode(NODE_ELEMENT, dealName, "")
                    ownConfigsEl.appendChild dealNode
            End If
            
            Dim exItem As IXMLDOMNode
            For Each exItem In dealNode.ChildNodes
                        If exItem.Attributes.getNamedItem("Name").Text = dealObj.listName Then
                                        dealNode.RemoveChild exItem
                        End If
            Next
            Set item = dealObj.BuildDealNode(Me)
            If item Is Nothing Then
                     Debug.Print "Item is nothing"
            Else: dealNode.appendChild item
            End If
            xmlDoc.Save path
End Sub

Public Sub MakeGlobalConfig(dealObj As deal, stockObj As stock, totinf As TotalInfo)
            Dim globEl As Object, settingsEl As IXMLDOMNode
            Set settingsEl = xmlDoc.getElementsByTagName("Settings")(0)
            Set globEl = settingsEl.SelectNodes("GlobalConfig")
            If globEl.Length > 0 Then
                    Set globEl = globEl.item(0)
            Else:
                    Set globEl = xmlDoc.createNode(NODE_ELEMENT, "GlobalConfig", "")
                    settingsEl.appendChild globEl
            End If
            
            Dim totcNode As Object, totpnlNode As Object
            Set totcNode = globEl.SelectNodes("TotalCostCell")
            If totcNode.Length > 0 Then
                        Set totcNode = totcNode.item(0)
            Else:
                        Set totcNode = xmlDoc.createNode(NODE_ELEMENT, "TotalCostCell", "")
                        globEl.appendChild totcNode
            End If
            If Not totinf.totalCostAddr = "" Then
                        totcNode.Text = totinf.totalCostAddr
            End If
            Set totpnlNode = globEl.SelectNodes("TotalPnlCell")
            If totpnlNode.Length > 0 Then
                        Set totpnlNode = totpnlNode.item(0)
            Else:
                        Set totpnlNode = xmlDoc.createNode(NODE_ELEMENT, "TotalPnlCell", "")
                        globEl.appendChild totpnlNode
            End If
            If Not totinf.totalPnlAddr = "" Then
                        totpnlNode.Text = totinf.totalPnlAddr
            End If
            
            Dim item As IXMLDOMNode, stockNode As Object, stockItem As IXMLDOMNode
            Dim dealNode As Object, itemEl As IXMLDOMNode
            Dim dealName As String
            If dealObj.deal = Buy Then
                     dealName = "Buy"
            Else: dealName = "Sell"
            End If
            Set stockNode = globEl.SelectNodes("Stock")
            If stockNode.Length > 0 Then
                    Set stockNode = stockNode.item(0)
            Else:
                    Set stockNode = xmlDoc.createNode(NODE_ELEMENT, "Stock", "")
                    globEl.appendChild stockNode
            End If
            
            Set dealNode = globEl.SelectNodes(dealName)
            If dealNode.Length > 0 Then
                    Set dealNode = dealNode.item(0)
            Else:
                    Set dealNode = xmlDoc.createNode(NODE_ELEMENT, dealName, "")
                    globEl.appendChild dealNode
            End If
           
           Dim ch As IXMLDOMNode
           
            Set stockItem = stockObj.BuildStockNode(Me)
            If Not stockItem Is Nothing Then
                        For Each ch In stockNode.ChildNodes
                                    stockNode.RemoveChild ch
                        Next
                        
                        For Each ch In stockItem.ChildNodes
                                    stockNode.appendChild ch
                        Next
            End If
            
            Set item = dealObj.BuildDealNode(Me)
            If Not item Is Nothing Then
                        For Each ch In dealNode.ChildNodes
                                    dealNode.RemoveChild ch
                        Next
                        
                        For Each ch In item.ChildNodes
                                    dealNode.appendChild ch
                        Next
            End If
            xmlDoc.Save path
End Sub

Public Sub ClearListCollision(ByRef list1 As Object, ByRef list2 As Object)
            If exceptList Is Nothing Then Exit Sub
            
            Dim it1 As Object, i As Integer
            For Each it1 In exceptList
                    For i = 0 To list1.ListCount - 1
                            If list1.list(i, 0) = it1.Text Then
                                        list2.AddItem (list1.list(i, 0))
                                        list1.RemoveItem (i)
                                        Exit For
                            End If
                    Next
            Next
End Sub
Public Sub ClearListCollision2(ByRef list1 As Object)
            If exceptList Is Nothing Then Exit Sub
            
            Dim it1 As Object, i As Integer
            For Each it1 In exceptList
                    For i = 0 To list1.ListCount - 1
                            If list1.list(i, 0) = it1.Text Then
                                        If hiddenmode = Remove Then
                                                list1.RemoveItem (i)
                                        Else: list1.list(i, 0) = "--" & list1.list(i, 0)
                                        End If
                                        Exit For
                            End If
                    Next
            Next
End Sub

Public Sub RefreshAll()
            RefreshOwnConfig
            RefreshExceptions
End Sub

Public Sub RefreshOwnConfig()
            GetOwnConfig
End Sub

Private Sub GetOwnConfig()
            Dim ucNode As Object
            Set ucNode = xmlDoc.getElementsByTagName("UniqueConfigs")
            If ucNode.Length > 0 Then
                            Set ucNode = ucNode.item(0)
                            Set ownConfigList = ucNode.SelectNodes("Item")
                            Exit Sub
            End If
            Set ownConfigList = Nothing
End Sub

Public Sub RefreshExceptions()
            GetExceptionsFromConfig
End Sub

Private Sub GetExceptionsFromConfig()
            Dim expNode As Object
            Set expNode = xmlDoc.getElementsByTagName("ExceptionList")
            If expNode.Length > 0 Then
                            Set expNode = expNode.item(0)
                            Set exceptList = expNode.getElementsByTagName("Item")
                            Exit Sub
            End If
            Set exceptList = Nothing
End Sub

Public Sub MakeExceptionsListConfig(ByRef curListBox As Object)
            Dim settingsEl As Object
            Set settingsEl = xmlDoc.getElementsByTagName("Settings")
            Set settingsEl = settingsEl.item(0)
            Dim expList As Object
            Set expList = xmlDoc.getElementsByTagName("ExceptionList")
            If expList.Length > 0 Then
                    Set expList = expList.item(0)
            Else:
                    Set expList = xmlDoc.createNode(NODE_ELEMENT, "ExceptionList", "")
                    settingsEl.appendChild expList
            End If
            Dim ch As IXMLDOMNode
            If expList.ChildNodes.Length > 0 Then
                            For Each ch In expList.ChildNodes
                                            expList.RemoveChild ch
            Next
            End If
            
            Dim i As Integer
            For i = 0 To curListBox.ListCount - 1
                        Dim expel As IXMLDOMNode
                        Set expel = xmlDoc.createNode(NODE_ELEMENT, "Item", "")
                        expel.Text = curListBox.list(i, 0)
                        expList.appendChild expel
            Next
            xmlDoc.Save path
            
            RefreshExceptions
End Sub
Public Sub RefreshList()
                LoadBook MainForm.BookListbox
                ClearListCollision2 MainForm.BookListbox
End Sub
