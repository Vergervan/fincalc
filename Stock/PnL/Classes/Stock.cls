VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public currentCost As Double
Public listName As String

Public currentCostAddr As String
Public stockPnlAddr As String

Public Function BuildStockNode(loader As ListBoxLoader) As IXMLDOMNode
                    If currentCostAddr = "" And stockPnlAddr = "" Then
                            Set BuildStockNode = Nothing
                            Exit Function
                    End If
                    Dim stockNode As IXMLDOMNode, nameAttr As IXMLDOMNode
                    Set stockNode = loader.Doc.createNode(NODE_ELEMENT, "Stock", "")
                    If Not listName = "" Then
                        Set nameAttr = loader.Doc.createNode(NODE_ATTRIBUTE, "Name", "")
                        nameAttr.Text = listName
                        stockNode.Attributes.setNamedItem nameAttr
                    End If
                    
                    Set BuildStockNode = stockNode
                    
                    Dim newNode As IXMLDOMNode
                
                    With loader
                                If Not currentCostAddr = "" Then
                                        Set newNode = loader.Doc.createNode(NODE_ELEMENT, "CurrentCostCell", "")
                                        newNode.Text = Range(currentCostAddr).AddressLocal
                                        stockNode.appendChild newNode
                                End If
                                
                                If Not stockPnlAddr = "" Then
                                        Set newNode = loader.Doc.createNode(NODE_ELEMENT, "StockPnlCell", "")
                                        newNode.Text = Range(stockPnlAddr).AddressLocal
                                        stockNode.appendChild newNode
                                End If
                    End With
                
End Function
