VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Deal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum DealType
            Buy = 0
            Sell = 1
End Enum

Public listName As String
Public dateAddr As String
Public costAddr As String
Public amountAddr As String
Public lotAddr As String

Public deal As DealType

Public Function BuildDealNode(loader As ListBoxLoader) As IXMLDOMNode
                If costAddr = "" And amountAddr = "" And lotAddr = "" Then
                            Set BuildDealNode = Nothing
                            Exit Function
                End If
                Dim itemEl As IXMLDOMNode
                Set itemEl = loader.Doc.createNode(NODE_ELEMENT, "Item", "")
                Dim nameAttr As IXMLDOMNode
                If Not listName = "" Then
                        Set nameAttr = loader.Doc.createNode(NODE_ATTRIBUTE, "Name", "")
                        nameAttr.Text = listName
                        itemEl.Attributes.setNamedItem nameAttr
                End If
                
                Set BuildDealNode = itemEl
                Dim newNode As IXMLDOMNode
                
                With loader
                            If Not dateAddr = "" Then
                                    Set newNode = loader.Doc.createNode(NODE_ELEMENT, "DateCell", "")
                                    newNode.Text = Range(dateAddr).AddressLocal
                                    itemEl.appendChild newNode
                            End If
                            
                            If Not costAddr = "" Then
                                    Set newNode = loader.Doc.createNode(NODE_ELEMENT, "CostCell", "")
                                    newNode.Text = Range(costAddr).AddressLocal
                                    itemEl.appendChild newNode
                            End If
                            
                            If Not amountAddr = "" Then
                                    Set newNode = loader.Doc.createNode(NODE_ELEMENT, "AmountCell", "")
                                    newNode.Text = Range(amountAddr).AddressLocal
                                    itemEl.appendChild newNode
                            End If
                            
                            If Not lotAddr = "" Then
                                    Set newNode = loader.Doc.createNode(NODE_ELEMENT, "LotCell", "")
                                    newNode.Text = Range(lotAddr).AddressLocal
                                    itemEl.appendChild newNode
                            End If
                End With
                
End Function

Private Sub Class_Initialize()
                deal = Buy
End Sub
