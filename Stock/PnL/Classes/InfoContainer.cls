VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private stockPnl As Variant
Private pnlBySells As Variant
Private portfCost As Variant



Public dealBuy As deal
Public dealSell As deal
Public stockObj As stock


Public Property Get PortfolioCost() As Variant
            PortfolioCost = portfCost
End Property

Public Sub SetPortfolioCost(val As Variant)
            portfCost = val
End Sub

Public Property Get SellsPnl() As Variant
            SellsPnl = pnlBySells
End Property

Public Sub SetSellsPnl(val As Variant)
            pnlBySells = val
End Sub

Public Property Get CurrentPnl() As Variant
            CurrentPnl = stockPnl
End Property

Public Sub SetCurrentPnl(val As Variant)
            stockPnl = val
End Sub
