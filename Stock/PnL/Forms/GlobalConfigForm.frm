VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GlobalConfigForm 
   Caption         =   "Настройка общего конфига"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4200
   OleObjectBlob   =   "GlobalConfigForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GlobalConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CurrentCostCellRef_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub LotCellRef_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
                CellRefEscape KeyCode
End Sub

Private Sub AmountCellRef_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
                CellRefEscape KeyCode
End Sub

Private Sub CostCellRef_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
                CellRefEscape KeyCode
End Sub

Private Sub DateCellRef_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
                CellRefEscape KeyCode
End Sub

Private Sub CellRefEscape(key As Integer)
                If key = 27 Then Me.Hide
End Sub

Private Sub SaveGlobalConfigButton_Click()
                Dim dl As deal, stk As stock, totinf As TotalInfo
                Set dl = New deal
                Set stk = New stock
                Set totinf = New TotalInfo
                With dl
                            .listName = ""
                            .dateAddr = DateCellRef.Text
                            .costAddr = CostCellRef.Text
                            .amountAddr = AmountCellRef.Text
                            .lotAddr = LotCellRef.Text
                            .deal = MainForm.GetDealByString(DealTypeBox.Text)
                End With
                With stk
                            .listName = ""
                            .currentCostAddr = CurrentCostCellRef.Text
                            .stockPnlAddr = StockPNLRef.Text
                End With
                totinf.totalCostAddr = TotalCostRef.Text
                totinf.totalPnlAddr = TotalPNLRef.Text
                MainForm.loader.MakeGlobalConfig dl, stk, totinf
                MsgBox "Успешно сохранено"
End Sub

Private Sub UserForm_Initialize()
                      DealBoxInit
End Sub

Private Sub DealBoxInit()
                With DealTypeBox
                                .AddItem "Покупка"
                                .AddItem "Продажа"
                                .ListIndex = 0
                End With
End Sub
