VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Настройка конфига"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6930
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public loader As ListBoxLoader
Private selectedName As String

Private Sub CommandButton1_Click()
                Dim totinf As TotalInfo
                Set totinf = loader.GetPnlInfo
                
                With totinf
                            If Not .totalCostAddr = "" Then
                                        Range(.totalCostAddr).value = .totalCost
                            End If
                            If Not .totalPnlAddr = "" Then
                                        Range(.totalPnlAddr).value = .totalPnl
                            End If
                            Debug.Print "Total PNL: " & .totalPnl & vbNewLine & _
                                                    "Total portfolio cost: " & .totalCost & vbNewLine & _
                                                    "Pnl addr: " & .totalPnlAddr
                End With
End Sub

Private Sub OwnConfigFrame_Click()
                If Not OwnConfigCheck.Enabled Then DontSelectedCell
End Sub

Private Sub ResetConfigButton_Click()
                   loader.ClearConfig
                   StartLoad
                   MsgBox "Конфиг очищен"
End Sub

Private Sub SaveOwnConfigButton_Click()
                Dim dl As deal
                Set dl = New deal
                With dl
                            .listName = selectedName
                            .deal = GetDealByString(DealTypeBox.Text)
                            If OwnConfigCheck.value = True Then
                                    .dateAddr = DateCellRef.Text
                                    .costAddr = CostCellRef.Text
                                    .amountAddr = AmountCellRef.Text
                                    .lotAddr = LotCellRef.Text
                                    .deal = GetDealByString(DealTypeBox.Text)
                            End If
                            Debug.Print "DealCode: " & .deal
                End With
                loader.MakeOwnConfig dl
                MsgBox "Успешно сохранено"
End Sub

Private Sub StartLoad()
                With loader
                        .LoadBook BookListbox
                        .RefreshList
                End With
End Sub

Private Sub TestButton_Click()
                If selectedName = "" Then
                            MsgBox "Выберите объект для расчёта"
                            Exit Sub
                End If
                Dim infoc1 As InfoContainer
                With loader
                            Set infoc1 = loader.BuildInfo(loader.FindDeal(selectedName, Buy), loader.FindDeal(selectedName, Sell), loader.FindStock(selectedName))
                End With
                loader.CalculateSum infoc1, SelectedSummarize
End Sub

Private Sub UserForm_Initialize() 'Инициализация формы
                Debug.Print "Init"
                Set loader = New ListBoxLoader
                StartLoad
                DealBoxInit
                CellsRefFrame.Enabled = False
                IsSelectedOne
End Sub
Private Sub BookListbox_Click()
                IsSelectedOne
End Sub

Public Function GetDealByString(str As String) As DealType
                Dim i As Integer
                For i = 0 To DealTypeBox.ListCount - 1
                            If str = DealTypeBox.list(i, 0) Then
                                        GetDealByString = i
                                        Exit Function
                            End If
                Next
End Function

Private Sub DealBoxInit()
                With DealTypeBox
                                .AddItem "Покупка"
                                .AddItem "Продажа"
                                .ListIndex = 0
                End With
End Sub

Private Sub HideExceptionsCheck_Change()
                With HideExceptionsCheck
                         If .value = False Then
                                loader.SetMode Highlight
                         Else: loader.SetMode Remove
                         End If
                End With
End Sub

Private Sub IsSelectedOne()
                Dim sel As Boolean, i As Integer
                sel = False
                For i = 0 To BookListbox.ListCount - 1
                            If BookListbox.Selected(i) Then
                                    sel = True
                                    selectedName = BookListbox.list(i, 0)
                            End If
                Next
                OwnConfigCheck.Enabled = sel
                SaveOwnConfigButton.Enabled = sel
End Sub

Private Sub DontSelectedCell()
            MsgBox "Не выбран лист!"
End Sub

Private Sub OwnConfigCheck_Change() 'Включение и выключение индивидуального конфига листа
                    Dim access As Boolean
                    Debug.Print "Update"
                    access = OwnConfigCheck.value
                    If Not OwnConfigCheck.Enabled Then
                                access = False
                    End If
                    CellsRefFrame.Enabled = access
End Sub

Private Sub GlobalConfigButton_Click()
                    Dim gcForm As GlobalConfigForm
                    Set gcForm = New GlobalConfigForm
                    gcForm.Show
End Sub

Private Sub ExceptionListButton_Click()
                    ExceptionsForm.Show
End Sub

