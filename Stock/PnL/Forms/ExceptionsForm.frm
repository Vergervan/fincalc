VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExceptionsForm 
   Caption         =   "Настройка исключений"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5145
   OleObjectBlob   =   "ExceptionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExceptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SwitchButton_SpinDown()
                        If ExceptionListBox.ListCount < 1 Then Exit Sub
                        Dim i As Integer
                        For i = 0 To ExceptionListBox.ListCount
                            If ExceptionListBox.Selected(i) Then
                                Debug.Print i
                                AllListsBox.AddItem ExceptionListBox.list(i, 0)
                                ExceptionListBox.RemoveItem i
                                Exit For
                            End If
                        Next
End Sub

Private Sub SwitchButton_SpinUp()
                    If AllListsBox.ListCount < 1 Then Exit Sub
                    Dim i As Integer
                    For i = 0 To AllListsBox.ListCount
                        If AllListsBox.Selected(i) Then
                            ExceptionListBox.AddItem AllListsBox.list(i, 0)
                            AllListsBox.RemoveItem i
                            Exit For
                        End If
                    Next
End Sub

Private Sub UserForm_Initialize()
                            MainForm.loader.LoadBook AllListsBox
                            MainForm.loader.ClearListCollision AllListsBox, ExceptionListBox
End Sub

Private Sub UserForm_Terminate()
            MainForm.loader.RefreshList
End Sub

Private Sub MakeConfig()
            MainForm.loader.MakeExceptionsListConfig ExceptionListBox
End Sub

Private Sub SaveButton_Click()
            MakeConfig
            MsgBox "Успешно сохранено"
End Sub
