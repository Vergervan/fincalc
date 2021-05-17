Attribute VB_Name = "RunModule"
Option Explicit

Public Sub Run()
Attribute Run.VB_ProcData.VB_Invoke_Func = "m\n14"
             Dim loader As ListBoxLoader
             Set loader = New ListBoxLoader
             MainForm.Show
End Sub
