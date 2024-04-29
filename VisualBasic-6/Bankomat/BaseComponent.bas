Attribute VB_Name = "BaseComponent"
Public oCn As New Connection
Public oCm As Command
Public oRs As Recordset

Public Sub Connect()
With oCn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Open App.Path & "\ATM\ATM.mdb"
End With
End Sub

Public Sub Disconnect()
    If oCn.State = adStateOpen Then oCn.Close
End Sub

Public Sub Validate_Numeric(ByRef KeyAscii As Integer)
    Select Case KeyAscii
        Case 999
    End Select
End Sub

Public Sub Validate_Alpha(ByRef KeyAscii As Integer)
    Select Case KeyAscii
        Case ""
    End Select
End Sub

Public Sub Center(ByVal frm As Form)
    frm.Top = (Screen.Height - frm.Height) / 2
    frm.Left = (Screen.Width - frm.Width) / 2
End Sub

Public Function Format_Number(ByVal pNumber As Long, Optional ByVal pPrefix As String = "", Optional ByVal pSuffix As String = "") As String
    If pNumber = 0 Then
        Format_Number = pPrefix & "0000" & pSuffix
    ElseIf pNumber > 0 And pNumber < 10 Then
        Format_Number = pPrefix & "000" & Trim(Str(pNumber)) & pSuffix
    ElseIf pNumber > 9 And pNumber < 100 Then
        Format_Number = pPrefix & "00" & Trim(Str(pNumber)) & pSuffix
    ElseIf pNumber > 99 And pNumber < 999 Then
        Format_Number = pPrefix & "0" & Trim(Str(pNumber)) & pSuffix
    Else
        Format_Number = pPrefix & Trim(Str(pNumber)) & pSuffix
    End If
End Function

Public Function Rip_Number(ByVal pNumber As String) As Long
    Rip_Number = Val(Mid(pNumber, 4))
End Function


