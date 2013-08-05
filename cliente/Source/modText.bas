Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal F As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal X, ByVal Y, ByVal Text As String, color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, X + 1, Y + 0, Text, Len(Text))
    Call TextOut(hDC, X + 0, Y + 1, Text, Len(Text))
    Call TextOut(hDC, X - 1, Y - 0, Text, Len(Text))
    Call TextOut(hDC, X - 0, Y - 1, Text, Len(Text))
    Call SetTextColor(hDC, color)
    Call TextOut(hDC, X, Y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text)
    frmMirage.txtChat.SelColor = QBColor(color)
    frmMirage.txtChat.SelText = vbNewLine & Msg
    frmMirage.txtChat.SelStart = Len(frmMirage.txtChat.Text) - 1

    If frmMirage.chkAutoScroll.value = Unchecked Then
        frmMirage.txtChat.SelStart = frmMirage.txtChat.SelStart
    End If
End Sub

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text & (Msg & vbNewLine)
    Else
        Txt.Text = Txt.Text & Msg
    End If

    Txt.SelStart = Len(Txt.Text)
End Sub


