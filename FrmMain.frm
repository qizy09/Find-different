VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "大家来找碴 Alpha"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   903
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrClickCheck 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   300
      Left            =   600
      SmallChange     =   200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6480
      Width           =   13575
   End
   Begin VB.PictureBox PicKid 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   4560
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
      Begin VB.Shape ShpCircle 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   735
         Index           =   0
         Left            =   240
         Shape           =   2  'Oval
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox PicMain 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MaxCircle As Long = 50
Const MaxPicNum As Long = 6
Public IsClickEnabled As Boolean
Public lngDiffFound As Long
Public lngCurrentPicNum As Long
Public IsExit As Boolean
Public IsClicking As Boolean

Private Function LoadPic(ByVal PicIndex As Long) As Long ' return value = how many differences are there in the pics
    If PicIndex > MaxPicNum Then
        ' Succeed
        MsgBox "祝贺你们成功跨过本关！", vbInformation
        On Error Resume Next
        Kill App.Path & "\g_diff.dat"
        On Error GoTo 0
        Open App.Path & "\g_diff.dat" For Append As #1
        Print #1, "finished"
        Close #1
        tmrClickCheck.Enabled = False
        SetStat "成功跨过本关。软件即将关闭……"
        DoEvents
        Sleep 3000
        LoadPic = -1
        IsExit = True
        Unload Me
        End
    End If
    On Error GoTo ErrHandler
    PicMain.Picture = LoadPicture(App.Path & "\Data\" & PicIndex & "_ori.jpg")
    PicKid.Picture = LoadPicture(App.Path & "\Data\" & PicIndex & "_new.jpg")
    ' Rearrange
    PicKid.Left = PicMain.Left + PicMain.Width + 10
    If PicMain.Width + PicKid.Width + 10 > FrmMain.Width Then
        HScroll.Max = PicMain.Width + PicKid.Width + 10 - FrmMain.Width
        HScroll.Enabled = True
    Else
        HScroll.Enabled = False
    End If
    ' Loading the info file
    Open App.Path & "\Data\" & PicIndex & ".inf" For Input As #1
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim lp As Long
    For lp = 1 To MaxCircle
        ShpCircle(lp).Visible = False
        ShpCircle(lp).Tag = ""
    Next lp
    lp = 0
    Do Until EOF(1)
        Input #1, x1, y1, x2, y2
        lp = lp + 1
        With ShpCircle(lp)
            .Left = x1
            .Top = y1
            .Width = x2 - x1
            .Height = y2 - y1
        End With
    Loop
    Close #1
    lngDiffFound = 0
    IsClickEnabled = True
    LoadPic = lp
    Form_Resize
    Exit Function
ErrHandler:
    If Err.Number = 53 Then
        MsgBox "无法找到对应的图片或信息文件。请与本软件技术负责人联系。", vbCritical
    Else
        MsgBox "读取图片文件时发生错误 " & Err.Number & ", " & Err.Description & ", 请与本软件技术负责人联系。", vbCritical
    End If
End Function

Private Sub Form_Load()
    Dim ret As Long
    Dim i As Long
    For i = 1 To MaxCircle
        Load ShpCircle(i)
    Next i
    lblStatus.Caption = "准备就绪"
    HScroll.Enabled = False
    lngCurrentPicNum = 1
    PicKid.Tag = LoadPic(1)
    '''''
    ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, ret
    Me.Show
    For i = 0 To 255 Step 2
        SetLayeredWindowAttributes Me.hWnd, 0, i, LWA_ALPHA
        DoEvents
        Sleep 1
    Next i
End Sub

Private Sub Form_Resize()
    If IsExit = True Then Exit Sub
    HScroll.Top = FrmMain.ScaleHeight - HScroll.Height
    HScroll.Width = FrmMain.ScaleWidth
    HScroll.Left = 0
    If PicMain.ScaleWidth + PicKid.ScaleWidth + 10 > FrmMain.ScaleWidth Then
        HScroll.Max = PicMain.ScaleWidth + PicKid.ScaleWidth + 10 - FrmMain.ScaleWidth
        HScroll.Enabled = True
        HScroll_Change
    Else
        ' Center it
        PicMain.Left = FrmMain.ScaleWidth / 2 - 5 - PicMain.ScaleWidth
        PicKid.Left = FrmMain.ScaleWidth / 2 + 5
        HScroll.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsExit = True
    Dim i As Long
    For i = 255 To 0 Step -2
        SetLayeredWindowAttributes Me.hWnd, 0, i, LWA_ALPHA
        DoEvents
        Sleep 1
    Next i
End Sub

Private Sub HScroll_Change()
    On Error Resume Next
    ' Redraw the pics
    PicMain.Left = -1 * HScroll.Value
    PicKid.Left = PicMain.Left + PicMain.Width + 10
End Sub

Private Sub PicKid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Do While IsClicking = True
        DoEvents
    Loop
    IsClicking = True
    If IsClickEnabled = True Then
        IsClickEnabled = False
        For i = 1 To CLng(PicKid.Tag)
            With ShpCircle(i)
                If X > .Left And X < .Left + .Width And Y > .Top And Y < .Top + .Height And .Tag <> "x" Then
                    ShpCircle(i).Visible = True
                    .Tag = "x"
                    lngDiffFound = lngDiffFound + 1
                    IsClickEnabled = True
                    SetStat ":)"
                    If CLng(PicKid.Tag) - lngDiffFound = 0 Then
                        ' Jump to the next pic
                        lngCurrentPicNum = lngCurrentPicNum + 1
                        SetStat "正在读取下一组图片..."
                        DoEvents
                        Sleep 1000
                        PicKid.Tag = LoadPic(lngCurrentPicNum)
                        If CLng(PicKid.Tag) <= 0 Then
                            IsClickEnabled = False
                            Exit Sub
                        End If
                        SetStat "准备就绪。"
                    End If
                    Exit For
                Else
                    SetStat ":("
                End If
            End With
        Next i
        tmrClickCheck.Enabled = True
    Else
        SetStat "点击过快！"
    End If
    IsClicking = False
End Sub

Private Sub PicMain_Click()
    SetStat "左边是原图，右边才是要点击的！"
End Sub

Private Sub tmrClickCheck_Timer()
    IsClickEnabled = True
    SetStat "等待下一次点击..."
    tmrClickCheck.Enabled = False
End Sub

Public Sub SetStat(ByVal strEvents As String)
    Dim i As Long, r As Integer
    lblStatus.Caption = "当前是第 " & lngCurrentPicNum & " 对图片。已找到 " & lngDiffFound & " 处不同，剩余 " & CLng(PicKid.Tag) - lngDiffFound & " 处。"
    lblStatus.Caption = lblStatus.Caption & vbCrLf & Date & " " & Time & " " & strEvents
End Sub
