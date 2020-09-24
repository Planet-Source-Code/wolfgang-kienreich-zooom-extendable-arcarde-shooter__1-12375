VERSION 5.00
Begin VB.Form fMessage 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'Kein
   Caption         =   "Zooom"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrTimer 
      Interval        =   500
      Left            =   300
      Top             =   1875
   End
   Begin VB.PictureBox picLamp 
      BackColor       =   &H00080000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   240
      MousePointer    =   12  'Nicht ablegen
      Picture         =   "fMessage.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   225
      Width           =   300
   End
   Begin VB.PictureBox picLamp 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   240
      MousePointer    =   12  'Nicht ablegen
      Picture         =   "fMessage.frx":04F2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   225
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblMove 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   405
      MousePointer    =   15  'Größenänderung alle
      TabIndex        =   3
      Top             =   30
      Width           =   2415
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   405
      TabIndex        =   0
      Top             =   855
      Width           =   3525
   End
End
Attribute VB_Name = "fMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private I_oSession As cSession

Private dMouseStart As POINTAPI  ' Offset for window movement on screen
Private dWindowStart As POINTAPI ' Offset for window movement on screen

' LBLMOVE_MOUSEDOWN: Remember window and mouse start coordinates
Private Sub lblMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Only move with left button pressed...
    If Button = 1 Then
        
        ' Get mouse and window position
        GetCursorPos dMouseStart
        dWindowStart.X = Me.Left \ Screen.TwipsPerPixelX
        dWindowStart.Y = Me.Top \ Screen.TwipsPerPixelY
        
    End If
    
End Sub


' LBLMOVE_MOUSEDOWN: Adjust window to current coordinates of mouse
Private Sub lblMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Only move with left button pressed...
    If Button = 1 Then
        
        Dim dMouseNow As POINTAPI       ' Hold current mouse position
        Dim dWindowNow As POINTAPI      ' Hold current window position
        
        ' Get mouse and window position
        GetCursorPos dMouseNow
        dWindowNow.X = Me.Left \ Screen.TwipsPerPixelX
        dWindowNow.Y = Me.Top \ Screen.TwipsPerPixelY
        
        ' Set window to new position
        Me.Left = (dWindowStart.X + (dMouseNow.X - dMouseStart.X)) * Screen.TwipsPerPixelX
        Me.Top = (dWindowStart.Y + (dMouseNow.Y - dMouseStart.Y)) * Screen.TwipsPerPixelY
        
    End If
    
End Sub


Private Sub Form_Load()

    Set Me.Picture = LoadPicture(App.Path + "\imessage.bmp")
    
    Dim L_dPolyPoint(0 To 8) As POINTAPI
    Dim L_nPointIndex As Long
    Dim L_nPolyRegion As Long

    
    L_dPolyPoint(0).X = 5
    L_dPolyPoint(0).Y = 0
    L_dPolyPoint(1).X = 200
    L_dPolyPoint(1).Y = 0
    L_dPolyPoint(2).X = 230
    L_dPolyPoint(2).Y = 30
    L_dPolyPoint(3).X = 300
    L_dPolyPoint(3).Y = 30
    L_dPolyPoint(4).X = 300
    L_dPolyPoint(4).Y = 195
    L_dPolyPoint(5).X = 295
    L_dPolyPoint(5).Y = 200
    L_dPolyPoint(6).X = 5
    L_dPolyPoint(6).Y = 200
    L_dPolyPoint(7).X = 0
    L_dPolyPoint(7).Y = 195
    L_dPolyPoint(8).X = 0
    L_dPolyPoint(8).Y = 5
    
    For L_nPointIndex = 0 To UBound(L_dPolyPoint)
        With L_dPolyPoint(L_nPointIndex)
            .X = .X + Me.Left / Screen.TwipsPerPixelX
            .Y = .Y + Me.Top / Screen.TwipsPerPixelY
        End With
    Next
    
    L_nPolyRegion = CreatePolygonRgn(L_dPolyPoint(0), 9, 1)
    SetWindowRgn Me.hwnd, L_nPolyRegion, True

End Sub

Private Sub picLamp_Click(Index As Integer)
    Unload Me
End Sub

Private Sub tmrTimer_Timer()
    Me.picLamp(0).Visible = Not Me.picLamp(0).Visible
    Me.picLamp(1).Visible = Not Me.picLamp(1).Visible
End Sub
