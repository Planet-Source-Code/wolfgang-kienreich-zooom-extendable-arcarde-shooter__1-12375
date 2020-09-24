VERSION 5.00
Begin VB.Form fZooom 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'Kein
   Caption         =   "Zooom"
   ClientHeight    =   8550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fZooom.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   StartUpPosition =   2  'Bildschirmmitte
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
      Picture         =   "fZooom.frx":27A2
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   2
      Top             =   195
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
      Index           =   0
      Left            =   240
      MousePointer    =   12  'Nicht ablegen
      Picture         =   "fZooom.frx":2C94
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   195
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
      Left            =   480
      MousePointer    =   15  'Größenänderung alle
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "fZooom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private I_oSession As cSession

Private dMouseStart As POINTAPI  ' Offset for window movement on screen
Private dWindowStart As POINTAPI ' Offset for window movement on screen



Private Sub Form_Activate()
    
    If I_oSession Is Nothing Then
        Set I_oSession = New cSession
        Me.Refresh
        I_oSession.Initialize Me.hwnd
        I_oSession.Execute
        I_oSession.Terminate
        Unload Me
        End
    End If
    
End Sub

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
        I_oSession.Viewport.Rebuild
        
    End If
    
End Sub


Private Sub Form_Load()

    Set Me.Picture = LoadPicture(App.Path + "\interface.bmp")
    
    Dim L_dPolyPoint(0 To 8) As POINTAPI
    Dim L_nPointIndex As Long
    Dim L_nPolyRegion1 As Long
    Dim L_nPolyRegion2 As Long
    Dim L_nPolyRegionR As Long
    
    L_dPolyPoint(0).X = 5
    L_dPolyPoint(0).Y = 0
    L_dPolyPoint(1).X = 200
    L_dPolyPoint(1).Y = 0
    L_dPolyPoint(2).X = 230
    L_dPolyPoint(2).Y = 30
    L_dPolyPoint(3).X = 440
    L_dPolyPoint(3).Y = 30
    L_dPolyPoint(4).X = 440
    L_dPolyPoint(4).Y = 565
    L_dPolyPoint(5).X = 435
    L_dPolyPoint(5).Y = 570
    L_dPolyPoint(6).X = 5
    L_dPolyPoint(6).Y = 570
    L_dPolyPoint(7).X = 0
    L_dPolyPoint(7).Y = 565
    L_dPolyPoint(8).X = 0
    L_dPolyPoint(8).Y = 5
    
    For L_nPointIndex = 0 To UBound(L_dPolyPoint)
        With L_dPolyPoint(L_nPointIndex)
            .X = .X + Me.Left / Screen.TwipsPerPixelX
            .Y = .Y + Me.Top / Screen.TwipsPerPixelY
        End With
    Next
    
    L_nPolyRegion1 = CreatePolygonRgn(L_dPolyPoint(0), 9, 1)
    L_nPolyRegion2 = CreateEllipticRgn(380 + Me.Left / Screen.TwipsPerPixelX, 30 + Me.Top / Screen.TwipsPerPixelY, 500 + Me.Left / Screen.TwipsPerPixelX, 150 + Me.Top / Screen.TwipsPerPixelY)
    CombineRgn L_nPolyRegion1, L_nPolyRegion1, L_nPolyRegion2, 2
    L_nPolyRegion2 = CreateEllipticRgn(380 + Me.Left / Screen.TwipsPerPixelX, 160 + Me.Top / Screen.TwipsPerPixelY, 500 + Me.Left / Screen.TwipsPerPixelX, 280 + Me.Top / Screen.TwipsPerPixelY)
    CombineRgn L_nPolyRegion1, L_nPolyRegion1, L_nPolyRegion2, 2
    SetWindowRgn Me.hwnd, L_nPolyRegion1, True

End Sub

Private Sub picLamp_Click(Index As Integer)
    I_oSession.Terminate
    End
End Sub
