VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cViewport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cSession
Public Terminating As Boolean

Private I_nHWND As Long
Private I_dWindowArea As RECT
Private I_oDDPrimarySurface As DirectDrawSurface7
Private I_oDDBackbufferSurface As DirectDrawSurface7
Private I_oDDClipper As DirectDrawClipper
Private I_dRedrawArea() As RECT

Public Function Left() As Long
    Left = I_dWindowArea.Left
End Function

Public Function Top() As Long
    Top = I_dWindowArea.Top
End Function

Public Function Primary() As DirectDrawSurface7
    Set Primary = I_oDDPrimarySurface
End Function

Public Property Get WindowHandle() As Long
    WindowHandle = I_nHWND
End Property
Public Property Let WindowHandle(nWindowHandle As Long)
    I_nHWND = nWindowHandle
    GetWindowRect I_nHWND, I_dWindowArea
End Property

Public Function Surface() As DirectDrawSurface7
    Set Surface = I_oDDBackbufferSurface
End Function

Public Function Area() As RECT
    With Area
        .Left = 0
        .Top = 0
        .Right = 400
        .Bottom = 500
    End With
End Function

Public Sub Rebuild()

    GetWindowRect I_nHWND, I_dWindowArea
    
End Sub
Public Sub Render()
    
    Dim L_dSrcArea As RECT
    Dim L_dDstArea As RECT
    
    If Parent.Level Is Nothing Then
        With L_dDstArea
            .Left = 0
            .Right = 400
            .Top = 0
            .Bottom = 500
        End With
        I_oDDBackbufferSurface.BltFast 0, 0, Parent.Frames.Item("IMENUBACK").Surface, L_dDstArea, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
    Else
    
        With L_dSrcArea
            .Left = 0
            .Right = 400
            .Top = Parent.Level.BackgroundPosition
            .Bottom = .Top + 500
        End With
        I_oDDBackbufferSurface.BltFast 0, 0, Parent.Level.Background, L_dSrcArea, DDBLTFAST_WAIT
        
    End If
    
    If Parent.Level Is Nothing Then
    
    Else
        Parent.Level.Enemies.Render
        Parent.Level.Shots.Render
        Parent.Player.Render
        Parent.Level.Effects.Render
    End If
    
    Parent.Menu.Render
    
    If Not Parent.Level Is Nothing Then
        
        I_oDDBackbufferSurface.SetForeColor RGB(192, 192, 224)
        
        fZooom.FontBold = True
        fZooom.FontSize = 12
        I_oDDBackbufferSurface.SetFont fZooom.Font
        I_oDDBackbufferSurface.DrawText 10, 10, Format(Parent.Player.Score, "00000000"), False
        
        fZooom.FontSize = 8
        I_oDDBackbufferSurface.SetFont fZooom.Font
        I_oDDBackbufferSurface.DrawText 310, 10, "UP", False
        I_oDDBackbufferSurface.DrawText 325, 10, Format(Parent.Campaign.NextExtraLife, "00000000"), False
        I_oDDBackbufferSurface.DrawText 310, 25, "HI", False
        I_oDDBackbufferSurface.DrawText 325, 25, Format(Parent.Campaign.HighScore, "00000000"), False
        
        If Parent.Level.MessageActive Then
            fZooom.FontSize = 8
            I_oDDBackbufferSurface.SetFont fZooom.Font
            I_oDDBackbufferSurface.DrawText 10, 28, Parent.Level.Message, False
        End If
        
        If Parent.DebugMode Then
            fZooom.FontSize = 8
            fZooom.FontBold = True
            I_oDDBackbufferSurface.SetFont fZooom.Font
            I_oDDBackbufferSurface.SetForeColor RGB(255, 48, 48)
            I_oDDBackbufferSurface.DrawText 160, 10, "Debug Mode", False
            I_oDDBackbufferSurface.DrawText 10, 480, "Level Position " & Format(Parent.Level.Position, "0"), False
            I_oDDBackbufferSurface.DrawText 310, 480, "(Collisions Off)", False

            I_oDDBackbufferSurface.DrawText 310, 235, "Update  " & Right(Format(Parent.UpdateTime, "00"), 2) & "ms", False
            I_oDDBackbufferSurface.DrawText 310, 250, "Render  " & Right(Format(Parent.RenderTime, "00"), 2) & "ms", False
            I_oDDBackbufferSurface.DrawText 310, 265, "Process " & Right(Format(Parent.ProcessTime, "00"), 2) & "ms", False
            
        End If
        
    End If
    
    If Parent.FrameCount Mod 20 = 0 Then
    
        If Parent.Level Is Nothing Then
            Parent.Frames.Item("METERBACK").RenderToWindow 438, 50, 0
            Parent.Frames.Item("LIFES").RenderToWindow 442, 185, 0
        Else
            Parent.Frames.Item("METERBACK").RenderToFrame Parent.Frames.Item("METER"), 0, 0, 0
            Parent.Frames.Item("HITMETER").RenderToFrame Parent.Frames.Item("METER"), 5, 3, 10 - (Parent.Player.Hitpoints / Parent.Player.MaxHitpoints) * 10
            Parent.Frames.Item("SHIELDMETER").RenderToFrame Parent.Frames.Item("METER"), 5, 17, 14 - (Parent.Player.Shield / Parent.Player.MaxHitpoints) * 14
            Parent.Frames.Item("METER").RenderToWindow 438, 50, 0
            Parent.Frames.Item("LIFES").RenderToWindow 442, 185, Parent.Player.Lifes
        End If
        
        Me.Primary.SetForeColor 0
        Me.Primary.DrawLine 175 + Left, 558 + Top, 276 + Left, 558 + Top
        Me.Primary.SetForeColor RGB(255 - Parent.FramesPerSecond * 2.5, Parent.FramesPerSecond * 2.5, 0)
        Me.Primary.DrawLine 175 + Left, 558 + Top, 175 + Left + Parent.FramesPerSecond, 558 + Top
                
    End If
        
    With L_dSrcArea
        .Left = 0
        .Right = 400
        .Top = 0
        .Bottom = 500
    End With
    
    With L_dDstArea
        .Left = I_dWindowArea.Left + 20
        .Right = .Left + 400
        .Top = I_dWindowArea.Top + 50
        .Bottom = .Top + 500
    End With
    
    I_oDDPrimarySurface.Blt L_dDstArea, I_oDDBackbufferSurface, L_dSrcArea, DDBLT_ASYNC
    
End Sub

Public Sub Initialize(nWindowHandle As Long)

    Dim L_dDDSD As DDSURFACEDESC2
    
    If nWindowHandle <> 0 Then Me.WindowHandle = nWindowHandle
    
    Parent.DDInstance.SetCooperativeLevel I_nHWND, DDSCL_NORMAL
    
    L_dDDSD.lFlags = DDSD_CAPS
    L_dDDSD.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set I_oDDPrimarySurface = Parent.DDInstance.CreateSurface(L_dDDSD)
    
    Set I_oDDClipper = Parent.DDInstance.CreateClipper(0)
    I_oDDClipper.SetHWnd I_nHWND
    I_oDDPrimarySurface.SetClipper I_oDDClipper
    
    L_dDDSD.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    L_dDDSD.lWidth = 400
    L_dDDSD.lHeight = 500
    L_dDDSD.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set I_oDDBackbufferSurface = Parent.DDInstance.CreateSurface(L_dDDSD)
    
    I_oDDBackbufferSurface.SetFont fZooom.Font
    I_oDDBackbufferSurface.SetFontTransparency True
   
End Sub

Public Sub Terminate()
    
    Set I_oDDBackbufferSurface = Nothing
    
    Set I_oDDClipper = Nothing
    Set I_oDDPrimarySurface = Nothing
    
End Sub


