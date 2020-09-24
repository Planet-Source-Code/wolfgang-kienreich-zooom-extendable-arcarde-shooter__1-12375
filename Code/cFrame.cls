VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cFrames

Private I_oDDRessourceSurface As DirectDrawSurface7

Public Terminating As Boolean
Public Sound As DirectSoundBuffer
Public SoundDelay As Long
Public AnimationDelay As Long

Private I_nTileWidth As Long
Private I_nTileHeight As Long

Private I_nTileCount As Long
Private I_dTileArea As RECT

Public Function RenderToViewport(nDstX As Long, nDstY As Long, ByVal nFrame As Long) As Boolean
    
    Dim L_dSrcArea As RECT
    Dim L_dDstArea As RECT
    Dim L_dMyArea As RECT
    Dim L_dVPArea As RECT
    Dim L_oDDDestinationSurface As DirectDrawSurface7
    
    Set L_oDDDestinationSurface = Parent.Parent.Viewport.Surface
    
    Let L_dVPArea = Parent.Parent.Viewport.Area
    
    With L_dMyArea
        .Left = nDstX
        .Top = nDstY
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    RenderToViewport = (IntersectRect(L_dDstArea, L_dMyArea, L_dVPArea) <> 0)
    
    If RenderToViewport Then
            
        With L_dSrcArea
        
            .Left = Width * (nFrame Mod I_nTileWidth)
            .Top = Height * (nFrame \ I_nTileWidth)
            .Right = .Left + Width
            .Bottom = .Top + Height

            .Top = .Top + IIf(L_dMyArea.Top < 0, Abs(L_dMyArea.Top), 0)
            .Left = .Left + IIf(L_dMyArea.Left < 0, Abs(L_dMyArea.Left), 0)
            .Bottom = .Bottom - IIf(L_dMyArea.Bottom > L_dVPArea.Bottom, L_dMyArea.Bottom - L_dVPArea.Bottom, 0)
            .Right = .Right - IIf(L_dMyArea.Right > L_dVPArea.Right, L_dMyArea.Right - L_dVPArea.Right, 0)
        
        End With
        
        nDstX = IIf(nDstX < 0, 0, nDstX)
        nDstY = IIf(nDstY < 0, 0, nDstY)
        
        L_oDDDestinationSurface.BltFast nDstX, nDstY, I_oDDRessourceSurface, L_dSrcArea, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    End If
    
End Function

Public Function RenderToWindow(nDstX As Long, nDstY As Long, ByVal nFrame As Long) As Boolean
    
    Dim L_dSrcArea As RECT
    Dim L_dDstArea As RECT
    Dim L_oDDDestinationSurface As DirectDrawSurface7
    
    Set L_oDDDestinationSurface = Parent.Parent.Viewport.Primary
    
    With L_dDstArea
        .Left = nDstX + Parent.Parent.Viewport.Left
        .Top = nDstY + Parent.Parent.Viewport.Top
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    With L_dSrcArea
        .Left = Width * (nFrame Mod I_nTileWidth)
        .Top = Height * (nFrame \ I_nTileWidth)
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    L_oDDDestinationSurface.Blt L_dDstArea, I_oDDRessourceSurface, L_dSrcArea, DDBLT_WAIT Or DDBLT_KEYSRC
        
    RenderToWindow = True
    
End Function

Public Function RenderToFrame(oDestination As cFrame, nDstX As Long, nDstY As Long, ByVal nFrame As Long) As Boolean
    
    Dim L_dSrcArea As RECT
    Dim L_dDstArea As RECT
    
    With L_dDstArea
        .Left = nDstX
        .Top = nDstY
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    With L_dSrcArea
        .Left = Width * (nFrame Mod I_nTileWidth)
        .Top = Height * (nFrame \ I_nTileWidth)
        .Right = .Left + Width
        .Bottom = .Top + Height
    End With
    
    oDestination.Surface.Blt L_dDstArea, I_oDDRessourceSurface, L_dSrcArea, DDBLT_WAIT Or DDBLT_KEYSRC
        
    RenderToFrame = True
    
End Function

Public Function Surface() As DirectDrawSurface7
    Set Surface = I_oDDRessourceSurface
End Function

Public Function TileCount() As Long
    TileCount = I_nTileCount
End Function

Public Function Width() As Long
    Width = I_dTileArea.Right - I_dTileArea.Left
End Function
Public Function Height() As Long
    Height = I_dTileArea.Bottom - I_dTileArea.Top
End Function

Public Sub Initialize(sFile As String, nTileWidth As Long, nTileHeight As Long)

    Dim L_dDDSD As DDSURFACEDESC2
    Dim L_dDDCK As DDCOLORKEY
    
    L_dDDSD.lFlags = DDSD_CAPS
    L_dDDSD.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
    If Not sFile = "" Then
        
        Set I_oDDRessourceSurface = Parent.Parent.DDInstance.CreateSurfaceFromFile(sFile, L_dDDSD)
        L_dDDCK.low = 0
        L_dDDCK.high = 0
        I_oDDRessourceSurface.SetColorKey DDCKEY_SRCBLT, L_dDDCK
        
        I_nTileCount = nTileWidth * nTileHeight
        
        I_oDDRessourceSurface.GetSurfaceDesc L_dDDSD
        With I_dTileArea
            .Top = 0
            .Bottom = L_dDDSD.lHeight \ nTileHeight
            .Left = 0
            .Right = L_dDDSD.lWidth \ nTileWidth
        End With
        
        I_nTileWidth = nTileWidth
        I_nTileHeight = nTileHeight
        
        AnimationDelay = 1
    
    End If
    
    
End Sub

Public Sub Terminate()

    Set Sound = Nothing
    Set I_oDDRessourceSurface = Nothing
    
End Sub

