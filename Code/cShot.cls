VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cShots
Public PlayerShot As Boolean
Public Frame As cFrame
Public ExploStyle As String
Public Power As Long

Public PositionX As Long
Public PositionY As Long
Public DeltaX As Long
Public DeltaY As Long

Public Terminating As Boolean

Public Sub Update()
    PositionX = PositionX + DeltaX
    PositionY = PositionY + DeltaY
End Sub

Public Sub Render()

    Dim L_nCurrentFrame As Long
       
    L_nCurrentFrame = (Parent.Parent.Parent.FrameCount Mod (Frame.TileCount * Frame.AnimationDelay)) \ Frame.AnimationDelay
   
    If Not Frame.RenderToViewport(PositionX \ 10 - Frame.Width \ 2, PositionY \ 10 - Frame.Height \ 2, L_nCurrentFrame) Then
        Terminating = True
    End If
    
End Sub


