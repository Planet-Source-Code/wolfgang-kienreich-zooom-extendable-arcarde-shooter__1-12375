VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWaypoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cWaypoints

Public Orientation As String
Public Position As Long
Public AttackPropability As Long

Public Direction As Long
Public Speed As Long

Public Style As Long

Public StopLevel As Boolean
Public ResumeLevel As Boolean

Public LoopTo As Long

Public CircleDir As Long
Public CircleRad As Long
Public CircleXit As Long

Public Sub Apply()

    If StopLevel Then
        Set Parent.Parent.Parent.Parent.Parent.Level.Stopper = Parent.Parent
    End If
    If ResumeLevel Then
        Set Parent.Parent.Parent.Parent.Parent.Level.Stopper = Nothing
    End If
    If Style = 1 Then CircleXit = Direction \ (2 * CircleRad)
    
    If Me.Direction > 500 Then
        
        Dim L_nDX As Long
        Dim L_nDY As Long
        
        L_nDX = Parent.Parent.Parent.Parent.Parent.Player.PositionX - Parent.Parent.PositionX
        L_nDY = Parent.Parent.Parent.Parent.Parent.Player.PositionY - Parent.Parent.PositionY
        
        If L_nDX = 0 Then L_nDX = 1

        If L_nDX >= 0 And L_nDY >= 0 Then
            Direction = 90 + Abs(Atn(L_nDY / L_nDX)) * 57.297469
        End If
        If L_nDX < 0 And L_nDY >= 0 Then
            Direction = 270 - Abs(Atn(L_nDY / L_nDX)) * 57.297469
        End If
        If L_nDX >= 0 And L_nDY < 0 Then
            Direction = 90 - Abs(Atn(L_nDY / L_nDX)) * 57.297469
        End If
        If L_nDX < 0 And L_nDY < 0 Then
            Direction = 270 + Atn(L_nDY / L_nDX) * 57.297469
        End If
                
    End If
    
    Set Parent.Parent.State = New cWaypoint
    With Parent.Parent.State
        .AttackPropability = Me.AttackPropability
        .CircleDir = Me.CircleDir
        .CircleRad = Me.CircleRad
        .CircleXit = Me.CircleXit
        .Direction = Me.Direction
        .LoopTo = Me.LoopTo
        .Orientation = Me.Orientation
        .Position = Me.Position
        .ResumeLevel = Me.ResumeLevel
        .Speed = Me.Speed
        .StopLevel = Me.StopLevel
        .Style = Me.Style
    End With
    
End Sub
