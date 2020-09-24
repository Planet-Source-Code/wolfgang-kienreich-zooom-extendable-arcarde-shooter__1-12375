VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Frame As cFrame

Public Terminating As Boolean
Private I_nMaxHitpoints As Long

Public ComponentX As Long
Public ComponentY As Long
Public ComponentParent As cEnemy
Private I_nCurrentHitpoints As Long
Public ExplosionSize As Long
Public PowerUpStyle As Long

Public NoCollide As Boolean
Public NoHit As Boolean

Public ScoreValue As Long

Public IsPowerup As Boolean

Public Hitpoints As Long
Public Animated As Boolean
Public Directed As Boolean
Public Targetting As Boolean
Public Shooting As Boolean
Public Parent As cEnemies
Private I_oWayPoints As cWaypoints

Public Active As Boolean

Public State As cWaypoint

Public LifeCount As Long
Public TriggerAt As Long

Public PositionX As Single
Public PositionY As Single

Public Weapon As Long

Private I_nShooting As Long
Private I_nTerminateCount As Long

Public Sub InitHitpoints()
    I_nMaxHitpoints = Hitpoints
    I_nCurrentHitpoints = Hitpoints
End Sub

Public Property Get Waypoints() As cWaypoints
    Set Waypoints = I_oWayPoints
End Property
Public Property Set Waypoints(oWaypoints As cWaypoints)
    Set I_oWayPoints = oWaypoints
End Property

Public Sub Initialize()

    Set I_oWayPoints = New cWaypoints
    Set I_oWayPoints.Parent = Me
    
End Sub

Public Sub Terminate()

End Sub

Public Sub Render()

    Dim L_nCurrentFrame As Long
    Dim L_nDX As Single
    Dim L_nDY As Single
    Dim L_nDirection As Single
    Dim L_nAngle As Single
    
    With Frame
    
         If Animated Then
             L_nCurrentFrame = (Parent.Parent.Parent.FrameCount Mod (.TileCount * .AnimationDelay)) \ .AnimationDelay
         End If
         
         If Directed Then
             L_nCurrentFrame = State.Direction \ (360 / (.TileCount + 1))
         End If
        
         If Targetting Then
            
            L_nDX = Parent.Parent.Parent.Player.PositionX - PositionX
            L_nDY = Parent.Parent.Parent.Player.PositionY - PositionY
            If L_nDX = 0 Then L_nDX = 1

            If L_nDX >= 0 And L_nDY >= 0 Then
                L_nAngle = 90 + Abs(Atn(L_nDY / L_nDX)) * 57.297469
            End If
            If L_nDX < 0 And L_nDY >= 0 Then
                L_nAngle = 270 - Abs(Atn(L_nDY / L_nDX)) * 57.297469
            End If
            If L_nDX >= 0 And L_nDY < 0 Then
                L_nAngle = 90 - Abs(Atn(L_nDY / L_nDX)) * 57.297469
            End If
            If L_nDX < 0 And L_nDY < 0 Then
                L_nAngle = 270 + Atn(L_nDY / L_nDX) * 57.297469
            End If
            
            L_nCurrentFrame = L_nAngle \ (360 / (.TileCount + 1))
            
         End If
         
         If Shooting Then
             L_nCurrentFrame = IIf(I_nShooting > 0, I_nShooting \ Frame.AnimationDelay, 0)
             If I_nShooting > 0 Then
                I_nShooting = I_nShooting + 1
                If I_nShooting = Frame.TileCount * Frame.AnimationDelay Then I_nShooting = 0
             End If
         End If
         
         If Not .RenderToViewport(PositionX \ 10 - Frame.Width \ 2, PositionY \ 10 - Frame.Height \ 2, L_nCurrentFrame) Then
             If ComponentParent Is Nothing Then
                If State.Style = 0 Then
                    If Parent.Parent.Parent.Level.Stopper Is Me Then Set Parent.Parent.Parent.Level.Stopper = Nothing
                    Active = False
                    Dim L_oEnemy As cEnemy
                    For Each L_oEnemy In Parent
                        If L_oEnemy.ComponentParent Is Me Then
                            L_oEnemy.Active = False
                        End If
                    Next
                End If
            End If
         End If
        
         If Parent.Parent.Parent.DebugMode Then
            If ComponentParent Is Nothing Then
        
                L_nDirection = (IIf(State.Direction - 90 < 0, State.Direction - 90 + 360, State.Direction - 90) * PIFACTOR)
                
                L_nDX = (State.Speed * Cos(L_nDirection) * 10)
                L_nDY = (State.Speed * Sin(L_nDirection) * 10)
                
                Parent.Parent.Parent.Viewport.Surface.SetForeColor RGB(255, 100, 100)
                Parent.Parent.Parent.Viewport.Surface.DrawLine PositionX / 10, PositionY / 10, PositionX / 10 + L_nDX, PositionY / 10 + L_nDY
         
            End If
         End If
         
    End With
    
End Sub

Public Sub Update()

    Dim L_nDY As Single
    Dim L_nDX As Single
    Dim L_nDirection As Single
    Dim L_oShot As cShot
    
    If Active Then
            
        If ComponentParent Is Nothing Then
        
            L_nDirection = (IIf(State.Direction - 90 < 0, State.Direction - 90 + 360, State.Direction - 90) * PIFACTOR)
            
            L_nDX = (State.Speed * Cos(L_nDirection))
            L_nDY = (State.Speed * Sin(L_nDirection))
            
            PositionX = PositionX + L_nDX
            PositionY = PositionY + L_nDY
            
        Else
            
            Set State = ComponentParent.State
            If ComponentParent.Terminating Then
                Terminating = True
                Parent.Parent.Effects.Add "EXPLO3", Int(PositionX), Int(PositionY), 0, 0
            End If
            PositionX = ComponentParent.PositionX + ComponentX * 10
            PositionY = ComponentParent.PositionY + ComponentY * 10
            
        End If
        
        If I_nTerminateCount > 0 Then
        
            If Not NoHit Then
                Set L_oShot = Parent.Parent.Shots.Hit(PositionX \ 10 - Frame.Width \ 2, PositionY \ 10 - Frame.Height \ 2, PositionX \ 10 + Frame.Width \ 2, PositionY \ 10 + Frame.Height \ 2, True)
                If Not (L_oShot Is Nothing) Then
                    L_oShot.Terminating = True
                    Parent.Parent.Parent.Player.Stat_OwnShotsHit = Parent.Parent.Parent.Player.Stat_OwnShotsHit + 1
                End If
            End If
            
            If I_nTerminateCount Mod 20 = 0 Then
                If ExplosionSize > 2 Then
                    Parent.Parent.Effects.Add "EXPLO3", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                End If
            End If
            If I_nTerminateCount Mod (10 + IIf(Not Parent.Parent.Parent.Detail, 5, 0)) = 0 Then
                Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
            End If
            If I_nTerminateCount Mod (15 + IIf(Not Parent.Parent.Parent.Detail, 5, 0)) = 0 Then
                Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
            End If

            I_nTerminateCount = I_nTerminateCount - 1
            If I_nTerminateCount = 0 Then
                Me.Terminating = True
                Parent.Parent.Parent.Player.Stat_Frags = Parent.Parent.Parent.Player.Stat_Frags + 1
                If Me.ComponentParent Is Nothing Then
                    If Parent.Parent.Parent.Level.Stopper Is Me Then Set Parent.Parent.Parent.Level.Stopper = Nothing
                End If
                Parent.Parent.Parent.Campaign.CheckExtraLife Me.ScoreValue
                
                Parent.Parent.Parent.Player.Score = Parent.Parent.Parent.Player.Score + Me.ScoreValue
                Parent.Parent.Effects.Add "EXPLO3", Int(PositionX), Int(PositionY), 0, IIf(Parent.Parent.Parent.Level.Stopper Is Nothing, 5, 0)
                If ExplosionSize = 1 Then
                    Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                End If
                If ExplosionSize = 2 Then
                    Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                    Parent.Parent.Effects.Add "EXPLO4", Int(PositionX), Int(PositionY), 0, 0, Me
                End If
                If ExplosionSize = 3 Then
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO3", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    Parent.Parent.Effects.Add "EXPLO3", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    Parent.Parent.Effects.Add "EXPLO4", Int(PositionX), Int(PositionY), 0, 0, Me
                End If
                If ExplosionSize = 4 Then
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO3", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO3", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    Parent.Parent.Effects.Add "EXPLO3", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    Parent.Parent.Effects.Add "EXPLO3", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO4", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO4", Int(PositionX) + Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(PositionY) + Int(Rnd * Frame.Height * 6) - Frame.Height * 3, 0, 0
                    Parent.Parent.Effects.Add "EXPLO4", Int(PositionX), Int(PositionY), 0, 0, Me
                End If
            End If
            
        Else
        
            If Me.ExplosionSize > 1 And Me.ExplosionSize < 4 Then
                If I_nCurrentHitpoints < Hitpoints \ 4 Then
                    If Me.ComponentParent Is Nothing Then
                        If Parent.Parent.Parent.FrameCount Mod (10 + IIf(Not Parent.Parent.Parent.Detail, 10, 0)) = 0 Then
                            Parent.Parent.Effects.Add "SMOKESHIP", Int(PositionX), Int(PositionY), -Int(L_nDX / 2), -Int(L_nDY / 2)
                        End If
                    End If
                End If
            End If
        
            If Parent.Parent.Parent.Player.Active And Active Then
            
                If Not IsPowerup Then
                
                    If Int(PositionY) < 4500 Then
                    
                        Select Case Weapon
                        
                            ' Single Ball, straight down
                            Case 1
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT1", Int(PositionX), Int(PositionY) + Frame.Height * 4, Int(PositionX), 5000
                                        I_nShooting = 1
                                    End If
                                End If
                        
                            ' Single Ball, directed at player
                            Case 2
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT1", Int(PositionX), Int(PositionY), Parent.Parent.Parent.Player.PositionX, Parent.Parent.Parent.Player.PositionY
                                        I_nShooting = 1
                                    End If
                                End If
                                
                            ' Single ball, in direction of enemy
                            Case 3
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT1", Int(PositionX), Int(PositionY), Int(PositionX) + L_nDX * 10000, Int(PositionY) + L_nDY * 10000
                                        I_nShooting = 1
                                    End If
                                End If
                                
                            ' Fireball, straight down
                            Case 4
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT3", Int(PositionX), Int(PositionY) + Frame.Height * 4, Int(PositionX), Int(PositionY) + 10000
                                        I_nShooting = 1
                                    End If
                                End If
                                
                            ' Fireball, directed at player
                            Case 5
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT3", Int(PositionX), Int(PositionY), Parent.Parent.Parent.Player.PositionX, Parent.Parent.Parent.Player.PositionY
                                        I_nShooting = 1
                                    End If
                                End If
                            
                            ' fireball ball, in direction of enemy
                            Case 6
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT3", Int(PositionX), Int(PositionY), Int(PositionX) + L_nDX * 10000, Int(PositionY) + L_nDY * 10000
                                        I_nShooting = 1
                                    End If
                                End If
                            
                            ' Fireline, straight down
                            Case 7
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT2", Int(PositionX), Int(PositionY) + Frame.Height * 4, Int(PositionX), Int(PositionY) + 10000
                                        I_nShooting = 1
                                    End If
                                End If
                            
                            '  Single fireball (down with deviation)
                            Case 8
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT3", Int(PositionX), Int(PositionY) + Frame.Height * 4, Int(PositionX) - 150 + Int(Rnd * 300), 5000
                                        I_nShooting = 1
                                    End If
                                End If
                            
                            '  Twin fireball (down with deviation)
                            Case 9
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT1", Int(PositionX) - Frame.Width * 2, Int(PositionY) + Frame.Height * 4, Int(PositionX) - Frame.Width * 2 - 150 + Int(Rnd * 300), 5000
                                        Parent.Parent.Shots.Add "ENEMYSHOT1", Int(PositionX) + Frame.Width * 2, Int(PositionY) + Frame.Height * 4, Int(PositionX) + Frame.Width * 2 - 150 + Int(Rnd * 300), 5000
                                        I_nShooting = 1
                                    End If
                                End If
                                
                            ' Mix: Overkill
                            Case 10
                                If Parent.Parent.Parent.FrameCount Mod 25 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT2", Int(PositionX), Int(PositionY) + Frame.Width * 4, Int(PositionX), 5000
                                        I_nShooting = 1
                                    End If
                                End If
                                If Parent.Parent.Parent.FrameCount Mod 35 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT3", Int(PositionX) - Frame.Width, Int(PositionY) + Frame.Width * 4, Int(PositionX) - Frame.Width - 150 + Rnd * 300, 5000
                                        I_nShooting = 1
                                    End If
                                End If
                                If Parent.Parent.Parent.FrameCount Mod 40 = 0 Then
                                    If Int(Rnd * 150) < State.AttackPropability Then
                                        Parent.Parent.Shots.Add "ENEMYSHOT3", Int(PositionX) + Frame.Width, Int(PositionY) + Frame.Width * 4, Int(PositionX) + Frame.Width - 150 + Rnd * 300, 5000
                                        I_nShooting = 1
                                    End If
                                End If
                                            
                        End Select
                    
                    End If
                    
                End If
            End If
            
            LifeCount = LifeCount + 1
            
            If Parent.Parent.Parent.FrameCount Mod 2 = 0 Then
            
                If Not IsPowerup Then
                    If Not NoHit Then
                        Set L_oShot = Parent.Parent.Shots.Hit(PositionX \ 10 - Frame.Width \ 2, PositionY \ 10 - Frame.Height \ 2, PositionX \ 10 + Frame.Width \ 2, PositionY \ 10 + Frame.Height \ 2, True)
                        If Not (L_oShot Is Nothing) Then
                            L_oShot.Terminating = True
                            I_nCurrentHitpoints = I_nCurrentHitpoints - L_oShot.Power * (3 - Parent.Parent.Parent.Difficulty)
                            Parent.Parent.Parent.Player.Stat_OwnShotsHit = Parent.Parent.Parent.Player.Stat_OwnShotsHit + 1
                            If I_nCurrentHitpoints < 0 Then
                                I_nTerminateCount = ExplosionSize * IIf(Parent.Parent.Parent.Detail, 30, 10) + 5
                            Else
                                Parent.Parent.Effects.Add L_oShot.ExploStyle, Int(PositionX), Int(PositionY), L_oShot.PositionX - Int(PositionX), L_oShot.PositionY - Int(PositionY), Me
                                If (I_nCurrentHitpoints - L_oShot.Power * (3 - Parent.Parent.Parent.Difficulty) < I_nMaxHitpoints \ 2) And (I_nCurrentHitpoints > I_nMaxHitpoints \ 2) And ExplosionSize > 2 Then
                                    Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO2", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                    If Parent.Parent.Parent.Detail Then Parent.Parent.Effects.Add "EXPLO1", Int(PositionX), Int(PositionY), Int(Rnd * Frame.Width * 6) - Frame.Width * 3, Int(Rnd * Frame.Height * 6) - Frame.Height * 3, Me
                                End If
                            End If
                        End If
                    End If
                End If
                
                If (Not NoCollide) And (I_nTerminateCount <= 0) And Not Parent.Parent.Parent.DebugMode Then
                    If Parent.Parent.Parent.Player.Active Then
                        With Parent.Parent.Parent.Player
                            If .PositionX + .Frame.Width * 2 > Int(PositionX) - Frame.Width * 5 Then
                                If .PositionX - .Frame.Width * 2 < Int(PositionX) + Frame.Width * 5 Then
                                    If .PositionY + .Frame.Height * 2 > Int(PositionY) - Frame.Height * 5 Then
                                        If .PositionY - .Frame.Height * 2 < Int(PositionY) + Frame.Height * 5 Then
                                        
                                            If IsPowerup Then
                                                Parent.Parent.Parent.Player.Stat_PowerupCollected = Parent.Parent.Parent.Player.Stat_PowerupCollected + 1
                                                Me.Terminating = True
                                                Parent.Parent.Effects.Add "EXPLOP", Int(PositionX), Int(PositionY), 0, IIf(Parent.Parent.Parent.Level.Stopper Is Nothing, 5, 0)
                                                
                                                Select Case PowerUpStyle
                                                    Case 1
                                                        If Parent.Parent.Parent.Player.Weapon = 0 Or Parent.Parent.Parent.Player.Weapon = 1 Or Parent.Parent.Parent.Player.Weapon = 2 Or Parent.Parent.Parent.Player.Weapon = 4 Then
                                                            Parent.Parent.Parent.Player.SetEquipment , IIf(Parent.Parent.Parent.Player.Mounts < 2, 2, Parent.Parent.Parent.Player.Mounts)
                                                            Parent.Parent.Message = "twin mount acquired"
                                                        End If
                                                    Case 2
                                                        If Parent.Parent.Parent.Player.Weapon = 0 Or Parent.Parent.Parent.Player.Weapon = 4 Then
                                                            Parent.Parent.Parent.Player.SetEquipment , 3
                                                            Parent.Parent.Message = "tripple mount acquired"
                                                        End If
                                                    Case 3
                                                        If Parent.Parent.Parent.Player.Weapon < 4 Then
                                                            Parent.Parent.Parent.Player.SetEquipment Parent.Parent.Parent.Player.Weapon + 1
                                                            Parent.Parent.Message = "weapon upgraded"
                                                        End If
                                                    Case 4
                                                        Parent.Parent.Parent.Player.Shield = Parent.Parent.Parent.Player.MaxHitpoints
                                                        Parent.Parent.Message = "shields charged"
                                                    Case 5
                                                        Parent.Parent.Parent.Player.Hitpoints = Parent.Parent.Parent.Player.MaxHitpoints
                                                        Parent.Parent.Message = "ship repaired"
                                                    Case 6
                                                        If Parent.Parent.Parent.Player.Orbiter < 4 Then
                                                            Parent.Parent.Parent.Player.SetEquipment , , Parent.Parent.Parent.Player.Orbiter + 1
                                                        End If
                                                        Parent.Parent.Message = "drones acquired"
                                                    Case 7
                                                        If Parent.Parent.Parent.Player.Lifes < 9 Then Parent.Parent.Parent.Player.Lifes = Parent.Parent.Parent.Player.Lifes + 1
                                                        Parent.Parent.Message = "extra ship acquired"
                                                End Select
                                                
                                            Else
                                                Parent.Parent.Parent.Player.Hitpoints = Parent.Parent.Parent.Player.Hitpoints - Int(I_nCurrentHitpoints * (Parent.Parent.Parent.Difficulty + 1) * 1.5)
                                                I_nCurrentHitpoints = -1
                                                I_nTerminateCount = 5
                                            End If
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
                
            End If
            
            If ComponentParent Is Nothing Then
            
                If State.Style = 1 Then
                    If LifeCount Mod State.CircleRad = 0 Then
                        State.Direction = State.Direction + 2 * State.CircleRad * State.CircleDir
                        If State.Direction > 359 Then State.Direction = State.Direction - 360
                        If State.Direction < 0 Then State.Direction = State.Direction + 360
                        If (State.Direction \ (2 * State.CircleRad)) = State.CircleXit Then
                            State.Style = 0
                        End If
                    End If
                End If
                
                If State.Style = 0 Then
                    If Not (I_oWayPoints.Current Is Nothing) Then
                        If I_oWayPoints.Current.Orientation = "H" Then
                            If L_nDX > 0 Then
                                If Int(PositionX) > Int(400 * (I_oWayPoints.Current.Position / 100)) * 10 Then
                                    I_oWayPoints.Current.Apply
                                    I_oWayPoints.Advance
                                    GoTo M_FinishedWaypoints
                                End If
                            End If
                            If L_nDX < 0 Then
                                If Int(PositionX) < Int(400 * (I_oWayPoints.Current.Position / 100)) * 10 Then
                                    I_oWayPoints.Current.Apply
                                    I_oWayPoints.Advance
                                    GoTo M_FinishedWaypoints
                                End If
                            End If
                        End If
                        If I_oWayPoints.Current.Orientation = "V" Then
                            If L_nDY > 0 Then
                                If Int(PositionY) > Int(500 * (I_oWayPoints.Current.Position / 100)) * 10 Then
                                    I_oWayPoints.Current.Apply
                                    I_oWayPoints.Advance
                                    GoTo M_FinishedWaypoints
                                End If
                            End If
                            If L_nDY < 0 Then
                                If Int(PositionY) < Int(500 * (I_oWayPoints.Current.Position / 100)) * 10 Then
                                    I_oWayPoints.Current.Apply
                                    I_oWayPoints.Advance
                                    GoTo M_FinishedWaypoints
                                End If
                            End If
                        End If
                    End If
                End If
M_FinishedWaypoints:
            End If
        End If
        
    Else
        
        If ComponentParent Is Nothing Then
        
            If Parent.Parent.Position = TriggerAt Then
                
                Active = True
                If IsPowerup Then
                Parent.Parent.Parent.Player.Stat_PowerupTotal = Parent.Parent.Parent.Player.Stat_PowerupTotal + 1
                End If
                I_nMaxHitpoints = Hitpoints
                I_nCurrentHitpoints = Hitpoints
                If Not I_oWayPoints.Current Is Nothing Then I_oWayPoints.Current.Apply
                I_oWayPoints.Advance
                            
                If State.Orientation = "H" Then
                    PositionX = 10 * (Int(400 * (State.Position / 100)))
                    PositionY = 10 * (-Frame.Height \ 2 + 1)
                End If
                
                If State.Orientation = "V" Then
                    PositionX = IIf(State.Direction > 180, 4000 + 10 * (Frame.Width \ 2 - 1), 10 * (-Frame.Width \ 2 + 1))
                    PositionY = 10 * (Int(500 * (State.Position / 100)))
                End If
                
                Dim L_oItem As cEnemy
                For Each L_oItem In Parent
                    If L_oItem.ComponentParent Is Me Then
                        With L_oItem
                            .Active = True
                            .InitHitpoints
                        End With
                    End If
                Next
                
            End If
            
            
        End If
        
    End If
    
End Sub

