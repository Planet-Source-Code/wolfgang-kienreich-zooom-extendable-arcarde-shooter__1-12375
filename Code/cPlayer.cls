VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Stat_OwnShotsFired As Long
Public Stat_OwnShotsHit As Long
Public Stat_EnemyShotsHit As Long
Public Stat_PowerupTotal As Long
Public Stat_PowerupCollected As Long
Public Stat_Frags As Long

Public Parent As cSession
Public Hitpoints As Long
Public Lifes As Long
Public MaxHitpoints As Long
Public Shield As Long

Public PositionX As Long
Public PositionY As Long

Public DeltaX As Long
Public DeltaY As Long

Public Weapon As Long
Public Mounts As Long
Public Orbiter As Long
Public Score As Long
Public Stars As Long

Private I_oMasterFrame As cFrame

Public StartDelay As Long

Private I_nShotCount As Long
Private I_nShotDelay As Long

Private I_nOrbiterCount As Long
Private I_nOrbiterDelay As Long
Private I_nOrbiterX As Long
Private I_nOrbiterY As Long
Private I_oOrbiterFrame As cFrame

Public Active As Boolean
Public Function Frame() As cFrame
    Set Frame = I_oMasterFrame
End Function
Public Sub SetEquipment(Optional ByVal nWeapon, Optional ByVal nMounts, Optional ByVal nOrbiter)

    If Not IsMissing(nWeapon) Then Weapon = nWeapon
    If Not IsMissing(nMounts) Then Mounts = nMounts
    If Not IsMissing(nOrbiter) Then Orbiter = nOrbiter
    
    Select Case Weapon
        
        Case 0
            I_nShotDelay = 15
            Set I_oMasterFrame = Parent.Frames.Item("PLAYER01")
            
        Case 1
            I_nShotDelay = 20
            Set I_oMasterFrame = Parent.Frames.Item("PLAYER02")
            If Mounts > 2 Then Mounts = 2
            
        Case 2
            I_nShotDelay = 25
            Set I_oMasterFrame = Parent.Frames.Item("PLAYER03")
            If Mounts > 2 Then Mounts = 2
            
        Case 3
            I_nShotDelay = 30
            Set I_oMasterFrame = Parent.Frames.Item("PLAYER04")
            If Mounts > 1 Then Mounts = 1
            
        Case 4
            I_nShotDelay = 35
            Set I_oMasterFrame = Parent.Frames.Item("PLAYER05" & IIf(Mounts = 1, "A", "B"))
            
    End Select
    
    I_nShotCount = I_nShotDelay
    
    I_nOrbiterDelay = 30
    I_nOrbiterCount = I_nOrbiterDelay
    
End Sub

Public Sub Initialize()

    MaxHitpoints = Parent.Campaign.PlayerHitpoints
    Lifes = Parent.Campaign.PlayerLifes
    Set I_oOrbiterFrame = Parent.Frames.Item("Orbiter")
    Reset
    
End Sub

Public Sub Reset(Optional bKeepEquipment As Boolean)

    PositionX = 2000
    PositionY = 6000
    I_nOrbiterX = PositionX
    I_nOrbiterY = PositionY
    DeltaX = 0
    DeltaY = 0
    If Not bKeepEquipment Then SetEquipment 0, 1, 0
    Hitpoints = MaxHitpoints
    Shield = 0
    
End Sub

Public Sub Render()

    Dim L_nCurrentFrame As Long
    
    L_nCurrentFrame = (DeltaX + 10) \ 2
    Select Case Weapon
        Case 0
            L_nCurrentFrame = L_nCurrentFrame + (Mounts - 1) * 11
        Case 1
            L_nCurrentFrame = L_nCurrentFrame + (Mounts - 1) * 11
        Case 2
            L_nCurrentFrame = L_nCurrentFrame + (Mounts - 1) * 11
        Case 3
            L_nCurrentFrame = L_nCurrentFrame
        Case 4
            L_nCurrentFrame = L_nCurrentFrame + IIf(Mounts = 1, 0, (Mounts - 2) * 11)
    End Select

    I_oMasterFrame.RenderToViewport PositionX \ 10 - I_oMasterFrame.Width \ 2, PositionY \ 10 - I_oMasterFrame.Height \ 2, L_nCurrentFrame
    
    Select Case Orbiter
    
        Case 1
            I_oOrbiterFrame.RenderToViewport I_nOrbiterX \ 10 - I_oOrbiterFrame.Width \ 2, I_nOrbiterY \ 10 - I_oOrbiterFrame.Height \ 2 + I_oMasterFrame.Height \ 2 + 10, 0
        
        Case 2
            I_oOrbiterFrame.RenderToViewport I_nOrbiterX \ 10 - I_oOrbiterFrame.Width \ 2 - I_oMasterFrame.Width \ 2 - 15, I_nOrbiterY \ 10 - I_oOrbiterFrame.Height \ 2 + I_oMasterFrame.Height \ 2 + 10, 0
            I_oOrbiterFrame.RenderToViewport I_nOrbiterX \ 10 - I_oOrbiterFrame.Width \ 2 + I_oMasterFrame.Width \ 2 + 15, I_nOrbiterY \ 10 - I_oOrbiterFrame.Height \ 2 + I_oMasterFrame.Height \ 2 + 10, 0
        
        Case 3
            I_oOrbiterFrame.RenderToViewport I_nOrbiterX \ 10 - I_oOrbiterFrame.Width \ 2 + 35 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR), I_nOrbiterY \ 10 - I_oOrbiterFrame.Height \ 2 + 35 * Sin((Parent.FrameCount Mod 180) * 2 * PIFACTOR), 0
        
        Case 4
            I_oOrbiterFrame.RenderToViewport I_nOrbiterX \ 10 - I_oOrbiterFrame.Width \ 2 + 35 * Cos(((Parent.FrameCount) Mod 180) * 2 * PIFACTOR), I_nOrbiterY \ 10 - I_oOrbiterFrame.Height \ 2 + 35 * Sin((Parent.FrameCount Mod 180) * 2 * PIFACTOR), 0
            I_oOrbiterFrame.RenderToViewport I_nOrbiterX \ 10 - I_oOrbiterFrame.Width \ 2 + 35 * Cos(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR), I_nOrbiterY \ 10 - I_oOrbiterFrame.Height \ 2 + 35 * Sin(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR), 0
        
    End Select
    
End Sub

Public Sub Update()

    If Not Active Then
    
        If StartDelay > 0 Then
            StartDelay = StartDelay - 1
            If StartDelay = 0 Then
                If Lifes < 0 Then
                    Parent.Mode = 6
                    Parent.Menu.DspItem = 0
                    Parent.Menu.DspOpen = True
                    Parent.Menu.VidOpen = True
                Else
                    Parent.Level.Reset
                End If
            End If
        Else
            PositionY = PositionY - 10
            If PositionY <= 4500 Then Active = True
        End If
        
        Exit Sub
        
    End If
    
    Dim L_oShot As cShot
    
    If Not Parent.DebugMode Then Set L_oShot = Parent.Level.Shots.Hit(PositionX \ 10 - I_oMasterFrame.Width \ 2, PositionY \ 10 - I_oMasterFrame.Width \ 2, PositionX \ 10 + I_oMasterFrame.Width \ 2, PositionY \ 10 + I_oMasterFrame.Height \ 2, False)
    If Not (L_oShot Is Nothing) Then
        Me.Stat_EnemyShotsHit = Me.Stat_EnemyShotsHit + 1
        Parent.Level.Effects.Add "SHIPHIT", L_oShot.PositionX, L_oShot.PositionY, 0, IIf(Parent.Level.Stopper Is Nothing, 5, 0)
        L_oShot.Terminating = True
        If Shield > 0 Then
            Shield = Shield - Int(L_oShot.Power * (Parent.Difficulty + 1))
            If Shield < 0 Then
                Hitpoints = Hitpoints + Shield
                Shield = 0
            End If
        Else
            Hitpoints = Hitpoints - Int(L_oShot.Power * (Parent.Difficulty + 1))
        End If
            
    End If
    
    If Hitpoints < 50 Then
        If Parent.FrameCount Mod (10 + IIf(Not Parent.Detail, 10, 0)) = 0 Then
            Parent.Level.Effects.Add "SMOKESHIP", PositionX, PositionY, 0, 10
            Parent.Level.Message = "critical damage"
        End If
    End If
    
    If Hitpoints < 0 Then
        
        Parent.Level.Message = "ship lost"
        
        Parent.Level.Effects.Add "EXPLO3", PositionX, PositionY, 0, 0
        Parent.Level.Effects.Add "EXPLO2", PositionX - I_oMasterFrame.Width * 5 + I_oMasterFrame.Width * 10 * Rnd, PositionY - I_oMasterFrame.Height \ 5 + I_oMasterFrame.Height * 10 * Rnd, 0, 0
        Parent.Level.Effects.Add "EXPLO2", PositionX - I_oMasterFrame.Width * 5 + I_oMasterFrame.Width * 10 * Rnd, PositionY - I_oMasterFrame.Height \ 5 + I_oMasterFrame.Height * 10 * Rnd, 0, 0
        Parent.Level.Effects.Add "EXPLO4", PositionX, PositionY, 0, 0
        
        Select Case Orbiter
        
            Case 1
                Parent.Level.Effects.Add "EXPLO3", I_nOrbiterX, I_nOrbiterY, 0, 0
                
            Case 2
                Parent.Level.Effects.Add "EXPLO3", I_nOrbiterX - I_oMasterFrame.Width * 5 - 150, I_nOrbiterY - I_oOrbiterFrame.Height * 5 + I_oMasterFrame.Height * 5 + 100, 0, 0
                Parent.Level.Effects.Add "EXPLO3", I_nOrbiterX + I_oMasterFrame.Width * 5 + 150, I_nOrbiterY - I_oOrbiterFrame.Height * 5 + I_oMasterFrame.Height * 5 + 100, 0, 0
            
            Case 3
                Parent.Level.Effects.Add "EXPLO3", (I_nOrbiterX + 350 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), (I_nOrbiterY + 350 * Sin((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), 0, 0
            
            Case 4
                Parent.Level.Effects.Add "EXPLO3", (I_nOrbiterX + 350 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), (I_nOrbiterY + 350 * Sin((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), 0, 0
                Parent.Level.Effects.Add "EXPLO3", (I_nOrbiterX + 350 * Cos(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR)), (I_nOrbiterY + 350 * Sin(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR)), 0, 0
            
        End Select
        
        Active = False
        Lifes = Lifes - 1
        If Lifes < 0 Then
            StartDelay = 300
            PositionY = 10000
            I_nOrbiterY = 10000
            If Score >= Parent.Campaign.HighScore Then
                Parent.Campaign.HighScore = Score
                Parent.Campaign.HighStars = Stars
                Parent.Campaign.WriteHigh
            End If
        Else
            StartDelay = 100
            Reset
        End If
        
        Exit Sub
        
    End If
    
    If I_nShotCount < I_nShotDelay Then I_nShotCount = I_nShotCount + 1
    If I_nOrbiterCount < I_nOrbiterDelay Then I_nOrbiterCount = I_nOrbiterCount + 1
    
    PositionX = PositionX + DeltaX
    PositionY = PositionY + DeltaY
    
    I_nOrbiterX = I_nOrbiterX + (PositionX - I_nOrbiterX) \ 20
    I_nOrbiterY = I_nOrbiterY + (PositionY - I_nOrbiterY) \ 20
    
    If PositionX < 300 Then
        DeltaX = 0
        PositionX = 300
    End If
    
    If PositionX > 3700 Then
        DeltaX = 0
        PositionX = 3700
    End If
    
    If PositionY < 3300 Then
        DeltaY = 0
        PositionY = 3300
    End If
    
    If PositionY > 4700 Then
        DeltaY = 0
        PositionY = 4700
    End If
    
End Sub

Public Sub Shoot()

     If I_nOrbiterCount = I_nOrbiterDelay Then
        
        Select Case Orbiter
        
            Case 1
                Parent.Level.Shots.Add "PLAYER2", I_nOrbiterX, I_nOrbiterY + I_oMasterFrame.Height * 5 + 10, I_nOrbiterX, 0, True
                
            Case 2
                Parent.Level.Shots.Add "PLAYER2", I_nOrbiterX - I_oMasterFrame.Width * 5 - 150, I_nOrbiterY - I_oOrbiterFrame.Height * 5 + I_oMasterFrame.Height * 5 + 100, I_nOrbiterX - I_oMasterFrame.Width * 5 - 150, 0, True
                Parent.Level.Shots.Add "PLAYER2", I_nOrbiterX + I_oMasterFrame.Width * 5 + 150, I_nOrbiterY - I_oOrbiterFrame.Height * 5 + I_oMasterFrame.Height * 5 + 100, I_nOrbiterX + I_oMasterFrame.Width * 5 + 150, 0, True
            
            Case 3
                Parent.Level.Shots.Add "PLAYER3", (I_nOrbiterX + 350 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), (I_nOrbiterY + 350 * Sin((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), (I_nOrbiterX + 250 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), 0, True
            
            Case 4
                Parent.Level.Shots.Add "PLAYER3", (I_nOrbiterX + 350 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), (I_nOrbiterY + 350 * Sin((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), (I_nOrbiterX + 250 * Cos((Parent.FrameCount Mod 180) * 2 * PIFACTOR)), 0, True
                Parent.Level.Shots.Add "PLAYER3", (I_nOrbiterX + 350 * Cos(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR)), (I_nOrbiterY + 350 * Sin(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR)), (I_nOrbiterX + 250 * Cos(((Parent.FrameCount + 90) Mod 180) * 2 * PIFACTOR)), 0, True
            
        End Select
        
        I_nOrbiterCount = 0
        
     End If
    
     If I_nShotCount = I_nShotDelay Then
   
        Select Case Weapon
        
            Case 0
                
                Select Case Mounts
                    Case 1
                        Parent.Level.Shots.Add "PLAYER1", PositionX + I_oMasterFrame.Width \ 2 - 20, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 20 + (Rnd * 500) - 250, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 1
                    Case 2
                        Parent.Level.Shots.Add "PLAYER1", PositionX + I_oMasterFrame.Width \ 2 - 95, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 95 + (Rnd * 500) - 250, 0, True
                        Parent.Level.Shots.Add "PLAYER1", PositionX + I_oMasterFrame.Width \ 2 + 55, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 + 55 + (Rnd * 500) - 250, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 2
                    Case Else
                        Parent.Level.Shots.Add "PLAYER1", PositionX + I_oMasterFrame.Width \ 2 - 95, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 95 + (Rnd * 500) - 250, 0, True
                        Parent.Level.Shots.Add "PLAYER1", PositionX + I_oMasterFrame.Width \ 2 - 20, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 20 + (Rnd * 500) - 250, 0, True
                        Parent.Level.Shots.Add "PLAYER1", PositionX + I_oMasterFrame.Width \ 2 + 55, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 + 55 + (Rnd * 500) - 250, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 3
                End Select
                
            Case 1
            
                Select Case Mounts
                    Case 1
                        Parent.Level.Shots.Add "PLAYER2", PositionX + I_oMasterFrame.Width \ 2 - 20, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 20, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 1
                    Case Else
                        Parent.Level.Shots.Add "PLAYER2", PositionX + I_oMasterFrame.Width \ 2 - 95, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 95, 0, True
                        Parent.Level.Shots.Add "PLAYER2", PositionX + I_oMasterFrame.Width \ 2 + 55, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 + 55, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 2
                End Select
                
            Case 2
            
                Select Case Mounts
                    Case 1
                        Parent.Level.Shots.Add "PLAYER3", PositionX + I_oMasterFrame.Width \ 2 - 160, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 160, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 1
                    Case Else
                        Parent.Level.Shots.Add "PLAYER3", PositionX + I_oMasterFrame.Width \ 2 - 160, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 160, 0, True
                        Parent.Level.Shots.Add "PLAYER3", PositionX + I_oMasterFrame.Width \ 2 + 110, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 + 110, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 2
                End Select
                
            Case 3
                Select Case Mounts
                    Case Else
                        Parent.Level.Shots.Add "PLAYER4", PositionX + I_oMasterFrame.Width \ 2 - 10, PositionY - 100, PositionX + I_oMasterFrame.Width \ 2 - 10, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 1
                End Select
                
            Case 4
                Select Case Mounts
                    Case 1
                        Parent.Level.Shots.Add "PLAYER5", PositionX + I_oMasterFrame.Width \ 2, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 1
                    Case 2
                        Parent.Level.Shots.Add "PLAYER5", PositionX + I_oMasterFrame.Width \ 2 - 150, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 1000, 0, True
                        Parent.Level.Shots.Add "PLAYER5", PositionX + I_oMasterFrame.Width \ 2 + 150, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 + 1000, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 2
                    Case Else
                        Parent.Level.Shots.Add "PLAYER5", PositionX + I_oMasterFrame.Width \ 2 - 150, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 - 1000, 0, True
                        Parent.Level.Shots.Add "PLAYER5", PositionX + I_oMasterFrame.Width \ 2, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2, 0, True
                        Parent.Level.Shots.Add "PLAYER5", PositionX + I_oMasterFrame.Width \ 2 + 150, PositionY - 50, PositionX + I_oMasterFrame.Width \ 2 + 1000, 0, True
                        Me.Stat_OwnShotsFired = Me.Stat_OwnShotsFired + 3
                End Select
                
        End Select
        
        I_nShotCount = 0
        
    End If
    
End Sub
