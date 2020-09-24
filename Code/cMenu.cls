VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cSession

Public BtnItem As Integer
Public BtnOpen As Boolean
Public DspItem As Integer
Public DspOpen As Boolean
Public VidOpen As Boolean

Private I_nBtnPos(3) As Single
Private I_nDspPos As Single
Private I_nVidPos As Single
Private I_nVidFrame As Long

Private I_sCampaign() As String
Private I_nHighscore() As Long
Private I_nHighStars() As Long
Private I_nCampaign As Long

Public MenuSound As DirectSoundBuffer

Public Sub Initialize()
    
    Dim L_dDSBufferDesc As DSBUFFERDESC
    Dim L_dWFMDesc As WAVEFORMATEX
    Dim L_nRun As Long
    
    With L_dDSBufferDesc
        .lBufferBytes = 0
        .lFlags = DSBCAPS_CTRLVOLUME
    End With
        
    If Parent.SoundPresent Then
        L_dWFMDesc = GetWaveFileFormat("zooom.wav")
        Set MenuSound = Parent.DSInstance.CreateSoundBufferFromFile(App.Path + "\zooom.wav", L_dDSBufferDesc, L_dWFMDesc)
        MenuSound.Play DSBPLAY_LOOPING
        Parent.MusicVolume = Parent.MusicVolume
    End If
    
    Call LoadCampaignInfo
    
    I_nCampaign = 0
    Parent.Campaign.Load I_sCampaign(I_nCampaign)
    
    For L_nRun = 0 To 3
        I_nBtnPos(L_nRun) = -L_nRun * 10
    Next
    I_nDspPos = 0
    I_nVidPos = 0
    I_nVidFrame = -1
    
    BtnItem = 0
    DspItem = 0
    BtnOpen = True
    DspOpen = False
    VidOpen = False
    
    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IBUTTONS").Sound)
    
End Sub

Public Sub LoadCampaignInfo()
    
    Dim L_sName As String
    Dim L_nCount As Long
    Dim L_nRun As Long
    
    ReDim I_sCampaign(0)
    ReDim I_nHighscore(0)
    ReDim I_nHighStars(0)
    
    L_sName = Dir(App.Path + "\level\*.cpn")
    Do While Not L_sName = ""
    
        ReDim Preserve I_sCampaign(L_nCount)
        ReDim Preserve I_nHighscore(L_nCount)
        ReDim Preserve I_nHighStars(L_nCount)
        
        I_sCampaign(L_nCount) = Left(L_sName, Len(L_sName) - 4)
    
        L_sName = Dir()
        L_nCount = L_nCount + 1
    Loop
    
    For L_nRun = 0 To L_nCount - 1
        If Dir(App.Path + "\level\" + I_sCampaign(L_nRun) + ".sco") <> "" Then
            Open App.Path + "\level\" + I_sCampaign(L_nRun) + ".sco" For Input As #7
            Input #7, I_nHighscore(L_nRun), I_nHighStars(L_nRun)
            Close #7
        End If
    Next
    
    Parent.Campaign.Load I_sCampaign(I_nCampaign)
    
End Sub

Public Sub Execute(Optional ByVal nAction As Integer)
    
    Select Case Parent.Mode
    
        Case 0
        
            If DspOpen Then
                
                If nAction = 5 Then
                    
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                    DspOpen = False
                    
                Else
            
                    Select Case BtnItem
                            
                        Case 1
                            
                            Select Case nAction
                            
                                Case 0
                                    Parent.Campaign.Clear
                                    Parent.Campaign.Load I_sCampaign(DspItem)
                                    I_nCampaign = DspItem
                                    DspOpen = False
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IBUTTONS").Sound)
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                                    
                                Case 3
                                    DspItem = DspItem - 1
                                    If DspItem < 0 Then DspItem = UBound(I_sCampaign)
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                                
                                Case 4
                                    DspItem = DspItem + 1
                                    If DspItem > UBound(I_sCampaign) Then DspItem = 0
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                                    
                            End Select
                            
                        Case 2
                        
                            Select Case nAction
                                Case 0
                                    DspOpen = False
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                                
                                Case 1, 2
                                
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                                    
                                    Select Case DspItem
                                        Case 0
                                            Parent.WaveVolume = Parent.WaveVolume + IIf(nAction = 2, 4, -4)
                                        Case 1
                                            Parent.MusicVolume = Parent.MusicVolume + IIf(nAction = 2, 4, -4)
                                        Case 2
                                            Parent.Detail = Not Parent.Detail
                                        Case 3
                                            If nAction = 1 Then
                                                If Parent.Difficulty > 0 Then Parent.Difficulty = Parent.Difficulty - 1
                                            End If
                                            If nAction = 2 Then
                                                If Parent.Difficulty < 2 Then Parent.Difficulty = Parent.Difficulty + 1
                                            End If
                                        Case 4
                                            Parent.Joystick = Not Parent.Joystick
                                        Case 5
                                            If nAction = 1 Then
                                                If Parent.JoystickDead > 10 Then Parent.JoystickDead = Parent.JoystickDead - 5
                                            End If
                                            If nAction = 2 Then
                                                If Parent.JoystickDead < 90 Then Parent.JoystickDead = Parent.JoystickDead + 5
                                            End If
                                            
                                    End Select
                                    
                                Case 3
                                
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                                
                                    DspItem = DspItem - 1
                                    If DspItem < 0 Then DspItem = 5
                                    
                                Case 4
                                
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                                
                                    DspItem = DspItem + 1
                                    If DspItem > 5 Then DspItem = 0
                            
                            End Select
                            
                        Case 3
                        
                            Select Case nAction
                            
                                Case 0
                                
                                    DspOpen = False
                                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                                    
                                Case 1, 2
                        
                                Case 3, 4
                                    
                            End Select
                        
                    End Select
                
                End If
                
            Else
            
                If nAction = 5 Then
                    Parent.Terminating = True
                Else
                
                    Select Case nAction
                    
                        Case 0
                        
                            If BtnItem = 0 Then
                            
                                If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IBUTTONS").Sound)
                                If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                                Parent.Campaign.CurrentLevel = 1
                                Set Parent.Player = New cPlayer
                                Set Parent.Player.Parent = Parent
                                Parent.Player.Initialize
                                Parent.Mode = 1
                                BtnOpen = False
                                If Parent.SoundPresent Then MenuSound.Stop
                            Else
                            
                                If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                            
                                DspOpen = True
                                If BtnItem = 1 Then LoadCampaignInfo
                                DspItem = IIf(BtnItem = 1, I_nCampaign, 0)
                                
                            End If
                        
                        Case 3
                        
                            If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IBUTTONS").Sound)
                        
                            BtnItem = BtnItem - 1
                            If BtnItem < 0 Then BtnItem = 3
                            DspOpen = False
                            
                        Case 4
                            
                            If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IBUTTONS").Sound)
                            
                            BtnItem = BtnItem + 1
                            If BtnItem > 3 Then BtnItem = 0
                            DspOpen = False
                    End Select
                    
                End If
            
            End If
            
        Case 2
        
        Case 3, 4
            
            Select Case nAction
            
                Case 0
                
                    DspOpen = False
                    VidOpen = False
                    I_nVidFrame = -1
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                    
                    Select Case DspItem
                        Case 0
                            If Not Parent.Level Is Nothing Then
                                Parent.Level.Terminate
                                Set Parent.Level = Nothing
                            End If
                            Parent.Player.Reset True
                            Parent.Mode = 1
                            
                        Case 1
                            If Parent.Mode = 4 Then
                                If Parent.Player.Score > Parent.Campaign.HighScore Then
                                    Parent.Campaign.HighScore = Parent.Player.Score
                                    Parent.Campaign.HighStars = Parent.Player.Stars
                                    Parent.Campaign.WriteHigh
                                End If
                            End If
                            If Not Parent.Level Is Nothing Then
                                Parent.Level.Terminate
                                Set Parent.Level = Nothing
                            End If
                            Parent.Mode = 0
                            BtnOpen = True
                            If Parent.SoundPresent Then
                                MenuSound.Play DSBPLAY_LOOPING
                                Parent.MusicVolume = Parent.MusicVolume
                            End If
                    End Select
                
                Case 3
                
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                
                    DspItem = DspItem - 1
                    If DspItem < 0 Then DspItem = 1
                    
                Case 4
                
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                
                    DspItem = DspItem + 1
                    If DspItem > 1 Then DspItem = 0
                
            End Select
            
        Case 6
        
            Select Case nAction
            
                Case 0
                
                    DspOpen = False
                    VidOpen = False
                    I_nVidFrame = -1
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                    
                    Select Case DspItem
                        Case 0
                            If Not Parent.Level Is Nothing Then
                                Parent.Level.Terminate
                                Set Parent.Level = Nothing
                            End If
                            Parent.Campaign.CurrentLevel = 1
                            Set Parent.Player = New cPlayer
                            Set Parent.Player.Parent = Parent
                            Parent.Player.Initialize
                            Parent.Mode = 1
                        Case 1
                            If Not Parent.Level Is Nothing Then
                                Parent.Level.Terminate
                                Set Parent.Level = Nothing
                            End If
                            Parent.Mode = 0
                            BtnOpen = True
                            If Parent.SoundPresent Then
                                MenuSound.Play DSBPLAY_LOOPING
                                Parent.MusicVolume = Parent.MusicVolume
                            End If
                    End Select
                
                Case 3
                
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                
                    DspItem = DspItem - 1
                    If DspItem < 0 Then DspItem = 1
                    
                Case 4
                
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
                
                    DspItem = DspItem + 1
                    If DspItem > 1 Then DspItem = 0
            
            End Select
            
        Case 5
        
            Select Case nAction
            
                Case 0
                
                    DspOpen = False
                    VidOpen = False
                    I_nVidFrame = -1
                    If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAY").Sound)
                    
                    Select Case DspItem
                        Case 0
                            Parent.Mode = 1
                        Case 1
                            If Not Parent.Level Is Nothing Then
                                Parent.Level.Terminate
                                Set Parent.Level = Nothing
                            End If
                            Parent.Campaign.CurrentLevel = 1
                            Set Parent.Player = New cPlayer
                            Set Parent.Player.Parent = Parent
                            Parent.Player.Initialize
                            Parent.Mode = 1
                            
                        Case 2
                            If Not Parent.Level Is Nothing Then
                                Parent.Level.Terminate
                                Set Parent.Level = Nothing
                            End If
                            BtnOpen = True
                            If Parent.SoundPresent Then
                                MenuSound.Play DSBPLAY_LOOPING
                                Parent.MusicVolume = Parent.MusicVolume
                            End If
                            Parent.Mode = 0
                            
                    End Select
                
            Case 3
            
                If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
            
                DspItem = DspItem - 1
                If DspItem < 0 Then DspItem = 2
                
            Case 4
            
                If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("IDISPLAYBACK").Sound)
            
                DspItem = DspItem + 1
                If DspItem > 2 Then DspItem = 0
        
        End Select
        
    End Select
    
End Sub
Public Sub Render()
    
    Dim L_nRun As Long
    
    For L_nRun = 0 To 3
        If I_nBtnPos(L_nRun) > 0 Then
            Parent.Frames.Item("IBUTTONS").RenderToViewport I_nBtnPos(L_nRun) - Parent.Frames.Item("IBUTTONS").Width, L_nRun * 50 + 50, L_nRun * 2 + IIf(BtnItem = L_nRun, 1, 0)
        End If
    Next
    
    If I_nDspPos > 0 Then
        Call RenderDisplayContents
        Parent.Frames.Item("IDISPLAY").RenderToViewport 200, I_nDspPos - Parent.Frames.Item("IDISPLAY").Height, 0
    End If
    
    If I_nVidPos > 0 Then
        Call RenderVideoContents
        Parent.Frames.Item("IVIDEO").RenderToViewport I_nVidPos - Parent.Frames.Item("IVIDEO").Width, 320, 0
    End If
    
End Sub

Public Sub Update()

    Dim L_nRun As Long
        
    For L_nRun = 0 To 3
        If BtnOpen Then
            If L_nRun = BtnItem Then
                I_nBtnPos(L_nRun) = I_nBtnPos(L_nRun) + (180 - I_nBtnPos(L_nRun)) / 10
            Else
                I_nBtnPos(L_nRun) = I_nBtnPos(L_nRun) + (160 - I_nBtnPos(L_nRun)) / 10
            End If
        Else
            I_nBtnPos(L_nRun) = I_nBtnPos(L_nRun) + (0 - I_nBtnPos(L_nRun)) / 10
        End If
    Next
    
    If DspOpen Then
        I_nDspPos = I_nDspPos + (300 - I_nDspPos) / 15
    Else
        I_nDspPos = I_nDspPos + (0 - I_nDspPos / 15)
    End If
    
    If VidOpen Then
        I_nVidPos = I_nVidPos + (200 - I_nVidPos) / 10
        If Int(I_nVidPos) = 199 And I_nVidFrame = -1 Then
            I_nVidFrame = 0
            If Parent.Mode = 3 Or Parent.Mode = 4 Then
                If Parent.SoundPresent Then Parent.Sounds.Add Parent.DSInstance.DuplicateSoundBuffer(Parent.Frames.Item("fanfare").Sound)
                If Parent.SoundPresent Then Parent.Sounds.Item(Parent.Sounds.Count).Sound.SetVolume (DSBVOLUME_MIN / 20)
            End If
        End If
    Else
        I_nVidFrame = -1
        I_nVidPos = I_nVidPos + (0 - I_nVidPos / 10)
    End If
    
End Sub

Public Sub Terminate()
    If Parent.SoundPresent Then MenuSound.Stop
    Set MenuSound = Nothing
End Sub

Private Sub RenderDisplayContents()

    Dim L_nRun As Long
    
    Parent.Frames.Item("IDISPLAYBACK").RenderToFrame Parent.Frames.Item("IDISPLAY"), 0, 0, 0
    
    With Parent.Frames.Item("IDISPLAY").Surface
        fZooom.FontSize = 8
        fZooom.FontBold = True
        .SetFont fZooom.Font
        .SetFontTransparency True
        .SetForeColor RGB(240, 240, 200)
    End With
    
    With Parent.Frames.Item("IDISPLAY").Surface
    
        Select Case Parent.Mode
        
            Case 0
            
                Select Case BtnItem
                    Case 0
                    Case 1
                        For L_nRun = 0 To UBound(I_sCampaign())
                            .SetForeColor IIf(L_nRun = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                            fZooom.FontBold = True
                            .SetFont fZooom.Font
                            .DrawText 25, L_nRun * 30 + 150, Left(UCase(I_sCampaign(L_nRun)), 16), False
                            fZooom.FontBold = False
                            .SetFont fZooom.Font
                            .DrawText 35, L_nRun * 30 + 165, "HI " & Format(I_nHighscore(L_nRun), "00000000") & "  " & Format(I_nHighStars(L_nRun), "0") & "*", False
                        Next
                    Case 2
                        .SetForeColor IIf(0 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                        .DrawText 25, 150, "SOUND", False
                        .DrawBox 70, 153, 123, 162
                        For L_nRun = 0 To Parent.WaveVolume \ 4
                            .DrawLine 70 + L_nRun * 2, 155, 70 + L_nRun * 2, 160
                        Next
                        
                        .SetForeColor IIf(1 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                        .DrawText 25, 165, "MUSIC", False
                        .DrawBox 70, 168, 123, 177
                        For L_nRun = 0 To Parent.MusicVolume \ 4
                            .DrawLine 70 + L_nRun * 2, 170, 70 + L_nRun * 2, 175
                        Next
                        
                        .SetForeColor IIf(2 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                        .DrawText 25, 180, "DETAIL", False
                        .DrawBox 70, 183, 123, 192
                        For L_nRun = 0 To (100 - IIf(Parent.Detail, 0, 50)) \ 4
                            .DrawLine 70 + L_nRun * 2, 185, 70 + L_nRun * 2, 190
                        Next
                        
                        .SetForeColor IIf(3 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                        .DrawText 25, 215, "DIFFICULTY " & IIf(Parent.Difficulty = 0, "LOW", IIf(Parent.Difficulty = 1, "AVG", "HIGH")), False
                        
                        .SetForeColor RGB(140, 140, 120)
                        .DrawText 25, 245, "JOYSTICK", False
                        
                        .SetForeColor IIf(4 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                        .DrawText 25, 265, "STATE", False
                        .DrawText 70, 265, IIf(Parent.Joystick, "ENABLED", "DISABLED"), False
                        
                        .SetForeColor IIf(5 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                        .DrawText 25, 280, "ZONE ", False
                        .DrawBox 70, 283, 123, 292
                        For L_nRun = 0 To Parent.JoystickDead \ 4
                            .DrawLine 70 + L_nRun * 2, 285, 70 + L_nRun * 2, 290
                        Next
                        
                    Case 3
                    
                        Parent.Frames.Item("NLSLOGO").RenderToFrame Parent.Frames.Item("IDISPLAY"), 20, 275, (Parent.FrameCount Mod 72) \ 2
                        .SetForeColor RGB(140, 140, 120)
                        
                        .SetForeColor RGB(140, 140, 120)
                        .DrawText 25, 150, "Zooom on the Web", False
                        .DrawText 25, 170, "     www.inode.at", False
                        .DrawText 25, 185, "      /nls/zooom", False
                        .DrawText 25, 205, "  levels*ships*fun", False
                        
                        .SetForeColor RGB(140, 140, 120)
                        .DrawText 25, 260, "Zooom realized by", False
                        .DrawText 65, 280, "NONLINEAR", False
                        .DrawText 65, 295, "SOLUTIONS", False
                        .DrawText 25, 315, "mail:  nls@inode.at", False
                    
                End Select
                
            Case 6
            
                .SetForeColor RGB(240, 240, 200)
                .DrawText 25, 150, "GAME OVER", False
                
                .SetForeColor IIf(0 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 280, "RESTART", False
                .SetForeColor IIf(1 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 295, "EXIT", False
            
                DrawStats Parent.Frames.Item("IDISPLAY").Surface

            Case 3
            
                .SetForeColor RGB(240, 240, 200)
                .DrawText 25, 150, "LEVEL COMPLETE", False
                
                .SetForeColor IIf(0 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 280, "CONTINUE", False
                .SetForeColor IIf(1 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 295, "EXIT", False
            
                DrawStats Parent.Frames.Item("IDISPLAY").Surface
            
            Case 4
            
                .SetForeColor RGB(240, 240, 200)
                .DrawText 25, 150, "CAMPAIGN DONE", False
                
                .SetForeColor IIf(0 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 280, "CONTINUE", False
                .SetForeColor IIf(1 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 295, "EXIT", False
            
                DrawStats Parent.Frames.Item("IDISPLAY").Surface
                
            Case 5
            
                .SetForeColor RGB(240, 240, 200)
                .DrawText 25, 150, "GAME PAUSED", False
                
                .SetForeColor IIf(0 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 280, "RESUME", False
                .SetForeColor IIf(1 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 295, "RESTART", False
                .SetForeColor IIf(2 = DspItem, RGB(240, 240, 200), RGB(140, 140, 120))
                .DrawText 25, 310, "EXIT", False
            
                DrawStats Parent.Frames.Item("IDISPLAY").Surface
                
        End Select
        
    End With

End Sub

Public Sub DrawStats(oSurface As DirectDrawSurface7)

    With oSurface
    
        .SetForeColor RGB(140, 140, 120)
        .DrawText 40, 180, "SCORE ", False
        .DrawText 80, 180, Format(Parent.Player.Score, "0"), False
        .DrawText 40, 200, "FIRED", False
        .DrawText 80, 200, Format(Parent.Player.Stat_OwnShotsFired, "0"), False
        .DrawText 40, 215, "HIT", False
        .DrawText 80, 215, Format((Parent.Player.Stat_OwnShotsHit * 100) / (Parent.Player.Stat_OwnShotsFired + 1), "0") & "%", False
        .DrawText 40, 230, "FRAGS", False
        .DrawText 80, 230, Format(Parent.Player.Stat_Frags, "0"), False
                
        If Parent.Player.Score >= Parent.Campaign.HighScore Then
            .DrawText 40, 250, "HIGHSCORE!", False
        End If
        
    End With
    
End Sub

Private Sub RenderVideoContents()

    If I_nVidFrame > -1 Then
        I_nVidFrame = I_nVidFrame + 1
        If I_nVidFrame > 287 Then I_nVidFrame = 287
            
        If I_nVidFrame Mod 6 = 0 Then
            Select Case Parent.Mode
                Case 3
                    Parent.Frames.Item("VLEVEL").RenderToFrame Parent.Frames.Item("IVIDEO"), 115, 20, I_nVidFrame \ 8
                Case 4
                    Parent.Frames.Item("VCAMPAIGN").RenderToFrame Parent.Frames.Item("IVIDEO"), 115, 20, I_nVidFrame \ 8
                Case 5
                    Parent.Frames.Item("VPAUSE").RenderToFrame Parent.Frames.Item("IVIDEO"), 115, 20, I_nVidFrame \ 8
                Case 6
                    Parent.Frames.Item("VOVER").RenderToFrame Parent.Frames.Item("IVIDEO"), 115, 20, I_nVidFrame \ 8
            End Select
        End If
    Else
        Parent.Frames.Item("IVIDEOBACK").RenderToFrame Parent.Frames.Item("IVIDEO"), 115, 20, 0
    End If
    
End Sub
