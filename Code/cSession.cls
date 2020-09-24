VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private I_oDXInstance As DirectX7
Private I_oDDInstance As DirectDraw7
Private I_oDSInstance As DirectSound

Public RenderTime As Long
Public UpdateTime As Long
Public ProcessTime As Long

Public Campaign As New cCampaign

Private I_sPlugin() As String
Public SoundPresent As Boolean

Private I_oViewport As cViewport
Private I_oLevel As cLevel
Private I_oMenu As cMenu
Private I_oFrames As cFrames
Private I_oSounds As cSounds
Private I_oEnemyTemplates As cEnemyTemplates
Private I_oPlayer As cPlayer

Private I_nWaveVolume As Integer
Private I_nMusicVolume As Integer

Public Detail As Boolean
Public Joystick As Boolean
Public JoystickDead As Long
Public Difficulty As Integer

Public Terminating As Boolean
Public Mode As Integer
Public DebugMode As Boolean

Public FrameCount As Long
Public FramesPerSecond As Long

Public Property Get Player() As cPlayer
    Set Player = I_oPlayer
End Property
Public Property Set Player(oPlayer As cPlayer)
    Set I_oPlayer = oPlayer
End Property
Public Sub RegisterPlugin(sName As String)
    ReDim Preserve I_sPlugin(UBound(I_sPlugin) + 1)
    I_sPlugin(UBound(I_sPlugin)) = sName
End Sub
Public Function PluginLoaded(sName As String) As Boolean

    Dim L_nRun As Long
    For L_nRun = 0 To UBound(I_sPlugin)
        If I_sPlugin(L_nRun) = sName Then
            PluginLoaded = True
            Exit Function
        End If
    Next
    
End Function

Public Property Get WaveVolume() As Integer
    WaveVolume = I_nWaveVolume
End Property
Public Property Let WaveVolume(nNew As Integer)
    If nNew > 100 Then nNew = 100
    If nNew < 0 Then nNew = 0
    SetWaveVolume nNew
    I_nWaveVolume = nNew
End Property

Public Property Get MusicVolume() As Integer
    MusicVolume = I_nMusicVolume
End Property
Public Property Let MusicVolume(nNew As Integer)
    If nNew > 100 Then nNew = 100
    If nNew < 0 Then nNew = 0
    SetMusicVolume nNew
    I_nMusicVolume = nNew
End Property

Public Sub SetMusicVolume(nVolume As Integer)
    On Error GoTo E_SetMusicVolume
    If SoundPresent Then
        If Not I_oMenu.MenuSound Is Nothing Then
            I_oMenu.MenuSound.SetVolume Int((DSBVOLUME_MIN / 100) * (100 - IIf(nVolume + 10 > 100, 100, nVolume + 10)))
        End If
    End If
    
    Dim L_nMidiDeviceHandle As Long
    Dim L_nVolume As Long
    L_nMidiDeviceHandle = mciGetDeviceID(I_oLevel.Music)
    L_nVolume = Int(327 * nVolume) + Int(327 * nVolume) * (2 ^ 16)
    midiOutSetVolume L_nMidiDeviceHandle, L_nVolume
    On Error GoTo 0
    Exit Sub
E_SetMusicVolume:
    
End Sub

Public Sub PlayMidiMusic(sFileName As String)

    Dim L_sFullPath As String * 255
    Dim L_sRealPath As String
    Dim L_nPathLength As String
    
    On Error GoTo E_IsPluginMusic
    Open App.Path + "\" + sFileName + ".mid" For Input As #4
    Close #4
    L_nPathLength = GetShortPathName(App.Path, L_sFullPath, 255)
    L_sRealPath = Left(L_sFullPath, L_nPathLength)
    GoTo E_IsStandardMusic
    
E_IsPluginMusic:
    L_nPathLength = GetShortPathName(App.Path + "\plugin\", L_sFullPath, 255)
    L_sRealPath = Left(L_sFullPath, L_nPathLength)
        
E_IsStandardMusic:
    On Error GoTo 0
    
    mciSendString "close " & sFileName, 0&, 0, 0
    mciSendString "open " + L_sRealPath + "\" + sFileName + ".mid type sequencer alias " & sFileName, 0&, 0, 0
    mciSendString "play " & sFileName, 0&, 0, 0

End Sub

Public Sub SetWaveVolume(nVolume As Integer)
    If Not SoundPresent Then Exit Sub
    Dim L_oFrame As cFrame
    For Each L_oFrame In I_oFrames
        If Not L_oFrame.Sound Is Nothing Then
            L_oFrame.Sound.SetVolume Int((DSBVOLUME_MIN / 100) * (100 - nVolume))
        End If
    Next
    
End Sub

Public Property Get DXInstance() As DirectX7
    Set DXInstance = I_oDXInstance
End Property
Public Property Set DXInstance(oDXInstance As DirectX7)
    Set I_oDXInstance = oDXInstance
End Property

Public Property Get DDInstance() As DirectDraw7
    Set DDInstance = I_oDDInstance
End Property
Public Property Set DDInstance(oDDInstance As DirectDraw7)
    Set I_oDDInstance = oDDInstance
End Property

Public Property Get DSInstance() As DirectSound
    Set DSInstance = I_oDSInstance
End Property
Public Property Set DSInstance(oDSInstance As DirectSound)
    Set I_oDSInstance = oDSInstance
End Property

Public Property Get EnemyTemplates() As cEnemyTemplates
    Set EnemyTemplates = I_oEnemyTemplates
End Property
Public Property Set EnemyTemplates(oEnemyTemplates As cEnemyTemplates)
    Set I_oEnemyTemplates = oEnemyTemplates
End Property

Public Property Get Viewport() As cViewport
    Set Viewport = I_oViewport
End Property
Public Property Set Viewport(oViewport As cViewport)
    Set I_oViewport = oViewport
End Property

Public Property Get Level() As cLevel
    Set Level = I_oLevel
End Property
Public Property Set Level(oLevel As cLevel)
    Set I_oLevel = oLevel
End Property

Public Property Get Sounds() As cSounds
    Set Sounds = I_oSounds
End Property
Public Property Set Sounds(oSounds As cSounds)
    Set I_oSounds = oSounds
End Property


Public Property Get Menu() As cMenu
    Set Menu = I_oMenu
End Property
Public Property Set Menu(oMenu As cMenu)
    Set I_oMenu = oMenu
End Property

Public Property Get Frames() As cFrames
    Set Frames = I_oFrames
End Property
Public Property Set Frames(oFrames As cFrames)
    Set I_oFrames = oFrames
End Property

Public Sub Initialize(nHWnd As Long)

    If Not HardwareCheck Then
        fMessage.Hide
        fMessage.lblMessage = "Wrong screen resolution and/or color depth detected. Zooom requires at least a resolution of 800x600 pixel and 16-bit color depth. Please change your system settings and restart Zooom."
        fMessage.Show 1
        Unload fZooom
        End
    End If
    
    Dim L_nValue As Long
    Dim L_nFactor As Long
    For L_nValue = 0 To 255
        G_nTranslucencyLookup(L_nValue, 100) = L_nValue
        For L_nFactor = 0 To 99
            G_nTranslucencyLookup(L_nValue, L_nFactor) = Int(L_nValue * ((100 - L_nFactor) / 100))
        Next
    Next
    
    ReDim I_sPlugin(0)
    Difficulty = 1
    
    On Error GoTo E_DirectX
    Set I_oDXInstance = New DirectX7
    Set I_oDDInstance = I_oDXInstance.DirectDrawCreate("")
    On Error GoTo 0
    
    On Error GoTo E_NoSound
    Set I_oDSInstance = I_oDXInstance.DirectSoundCreate("")
    I_oDSInstance.SetCooperativeLevel fZooom.hwnd, DSSCL_NORMAL
    SoundPresent = True
E_NoSound:
    On Error GoTo 0
    
    Set Campaign = New cCampaign
    Set Campaign.Parent = Me
    Campaign.Initialize
    
    Set I_oViewport = New cViewport
    Set I_oViewport.Parent = Me
    I_oViewport.Initialize nHWnd
    
    Set I_oSounds = New cSounds
    Set I_oSounds.Parent = Me
    
    Set I_oFrames = New cFrames
    Set I_oFrames.Parent = Me
    I_oFrames.Initialize
    
    Set I_oEnemyTemplates = New cEnemyTemplates
    Set I_oEnemyTemplates.Parent = Me
    I_oEnemyTemplates.Initialize
    
    WaveVolume = 80
    MusicVolume = 80
    
    Set I_oMenu = New cMenu
    Set I_oMenu.Parent = Me
    I_oMenu.Initialize
    
    Detail = True
    JoystickDead = 50
    Joystick = False
    Call LoadSettings
    WaveVolume = I_nWaveVolume
    MusicVolume = I_nMusicVolume
    Exit Sub
    
E_DirectX:

        fMessage.Hide
        fMessage.lblMessage = "DirectX could not be initialized. Zooom requires Microsoft DirectX 7.x to be installed on your system. You can download the latest DirectX drivers under www.microsoft.com/directx."
        fMessage.Show 1
        Unload fZooom
        End

End Sub

Public Sub Execute()

    Dim L_nNextFrameTime As Long
    Dim L_nNextSecondTime As Long
    Dim L_nFrameCount As Long
    Dim L_nCurrentTime As Long
    
    fZooom.Refresh

    Do
        If (Mode = 1) And Not (GetForegroundWindow = fZooom.hwnd) Then
            Mode = 5
            I_oMenu.DspOpen = True
            I_oMenu.VidOpen = True
            If SoundPresent Then Sounds.Add DSInstance.DuplicateSoundBuffer(Frames.Item("IDISPLAY").Sound)
        End If
        
        FrameCount = FrameCount + 1
        
        L_nNextFrameTime = timeGetTime + 10
        
        L_nFrameCount = L_nFrameCount + 1
        If timeGetTime > L_nNextSecondTime Then
            FramesPerSecond = L_nFrameCount
            L_nNextSecondTime = timeGetTime + 1000
            L_nFrameCount = 0
        End If
            
        fZooom.picLamp(0).Visible = (FrameCount Mod 100 > 50)
        fZooom.picLamp(1).Visible = (FrameCount Mod 100 <= 50)
        
        If Mode = 1 Then
            If I_oLevel Is Nothing Then
                Set I_oLevel = New cLevel
                Set I_oLevel.Parent = Me
                I_oLevel.Initialize Campaign.LevelName(Campaign.CurrentLevel)
            Else
                If I_oLevel.Position = I_oLevel.LevelLength Then Mode = 2
            End If
        End If
        
        If Mode = 2 Then
            
            If Campaign.CurrentLevel = Campaign.LevelCount Then
                I_oPlayer.Stars = I_oPlayer.Stars + 1
                Campaign.CurrentLevel = 1
                Mode = 4
                I_oMenu.DspItem = 0
                I_oMenu.DspOpen = True
                I_oMenu.VidOpen = True
            Else
                Campaign.CurrentLevel = Campaign.CurrentLevel + 1
                Mode = 3
                I_oMenu.DspItem = 0
                I_oMenu.DspOpen = True
                I_oMenu.VidOpen = True
            End If
            
        End If
        
        If FrameCount Mod 20 = 0 Then UpdateTime = timeGetTime
        If Not I_oLevel Is Nothing Then
            If Mode = 1 Then I_oLevel.Update
        End If
        I_oMenu.Update
        If SoundPresent Then I_oSounds.Update
        If FrameCount Mod 20 = 0 Then UpdateTime = timeGetTime - UpdateTime
        
        If FrameCount Mod 20 = 0 Then RenderTime = timeGetTime
        I_oViewport.Render
        If FrameCount Mod 20 = 0 Then RenderTime = timeGetTime - RenderTime
        
        If FrameCount Mod 20 = 0 Then ProcessTime = timeGetTime
        Process
        If FrameCount Mod 20 = 0 Then ProcessTime = timeGetTime - ProcessTime
        
        If FrameCount Mod 10 = 0 Then DoEvents
        
        Do
        Loop Until timeGetTime >= L_nNextFrameTime
        
    Loop Until Terminating

End Sub

Public Sub Terminate()

    Set I_oDSInstance = Nothing
    
    I_oViewport.Terminate
    Set I_oViewport = Nothing
    
    If Not (I_oLevel Is Nothing) Then I_oLevel.Terminate
    Set I_oLevel = Nothing
    
    If Not (I_oMenu Is Nothing) Then I_oMenu.Terminate
    Set I_oMenu = Nothing
    
    Set I_oEnemyTemplates = Nothing
    Set I_oFrames = Nothing
    Set I_oSounds = Nothing
    
    Set I_oDSInstance = Nothing
    Set I_oDDInstance = Nothing
    Set I_oDXInstance = Nothing
    
    Call SaveSettings
    
End Sub

Public Sub Process()

    If Not (GetForegroundWindow = fZooom.hwnd) Then Exit Sub
    
    Static L_nKeyDelay As Long
    
    Dim L_bKeyL As Boolean
    Dim L_bKeyU As Boolean
    Dim L_bKeyD As Boolean
    Dim L_bKeyR As Boolean
    Dim L_bKeyS As Boolean
    Dim L_bKeyE As Boolean
    Dim L_bKeyC As Boolean
    
    Dim L_dJoy As JOYINFOEX
    
    L_dJoy.dwSize = Len(L_dJoy)
    L_dJoy.dwFlags = 131

    L_bKeyS = (GetAsyncKeyState(&H20) = -32767 - IIf(Mode = 1, 1, 0))
    L_bKeyE = (GetAsyncKeyState(&H1B) = -32767 - IIf(Mode = 1, 1, 0))
    L_bKeyU = (GetAsyncKeyState(&H26) = -32767 - IIf(Mode = 1, 1, 0))
    L_bKeyD = (GetAsyncKeyState(&H28) = -32767 - IIf(Mode = 1, 1, 0))
    L_bKeyL = (GetAsyncKeyState(&H25) = -32767 - IIf(Mode = 1, 1, 0))
    L_bKeyR = (GetAsyncKeyState(&H27) = -32767 - IIf(Mode = 1, 1, 0))
    L_bKeyC = (GetAsyncKeyState(&H11) = -32768)
    
    If Joystick Then
        If joyGetPosEx(0, L_dJoy) = 0 Then
            L_bKeyS = L_bKeyS Or (L_dJoy.dwButtonNumber = 1)
            L_bKeyL = L_bKeyL Or (L_dJoy.dwXpos < 32767& - 300& * IIf(Mode = 1, 0, JoystickDead))
            L_bKeyR = L_bKeyR Or (L_dJoy.dwXpos > 32767& + 300& * IIf(Mode = 1, 0, JoystickDead))
            L_bKeyU = L_bKeyU Or (L_dJoy.dwYpos < 32767& - 300& * IIf(Mode = 1, 0, JoystickDead))
            L_bKeyD = L_bKeyD Or (L_dJoy.dwYpos > 32767& + 300& * IIf(Mode = 1, 0, JoystickDead))
        End If
    End If
    
    If L_nKeyDelay > 0 Then L_nKeyDelay = L_nKeyDelay - 1
    
    If L_nKeyDelay = 0 Then
    
        If Mode = 1 Then
        
            If Not I_oLevel Is Nothing Then
                If Player.Active Then
                
                    If L_bKeyC Then
                        If (GetAsyncKeyState(&H31) = -32767) Then Player.SetEquipment IIf(Player.Weapon < 4, Player.Weapon + 1, Player.Weapon)
                        If (GetAsyncKeyState(&H32) = -32767) Then Player.SetEquipment , IIf(Player.Mounts < 4, Player.Mounts + 1, Player.Mounts)
                        If (GetAsyncKeyState(&H33) = -32767) Then Player.SetEquipment , , IIf(Player.Orbiter < 4, Player.Orbiter + 1, Player.Orbiter)
                        If (GetAsyncKeyState(&H34) = -32767) Then Player.Hitpoints = Player.MaxHitpoints
                        If (GetAsyncKeyState(&H35) = -32767) Then Player.Shield = Player.MaxHitpoints
                        If (GetAsyncKeyState(&H36) = -32767) Then Player.Lifes = IIf(Player.Lifes < 9, Player.Lifes + 1, Player.Lifes)
                        If (GetAsyncKeyState(&H37) = -32767) Then Player.Hitpoints = -1
                        If (GetAsyncKeyState(&H38) = -32767) Then I_oLevel.Position = I_oLevel.LevelLength - 1
                        If (GetAsyncKeyState(&H39) = -32767) Then DebugMode = Not DebugMode
                        If (GetAsyncKeyState(&H30) = -32767) Then I_oLevel.Position = IIf(I_oLevel.Position < I_oLevel.LevelLength - 500, I_oLevel.Position + 500, I_oLevel.Position)
                        
                        L_nKeyDelay = 10
                    End If
                    
                    If L_bKeyL Then
                        Player.DeltaX = Player.DeltaX - 1
                        If Player.DeltaX < -10 Then Player.DeltaX = -10
                        L_nKeyDelay = 0
                    End If
                    
                    If L_bKeyR Then
                        Player.DeltaX = Player.DeltaX + 1
                        If Player.DeltaX > 10 Then Player.DeltaX = 10
                        L_nKeyDelay = 0
                    End If
                    
                    If L_bKeyU Then
                        Player.DeltaY = Player.DeltaY - 1
                        If Player.DeltaY < -10 Then Player.DeltaY = -10
                        L_nKeyDelay = 0
                    End If
                    
                    If L_bKeyD Then
                        Player.DeltaY = Player.DeltaY + 1
                        If Player.DeltaY > 10 Then Player.DeltaY = 10
                        L_nKeyDelay = 0
                    End If
                    
                    If L_bKeyS Then
                        Player.Shoot
                        L_nKeyDelay = 0
                    End If
                
                    If L_bKeyE Then
                        Mode = 5
                        I_oMenu.DspOpen = True
                        I_oMenu.VidOpen = True
                        If SoundPresent Then Sounds.Add DSInstance.DuplicateSoundBuffer(Frames.Item("IDISPLAY").Sound)
                        L_nKeyDelay = 0
                    End If
                
                End If
            End If
                
        Else
        
            If L_bKeyL Then
                I_oMenu.Execute 1
                L_nKeyDelay = 10
            End If
            
            If L_bKeyR Then
                I_oMenu.Execute 2
                L_nKeyDelay = 10
            End If
            
            If L_bKeyU Then
                I_oMenu.Execute 3
                L_nKeyDelay = 10
            End If
            
            If L_bKeyD Then
                I_oMenu.Execute 4
                L_nKeyDelay = 10
            End If
            
            If L_bKeyS Then
                I_oMenu.Execute 0
                L_nKeyDelay = 10
            End If
        
            If L_bKeyE Then
                I_oMenu.Execute 5
                L_nKeyDelay = 10
            End If
            
        End If
        
    End If
    
End Sub

Private Function HardwareCheck() As Boolean

    HardwareCheck = True
    If GetDeviceCaps(fZooom.hdc, 8) < 800 Then HardwareCheck = False
    If GetDeviceCaps(fZooom.hdc, 10) < 600 Then HardwareCheck = False
    If GetDeviceCaps(fZooom.hdc, 24) > -1 Then HardwareCheck = False

End Function

Private Sub LoadSettings()

    If Dir(App.Path + "\zooom.cfg") <> "" Then
        Open App.Path + "\zooom.cfg" For Input As #8
        Input #8, Detail, Difficulty, Joystick, JoystickDead, I_nMusicVolume, I_nWaveVolume
        Close #8
    End If
        
End Sub

Private Sub SaveSettings()

    Open App.Path + "\zooom.cfg" For Output As #8
    Write #8, Me.Detail, Me.Difficulty, Me.Joystick, Me.JoystickDead, Me.MusicVolume, Me.WaveVolume
    Close #8
    
End Sub
