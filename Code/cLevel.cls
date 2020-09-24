VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cSession
Public Position As Long
Public Stopper As cEnemy
Public Music As String

Private I_sMessage As String
Private I_nMessageCount As String

Private I_nBackgroundPosition As Long
Private I_nBackgroundLength As Long
Private I_nLevelLength As Long
Private I_oDDSurfaceBackground As DirectDrawSurface7

Private I_oEnemies As cEnemies
Private I_oShots As cShots
Private I_oEffects As cEffects

Public Function MessageActive() As Boolean
    MessageActive = (I_nMessageCount > 0)
End Function
Public Property Get Message() As String
    Message = I_sMessage
End Property
Public Property Let Message(sMessage As String)
    I_sMessage = sMessage
    I_nMessageCount = 150
End Property


Public Property Get Effects() As cEffects
    Set Effects = I_oEffects
End Property
Public Property Set Effects(oEffects As cEffects)
    Set I_oEffects = oEffects
End Property

Public Property Get Enemies() As cEnemies
    Set Enemies = I_oEnemies
End Property
Public Property Set Enemies(oEnemies As cEnemies)
    Set I_oEnemies = oEnemies
End Property

Public Property Get Shots() As cShots
    Set Shots = I_oShots
End Property
Public Property Set Shots(oShots As cShots)
    Set I_oShots = oShots
End Property

Public Property Get Background() As DirectDrawSurface7
    Set Background = I_oDDSurfaceBackground
End Property
Public Property Set Background(oBackground As DirectDrawSurface7)
    Set I_oDDSurfaceBackground = oBackground
End Property

Public Property Get BackgroundPosition() As Long
    BackgroundPosition = I_nBackgroundPosition
End Property

Public Property Get LevelLength() As Long
    LevelLength = I_nLevelLength
End Property

Public Sub Reset()
    
    Dim L_oEnemy As cEnemy
    For Each L_oEnemy In I_oEnemies
        L_oEnemy.Active = False
        L_oEnemy.Terminating = False
        L_oEnemy.Waypoints.Reset
        L_oEnemy.LifeCount = 0
    Next
    I_oShots.Clear
    I_oEffects.Clear
    
    Position = 0
    Set Stopper = Nothing
    If Not Music = "" Then
        Parent.PlayMidiMusic Music
        Parent.SetMusicVolume 50
    End If
    Message = "get ready ..."
    
End Sub
Public Sub Initialize(sName As String)
    
    Dim L_dParsingResults As tParsingResults
    Dim L_sLine As String
    Dim L_nPosition As Long
    Dim L_nWaypoint As Long
    Dim L_nLineCount As Long
    
    Dim L_dDDSD As DDSURFACEDESC2
        
    ReDim I_sPlugin(0)
    
    Set I_oEnemies = New cEnemies
    Set I_oEnemies.Parent = Me
    
    Set I_oShots = New cShots
    Set I_oShots.Parent = Me
    
    Set I_oEffects = New cEffects
    Set I_oEffects.Parent = Me
    
    Dim L_oEnemy As cEnemy
    Dim L_oParent As cEnemy
    
    Dim L_oWaypoint As cWaypoint
    
    Open App.Path + "\level\" + sName + ".lvl" For Input As #1
    
    Do
        Input #1, L_sLine
        L_nLineCount = L_nLineCount + 1
        Let L_dParsingResults = Parse(L_sLine)
        
    Loop Until L_dParsingResults.sCommand = "BEGIN"
    
    On Error GoTo E_LineError
    
    Do

        Input #1, L_sLine
        L_nLineCount = L_nLineCount + 1
        Let L_dParsingResults = Parse(L_sLine)
        
        Select Case L_dParsingResults.sCommand
            
            Case "PLUGIN"
                Call LoadPlugin(L_dParsingResults.sArgument(0))
                
            Case "MUSIC"
                Me.Music = L_dParsingResults.sArgument(0)
                
            Case "BACKGROUND"
            
                L_dDDSD.lFlags = DDSD_CAPS
                L_dDDSD.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
                
                On Error GoTo E_IsPluginBackground
                Open App.Path + "\" + L_dParsingResults.sArgument(0) + ".bmp" For Input As #3
                Close #3
                Set I_oDDSurfaceBackground = Parent.DDInstance.CreateSurfaceFromFile(App.Path + "\" + L_dParsingResults.sArgument(0) + ".bmp", L_dDDSD)
                GoTo E_IsStandardBackground
E_IsPluginBackground:
                On Error GoTo E_LineError
                Set I_oDDSurfaceBackground = Parent.DDInstance.CreateSurfaceFromFile(App.Path + "\plugin\" + L_dParsingResults.sArgument(0) + ".bmp", L_dDDSD)
E_IsStandardBackground:
                On Error GoTo E_LineError
                I_oDDSurfaceBackground.GetSurfaceDesc L_dDDSD
                I_nBackgroundLength = L_dDDSD.lHeight - 500
                
            Case "LENGTH"
                I_nLevelLength = Val(L_dParsingResults.sArgument(0))
                
            Case "AT"
                L_nPosition = Val(L_dParsingResults.sArgument(0))
                
            Case "OFFSET"
                L_nPosition = L_nPosition + Val(L_dParsingResults.sArgument(0))
                
            Case "POWERUP"
            
                Set L_oEnemy = I_oEnemies.Add(L_dParsingResults.sArgument(0))
                If L_oEnemy Is Nothing Then
                    Err.Raise 13
                End If
                L_oEnemy.TriggerAt = L_nPosition
                L_oEnemy.IsPowerup = True
                L_oEnemy.NoHit = True
            
            Case "SHIP"
            
                Set L_oEnemy = I_oEnemies.Add(L_dParsingResults.sArgument(0))
                If L_oEnemy Is Nothing Then
                    Err.Raise 13
                End If
                
                L_oEnemy.TriggerAt = L_nPosition
                Set L_oParent = L_oEnemy
                
                Case "DIRECTION"
                    
                    If L_dParsingResults.sArgument(0) = "PLAYER" Then
                        L_oWaypoint.Direction = 1000
                    Else
                        L_oWaypoint.Direction = Val(L_dParsingResults.sArgument(0))
                    End If
                    
                Case "CIRCLE"
                    L_oWaypoint.Style = 1
                    L_oWaypoint.CircleDir = IIf(L_dParsingResults.sArgument(0) = "CLOCKWISE", 1, -1)
                    L_oWaypoint.CircleRad = Val(L_dParsingResults.sArgument(1))
                    
                Case "SPEED"
                    L_oWaypoint.Speed = Val(L_dParsingResults.sArgument(0))
                                    
                Case "STOP"
                    L_oWaypoint.StopLevel = True
                                    
                Case "RESUME"
                    L_oWaypoint.ResumeLevel = True
                                    
                Case "ATTACK"
                    L_oWaypoint.AttackPropability = Val(L_dParsingResults.sArgument(0)) + Parent.Difficulty * 25
                    If L_oWaypoint.AttackPropability > 100 Then L_oWaypoint.AttackPropability = 100
                    
                Case "LOOP"
                    L_oWaypoint.LoopTo = Val(L_dParsingResults.sArgument(0))
                                    
                Case "WAYPOINT"
                    Set L_oWaypoint = L_oEnemy.Waypoints.Add(IIf(L_dParsingResults.sArgument(0) = "VERTICAL", "V", "H"), Val(L_dParsingResults.sArgument(1)))
                    If L_oEnemy.Waypoints.Count = 1 Then L_oEnemy.Waypoints.Advance
                    
            Case "END"
                Exit Do
            
        End Select
        
        L_dParsingResults.sCommand = ""
        
    Loop Until EOF(1)
    Close #1
    On Error GoTo 0
    
    If Not Music = "" Then
        Parent.PlayMidiMusic Music
    End If
    
    fZooom.FontSize = 8
    Message = "get ready ..."
    Exit Sub
    
E_LineError:
    fMessage.Hide
    fMessage.lblMessage.Caption = "Error reading level " & sName & vbCrLf & "Syntax error in line " & L_nLineCount & " (Command " & IIf(L_dParsingResults.sCommand = "", "unknown", L_dParsingResults.sCommand) & ")" & vbCrLf & vbCrLf & "Please check syntax and structure of this level. When designing custom levels, please stick tightly to the command reference provided with Zooom. If the error persists, feel free to contact the author."
    fMessage.Show 1
    Terminate
    Parent.Terminate
    End
        
End Sub

Public Sub Update()

    If (Stopper Is Nothing) And (Parent.Player.PositionY < 5000) Then Position = Position + 1
    I_nBackgroundPosition = I_nBackgroundLength - Position * (I_nBackgroundLength / I_nLevelLength)
    If I_nBackgroundPosition < 0 Then I_nBackgroundPosition = 0
    
    If I_nMessageCount > 0 Then I_nMessageCount = I_nMessageCount - 1
    
    
    I_oEnemies.Update
    I_oShots.Update
    Parent.Player.Update
    I_oEffects.Update
    
End Sub

Public Sub Terminate()
        
    mciSendString "close " & Music, 0&, 0, 0
    
    Set I_oEnemies = Nothing
    If Not I_oShots Is Nothing Then I_oShots.Clear
    Set I_oShots = Nothing
    Set I_oEffects = Nothing
    
    Set I_oDDSurfaceBackground = Nothing
    
End Sub

Public Sub LoadPlugin(sName As String)
        
    If Parent.PluginLoaded(sName) Then Exit Sub
    Parent.RegisterPlugin sName
    
    Dim L_oET As cEnemyTemplate
    
    Dim L_dParsingResults As tParsingResults
    Dim L_sLine As String
    Dim L_nPosition As Long
    Dim L_nLineCount As Long
    
    Open App.Path + "\plugin\" + sName + ".plg" For Input As #2
    
    Do
        Input #2, L_sLine
        L_nLineCount = L_nLineCount + 1
        Let L_dParsingResults = Parse(L_sLine)
    Loop Until L_dParsingResults.sCommand = "BEGIN"
    
    Do

        Input #2, L_sLine
        L_nLineCount = L_nLineCount + 1
        Let L_dParsingResults = Parse(L_sLine)
        
        Select Case L_dParsingResults.sCommand
            
            Case "FRAME"
                If L_oET Is Nothing Then
                    Parent.Frames.Add L_dParsingResults.sArgument(0), App.Path + "\plugin\" + L_dParsingResults.sArgument(1) + ".bmp", "", Val(L_dParsingResults.sArgument(2)), Val(L_dParsingResults.sArgument(3)), Val(L_dParsingResults.sArgument(4))
                Else
                    Set L_oET.Frame = Parent.Frames.Item(L_dParsingResults.sArgument(0))
                    If L_oET Is Nothing Then Err.Raise 13

                End If

            Case "SHIP"
                Set L_oET = Parent.EnemyTemplates.Add(L_dParsingResults.sArgument(0))
                
            Case "HITPOINTS"
                L_oET.Hitpoints = Val(L_dParsingResults.sArgument(0))
            
            Case "SCORE"
                L_oET.ScoreValue = Val(L_dParsingResults.sArgument(0))
                    
            Case "ANIMATED"
                L_oET.Animated = True
            
            Case "DIRECTED"
                L_oET.Directed = True
            
            Case "TARGETTING"
                L_oET.Targetting = True
                
            Case "SHOOTING"
                L_oET.Shooting = True
                
            Case "EXPLOSION"
                L_oET.ExplosionSize = Val(L_dParsingResults.sArgument(0))
                
            Case "WEAPON"
                L_oET.Weapon = Val(L_dParsingResults.sArgument(0))
                
            Case "HITS"
                If L_dParsingResults.sArgument(0) = "OFF" Then L_oET.NoHit = True
                
            Case "COLLISSIONS"
                If L_dParsingResults.sArgument(0) = "OFF" Then L_oET.NoCollide = True
                
            Case "COMPONENT"
                L_oET.AddComponent L_dParsingResults.sArgument(0), Val(L_dParsingResults.sArgument(1)), Val(L_dParsingResults.sArgument(2))
                
            Case "END"
                Exit Do
            
        End Select
        
    Loop Until EOF(2)
    Close #2
    
E_PlgError:
    fMessage.Hide
    fMessage.lblMessage.Caption = "Error reading plugin " & sName & vbCrLf & "Syntax error in line " & L_nLineCount & " (Command " & IIf(L_dParsingResults.sCommand = "", "unknown", L_dParsingResults.sCommand) & ")" & vbCrLf & vbCrLf & "Please check syntax and structure of this level. When designing custom levels, please stick tightly to the command reference provided with Zooom. If the error persists, feel free to contact the author."
    fMessage.Show 1
    Terminate
    Parent.Terminate
    End
    
End Sub

