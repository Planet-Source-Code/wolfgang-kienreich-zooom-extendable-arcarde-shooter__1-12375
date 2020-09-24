VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCampaign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cSession
Public CurrentLevel As Long
Private I_sName As String
Private I_sLevelName() As String
Private I_nExtraLife() As Long
Private I_nNextExtraLife As Long
Public PlayerHitpoints As Long
Public PlayerLifes As Long
Public HighScore As Long
Public HighStars As Long

Public Function NextExtraLife() As Long
    NextExtraLife = I_nNextExtraLife
End Function

Public Sub Initialize()

    ReDim I_sLevelName(0)
    ReDim I_nExtraLife(0)

End Sub

Public Sub Clear()
    ReDim I_sLevelName(0)
    ReDim I_nExtraLife(0)
    CurrentLevel = 0
    HighScore = 0
    HighStars = 0
End Sub

Public Sub Load(sName As String)
    Clear
    
    Dim L_dParsingResults As tParsingResults
    Dim L_sLine As String
    
    Open App.Path + "\level\" + sName + ".cpn" For Input As #3
    
    Do
        Input #3, L_sLine
        Let L_dParsingResults = Parse(L_sLine)
    Loop Until L_dParsingResults.sCommand = "BEGIN"
    
    Do

        Input #3, L_sLine
        Let L_dParsingResults = Parse(L_sLine)
        
        Select Case L_dParsingResults.sCommand
            
            Case "LEVEL"
                ReDim Preserve I_sLevelName(UBound(I_sLevelName) + 1)
                I_sLevelName(UBound(I_sLevelName)) = L_dParsingResults.sArgument(0)
            Case "LIFE"
                I_nExtraLife(UBound(I_nExtraLife)) = L_dParsingResults.sArgument(0)
                ReDim Preserve I_nExtraLife(UBound(I_nExtraLife) + 1)
                
            Case "PLAYER"
            
               Case "HITPOINTS"
                     PlayerHitpoints = Val(L_dParsingResults.sArgument(0))
                
                Case "LIFES"
                     PlayerLifes = Val(L_dParsingResults.sArgument(0))
                
            Case "END"
                Exit Do
            
        End Select
        
    Loop Until EOF(3)
    Close #3
    
    On Error GoTo E_NoScoreFile
    Open App.Path + "\level\" + sName + ".sco" For Input As #7
    Input #7, HighScore, HighStars
    Close #7
    
E_NoScoreFile:
        
    I_sName = sName
    I_nNextExtraLife = I_nExtraLife(0)
    
    CurrentLevel = 1
    
End Sub

Public Sub WriteHigh()
    Open App.Path + "\level\" + I_sName + ".sco" For Output As #7
    Write #7, HighScore, HighStars
    Close #7
End Sub

Public Sub CheckExtraLife(nScore As Long)

    Dim L_nRun As Long
    For L_nRun = 0 To UBound(I_nExtraLife)
        
        If L_nRun = 0 Then
            If I_nExtraLife(L_nRun) > Parent.Player.Score Then
                I_nNextExtraLife = I_nExtraLife(L_nRun)
            End If
        End If
        
        If L_nRun > 0 Then
            If I_nExtraLife(L_nRun) > Parent.Player.Score Then
                If I_nExtraLife(L_nRun - 1) < Parent.Player.Score Then
                    I_nNextExtraLife = I_nExtraLife(L_nRun)
                End If
            End If
        End If
        
        If Parent.Player.Score < I_nExtraLife(L_nRun) And Parent.Player.Score + nScore >= I_nExtraLife(L_nRun) Then
            If Parent.Player.Lifes < 9 Then Parent.Player.Lifes = Parent.Player.Lifes + 1
            Parent.Level.Effects.Add "EXPLOP", Parent.Player.PositionX, Parent.Player.PositionY, 0, IIf(Parent.Level.Stopper Is Nothing, 5, 0)
            Parent.Level.Message = "extra ship awarded"
        End If
        
    Next
    
End Sub

Public Function LevelCount()

    LevelCount = UBound(I_sLevelName())
    
End Function

Public Function LevelName(nIndex As Long)
    LevelName = I_sLevelName(nIndex)
End Function

