VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Member0" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cLevel
Private I_oCollection As Collection

Public Function Hit(ByVal nLeft As Long, ByVal nTop As Long, ByVal nRight As Long, ByVal nBottom As Long, ByVal bPlayer As Boolean) As cShot
    
    Dim L_oItem As cShot
     
    For Each L_oItem In I_oCollection
        If L_oItem.PlayerShot = bPlayer Then
            With L_oItem
                If .PositionX \ 10 > nLeft And .PositionY \ 10 > nTop And .PositionX \ 10 < nRight And .PositionY \ 10 < nBottom Then
                    Set Hit = L_oItem
                    Exit For
                End If
            End With
        End If
    Next
    
End Function

Public Function Add(sType As String, ByVal nPositionX As Long, ByVal nPositionY As Long, ByVal nTargetX As Long, ByVal nTargetY As Long, Optional ByVal bPlayer As Boolean) As cShot

    Dim L_nLength As Long
    Dim L_nSpeed As Long
    
    Dim L_oNew As cShot
    
    Set L_oNew = New cShot
    
    Set L_oNew.Parent = Me
    I_oCollection.Add L_oNew
    
    L_oNew.PositionX = nPositionX
    L_oNew.PositionY = nPositionY
    
    nTargetX = nTargetX - nPositionX
    nTargetY = nTargetY - nPositionY
    
    L_nLength = Sqr(nTargetX ^ 2 + nTargetY ^ 2)

    Select Case sType
    
        Case "ENEMYSHOT1"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("ENEMYSHOT1")
            L_nSpeed = 25
            L_oNew.Power = 40
            
        Case "ENEMYSHOT2"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("ENEMYSHOT2")
            L_nSpeed = 25
            L_oNew.Power = 60
            
        Case "ENEMYSHOT3"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("ENEMYSHOT3")
            L_nSpeed = 30
            L_oNew.Power = 80
            
        Case "PLAYER1"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("PLAYERSHOT1")
            L_nSpeed = 30
            L_oNew.Power = 5
            L_oNew.ExploStyle = "HITP1"
            
        Case "PLAYER2"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("PLAYERSHOT2")
            L_nSpeed = 25
            L_oNew.Power = 10
            L_oNew.ExploStyle = "HITP2"
                        
        Case "PLAYER3"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("PLAYERSHOT3")
            L_nSpeed = 20
            L_oNew.Power = 15
            L_oNew.ExploStyle = "HITP3"
                        
        Case "PLAYER4"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("PLAYERSHOT4")
            L_nSpeed = 15
            L_oNew.Power = 40
            L_oNew.ExploStyle = "HITP4"
            
            
        Case "PLAYER5"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("PLAYERSHOT5")
            L_nSpeed = 30
            L_oNew.Power = 25
            L_oNew.ExploStyle = "HITP5"
            
    End Select
    
    L_oNew.DeltaX = Int((nTargetX / L_nLength) * L_nSpeed)
    L_oNew.DeltaY = Int((nTargetY / L_nLength) * L_nSpeed)
    L_oNew.PlayerShot = bPlayer
    
    If Parent.Parent.SoundPresent Then
        If Not L_oNew.Frame.Sound Is Nothing Then
            If L_oNew.Frame.SoundDelay < timeGetTime Then
                L_oNew.Frame.SoundDelay = timeGetTime + 100
                Parent.Parent.Sounds.Add Parent.Parent.DSInstance.DuplicateSoundBuffer(L_oNew.Frame.Sound)
            End If
        End If
    End If
    
    Set Add = L_oNew
    
    Set L_oNew = Nothing

End Function

Public Property Get Item(nIndex As Long) As cShot
Attribute Item.VB_UserMemId = 0
    Set Item = I_oCollection(nIndex)
End Property

Public Property Get Count() As Long
    Count = I_oCollection.Count
End Property

Public Sub Remove(nIndex As Long)
    I_oCollection.Remove nIndex
End Sub

Public Sub Update()

    Dim L_oItem As cShot
    For Each L_oItem In I_oCollection
        L_oItem.Update
    Next
    
    Dim L_nIndex As Long
    L_nIndex = 1
    Do Until L_nIndex > I_oCollection.Count
        If I_oCollection.Item(L_nIndex).Terminating Then
            I_oCollection.Remove L_nIndex
        Else
            L_nIndex = L_nIndex + 1
        End If
    Loop
    
End Sub

Public Sub Render()

    Dim L_oItem As cShot
    For Each L_oItem In I_oCollection
        L_oItem.Render
    Next

End Sub


Public Sub Clear()
    Do While I_oCollection.Count > 0
        I_oCollection.Remove I_oCollection.Count
    Loop
    Set I_oCollection = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = I_oCollection.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set I_oCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set I_oCollection = Nothing
End Sub

