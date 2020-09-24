VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemies"
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

Public Function Add(sType As String) As cEnemy

    Dim L_oNew As cEnemy
    Dim L_oTpl As cEnemyTemplate
    
    Set L_oNew = New cEnemy
    
    Set L_oNew.Parent = Me
    L_oNew.Initialize
    I_oCollection.Add L_oNew
    Set L_oTpl = Parent.Parent.EnemyTemplates.Item(sType)
    If L_oTpl Is Nothing Then
        Set Add = Nothing
        Exit Function
    End If
    
    With L_oNew
        .Animated = L_oTpl.Animated
        .Directed = L_oTpl.Directed
        .ExplosionSize = L_oTpl.ExplosionSize
        Set .Frame = L_oTpl.Frame
        .Hitpoints = L_oTpl.Hitpoints
        .NoCollide = L_oTpl.NoCollide
        .NoHit = L_oTpl.NoHit
        .IsPowerup = L_oTpl.PowerUpStyle > 0
        .PowerUpStyle = L_oTpl.PowerUpStyle
        .ScoreValue = L_oTpl.ScoreValue
        .Targetting = L_oTpl.Targetting
        .Shooting = L_oTpl.Shooting
        .Weapon = L_oTpl.Weapon
        
        If L_oTpl.ComponentCount > 0 Then
            Dim L_nIndex As Long
            For L_nIndex = 1 To L_oTpl.ComponentCount
                With Add(L_oTpl.Component(L_nIndex).sName)
                    .ComponentX = L_oTpl.Component(L_nIndex).OffsetX
                    .ComponentY = L_oTpl.Component(L_nIndex).OffsetY
                    Set .ComponentParent = L_oNew
                End With
            Next
        End If
        
    End With
    
    Set Add = L_oNew
    Set L_oNew = Nothing

End Function

Public Property Get Item(nIndex As Long) As cEnemy
Attribute Item.VB_UserMemId = 0
    Set Item = I_oCollection(nIndex)
End Property

Public Property Get Count() As Long
    Count = I_oCollection.Count
End Property

Public Sub Remove(nIndex As Long)
    I_oCollection.Remove nIndex
End Sub

Public Sub Clear()
    Do While I_oCollection.Count > 0
        Set I_oCollection.Item(I_oCollection.Count) = Nothing
    Loop
    Set I_oCollection = New Collection
End Sub

Public Sub Update()
    
    Dim L_oItem As cEnemy
    For Each L_oItem In I_oCollection
        L_oItem.Update
        If L_oItem.Terminating Then
            L_oItem.Active = False
        End If
    Next
    
End Sub

Public Sub Render()
    
    Dim L_oItem As cEnemy
    For Each L_oItem In I_oCollection
        If L_oItem.Active Then L_oItem.Render
    Next

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

    Dim L_oItem As cEnemy
    For Each L_oItem In I_oCollection
        L_oItem.Terminate
    Next
    
    Set I_oCollection = Nothing
    
End Sub

