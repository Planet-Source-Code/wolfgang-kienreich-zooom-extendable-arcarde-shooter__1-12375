VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemyTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Parent As cEnemyTemplates
Public Frame As cFrame
Public Hitpoints As Long
Public Animated As Boolean
Public Directed As Boolean
Public Targetting As Boolean
Public Shooting As Boolean
Public Weapon As Long
Public ExplosionSize As Long
Public ScoreValue As Long
Public NoCollide As Boolean
Public NoHit As Boolean
Public PowerUpStyle As Long
Private I_oCollection As Collection

Public Sub AddComponent(sType As String, nOffsetX As Long, nOffsetY As Long)
    Dim L_oItem As New cComponentTemplate
    With L_oItem
        .sName = sType
        Set .Parent = Me
        .OffsetX = nOffsetX
        .OffsetY = nOffsetY
    End With
    I_oCollection.Add L_oItem
    Set L_oItem = Nothing
End Sub

Public Function ComponentCount()
    ComponentCount = I_oCollection.Count
End Function

Public Function Component(nIndex As Long) As cComponentTemplate
    Set Component = I_oCollection.Item(nIndex)
End Function

Private Sub Class_Initialize()
    Set I_oCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set I_oCollection = Nothing
End Sub
