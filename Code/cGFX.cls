VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Member0" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cSession
Private I_oCollection As Collection

Public Function Add(sName As String) As cGFX

    Dim L_oNew As cGFX
    Set L_oNew = New cGFX
    
    Set L_oNew.Parent = Me
    
    I_oCollection.Add L_oNew, sName

    Set Add = L_oNew
    
    Set L_oNew = Nothing

End Function

Public Property Get Item(sName As String) As cGFX
Attribute Item.VB_UserMemId = 0
    Set Item = I_oCollection(sName)
End Property

Public Property Get Count() As Long
    Count = I_oCollection.Count
End Property

Public Sub Remove(sName As String)
    I_oCollection.Remove sName
End Sub

Public Sub Clear()
    Do While I_oCollection.Count > 0
        Set I_oCollection.Item(I_oCollection.Count) = Nothing
    Loop
    Set I_oCollection = New Collection
End Sub

Public Sub Initialize()
    
    Dim I_oItem As cGFX
    
    With Add("SHIP01")
        .Initialize "shiptest.bmp", 8, 4
    End With
    
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

    Dim L_oItem As cGFX
    For Each L_oItem In I_oCollection
        L_oItem.Terminate
    Next

    Set I_oCollection = Nothing
    
End Sub

