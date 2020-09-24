VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSounds"
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

Public Sub Add(oSound As DirectSoundBuffer)
    
    Dim L_oNew As cSound
    Set L_oNew = New cSound
    
    With L_oNew
        Set .Sound = oSound
        .Sound.Play DSBPLAY_DEFAULT
        .StartTime = timeGetTime
    End With
    
    I_oCollection.Add L_oNew
    Set L_oNew = Nothing
    
End Sub

Public Property Get Item(nIndex As Long) As cSound
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
    
    Dim L_nTime As Long
    
    Dim L_nIndex As Long
    L_nIndex = 1
    
    L_nTime = timeGetTime
    
    Do Until L_nIndex > I_oCollection.Count
        If I_oCollection.Item(L_nIndex).StartTime < L_nTime - 1500 Then
            I_oCollection.Remove L_nIndex
        Else
            L_nIndex = L_nIndex + 1
        End If
    Loop
    
End Sub

Public Sub Clear()
    Do While I_oCollection.Count > 0
        Set I_oCollection.Item(I_oCollection.Count) = Nothing
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

