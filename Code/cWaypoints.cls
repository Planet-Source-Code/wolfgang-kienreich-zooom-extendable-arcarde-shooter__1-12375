VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWaypoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Member0" ,"cFrameTemplate"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Parent As cEnemy
Private I_nCurrentItem As Long
Private I_oCollection As Collection

Public Sub Reset()
    I_nCurrentItem = 1
End Sub
Public Function Advance()

    If Not ((I_nCurrentItem < 1) Or (I_nCurrentItem > I_oCollection.Count)) Then
        If I_oCollection.Item(I_nCurrentItem).LoopTo > 0 Then
            I_nCurrentItem = I_oCollection.Item(I_nCurrentItem).LoopTo + 1
        Else
            I_nCurrentItem = I_nCurrentItem + 1
        End If
    Else
        If I_nCurrentItem < I_oCollection.Count Then I_nCurrentItem = I_nCurrentItem + 1
    End If
    
End Function

Public Function Current() As cWaypoint
    If I_nCurrentItem > I_oCollection.Count Then
        Set Current = Nothing
    Else
        Set Current = I_oCollection.Item(I_nCurrentItem)
    End If
End Function

Public Function Add(sOrientation As String, nPosition As Long) As cWaypoint

    Dim L_oNew As cWaypoint
    Set L_oNew = New cWaypoint
    
    Set L_oNew.Parent = Me
    L_oNew.Position = nPosition
    L_oNew.Orientation = sOrientation
    
    I_oCollection.Add L_oNew

    Set Add = L_oNew
    
    Set L_oNew = Nothing

End Function

Public Property Get Item(nIndex As Long) As cWaypoint
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

