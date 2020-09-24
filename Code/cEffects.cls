VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEffects"
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

Public Function Add(sType As String, nPositionX As Long, nPositionY As Long, nDeltaX As Long, nDeltaY As Long, Optional oParent As cEnemy) As cEffect

    
    Dim L_oNew As cEffect
    Set L_oNew = New cEffect
    Set L_oNew.Parent = Me
    I_oCollection.Add L_oNew
    
    L_oNew.PositionX = nPositionX
    L_oNew.PositionY = nPositionY
    L_oNew.DeltaX = nDeltaX
    L_oNew.DeltaY = nDeltaY
    Set L_oNew.ParentObject = oParent
    
    Select Case sType
    
        Case "SMOKESHIP"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("SMOKE")
            
        Case "EXPLO1"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("EXPLO1")
            
        Case "EXPLO2"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("EXPLO2")
            
        Case "EXPLO3"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("EXPLO3")
            
        Case "EXPLO4"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("EXPLO4")
            
        Case "EXPLOP"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("EXPLOP")
            
        Case "HITP1"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("HITP1")
            
        Case "HITP2"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("HITP2")
            
        Case "HITP3"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("HITP3")
            
        Case "HITP4"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("HITP4")
            
        Case "HITP5"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("HITP5")
            
        Case "SHIPHIT"
            Set L_oNew.Frame = Parent.Parent.Frames.Item("EXPLO1")
            
    End Select

    If Parent.Parent.SoundPresent Then
        If Not L_oNew.Frame.Sound Is Nothing Then
            If L_oNew.Frame.SoundDelay < timeGetTime Then
                L_oNew.Frame.SoundDelay = timeGetTime + 250
                Parent.Parent.Sounds.Add Parent.Parent.DSInstance.DuplicateSoundBuffer(L_oNew.Frame.Sound)
            End If
        End If
    End If
    
    Set Add = L_oNew

    Set L_oNew = Nothing
    
End Function

Public Property Get Item(nIndex As Long) As cEffect
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

    Dim L_oItem As cEffect
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

    Dim L_oItem As cEffect
    For Each L_oItem In I_oCollection
        L_oItem.Render
    Next

End Sub


Public Sub Clear()
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

