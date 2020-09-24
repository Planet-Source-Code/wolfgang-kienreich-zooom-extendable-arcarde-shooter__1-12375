VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFrames"
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

Public Function Add(sName As String, sImageFile As String, sSoundFile As String, nTileWidth As Long, nTileHeight As Long, nAnimationDelay As Long) As cFrame

    Dim L_dDSBufferDesc As DSBUFFERDESC
    Dim L_dWFMDesc As WAVEFORMATEX
    Dim L_oNew As cFrame
    
    On Error Resume Next
    Set L_oNew = I_oCollection.Item(sName)
    On Error GoTo 0
    
    If (L_oNew Is Nothing) Then
        
        With L_dDSBufferDesc
            .lBufferBytes = 0
            .lFlags = DSBCAPS_CTRLVOLUME
        End With
        
        Set L_oNew = New cFrame
        Set L_oNew.Parent = Me
        
        With L_oNew
            
            .Initialize sImageFile, nTileWidth, nTileHeight
            .AnimationDelay = nAnimationDelay
                    
            If Not sSoundFile = "" Then
                If Parent.SoundPresent Then
                    L_dWFMDesc = GetWaveFileFormat(sSoundFile)
                    Set .Sound = Parent.DSInstance.CreateSoundBufferFromFile(App.Path + "\" + sSoundFile, L_dDSBufferDesc, L_dWFMDesc)
                End If
            End If
            
        End With
        
        I_oCollection.Add L_oNew, sName
        Set Add = L_oNew
        Set L_oNew = Nothing
        
    Else
    
        Set Add = I_oCollection.Item(sName)
        
    End If

End Function

Public Property Get Item(sName As String) As cFrame
Attribute Item.VB_UserMemId = 0
    On Error GoTo E_NotFound
    Set Item = I_oCollection(sName)
    Exit Property
E_NotFound:
    Set Item = Nothing
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
    
    Add "PLAYER01", App.Path + "\player01.bmp", "", 11, 3, 1
    Add "PLAYER02", App.Path + "\player02.bmp", "", 11, 2, 1
    Add "PLAYER03", App.Path + "\player03.bmp", "", 11, 2, 1
    Add "PLAYER04", App.Path + "\player04.bmp", "", 11, 1, 1
    Add "PLAYER05A", App.Path + "\player05a.bmp", "", 11, 1, 1
    Add "PLAYER05B", App.Path + "\player05b.bmp", "", 11, 2, 1
    Add "ORBITER", App.Path + "\player0c.bmp", "", 1, 1, 1

    Add "SHIP01", App.Path + "\ship01.bmp", "", 6, 3, 5
    Add "SHIP02", App.Path + "\ship02.bmp", "", 6, 3, 5
    Add "SHIP03", App.Path + "\ship03.bmp", "", 6, 2, 5
    Add "SHIP04", App.Path + "\ship04.bmp", "", 5, 2, 3
    Add "SHIP05", App.Path + "\ship05.bmp", "", 5, 2, 3
    Add "SHIP06", App.Path + "\ship06.bmp", "", 6, 6, 5
    Add "SHIP07", App.Path + "\ship07.bmp", "", 6, 6, 5
    Add "SHIP08", App.Path + "\ship08.bmp", "", 6, 1, 10
    Add "SHIP09", App.Path + "\ship09.bmp", "", 6, 1, 10
    Add "SHIP10", App.Path + "\ship10.bmp", "", 1, 1, 30
    Add "SHIP11", App.Path + "\ship11.bmp", "", 3, 2, 3
    Add "SHIP0", App.Path + "\ship00.bmp", "", 6, 6, 8
    Add "ROCK1", App.Path + "\shiprock1.bmp", "", 6, 6, 8
    Add "ROCK2", App.Path + "\shiprock2.bmp", "", 6, 6, 8
    
    Add "EXPLO1", App.Path + "\explo1.bmp", "explo1.wav", 4, 3, 5
    Add "EXPLO2", App.Path + "\explo2.bmp", "explo2.wav", 4, 3, 10
    Add "EXPLO3", App.Path + "\explo3.bmp", "explo3.wav", 4, 3, 5
    Add "EXPLO4", App.Path + "\explo4.bmp", "explo4.wav", 8, 1, 4
    Add "EXPLOP", App.Path + "\explop.bmp", "explop.wav", 8, 1, 3
    Add "SMOKE", App.Path + "\smoke.bmp", "", 4, 2, 5
    
    Add "PLAYERSHOT1", App.Path + "\shotp1.bmp", "shotp1.wav", 1, 1, 1
    Add "PLAYERSHOT2", App.Path + "\shotp2.bmp", "shotp2.wav", 1, 1, 1
    Add "PLAYERSHOT3", App.Path + "\shotp3.bmp", "shotp3.wav", 16, 1, 1
    Add "PLAYERSHOT4", App.Path + "\shotp4.bmp", "shotp4.wav", 1, 1, 1
    Add "PLAYERSHOT5", App.Path + "\shotp5.bmp", "shotp5.wav", 8, 1, 5
    
    Add "ENEMYSHOT1", App.Path + "\shote1.bmp", "shote.wav", 8, 1, 5
    Add "ENEMYSHOT2", App.Path + "\shote2.bmp", "shote.wav", 2, 1, 5
    Add "ENEMYSHOT3", App.Path + "\shote3.bmp", "shote.wav", 8, 1, 3
    
    Add "HITP1", App.Path + "\hitp1.bmp", "hitp.wav", 5, 1, 5
    Add "HITP2", App.Path + "\hitp2.bmp", "hitp.wav", 8, 1, 3
    Add "HITP3", App.Path + "\hitp3.bmp", "hitp.wav", 9, 1, 3
    Add "HITP4", App.Path + "\hitp4.bmp", "hitp.wav", 1, 8, 3
    Add "HITP5", App.Path + "\hitp5.bmp", "hitp.wav", 8, 1, 3
    
    Add "UPTWINMOUNT", App.Path + "\up01.bmp", "", 9, 1, 5
    Add "UPTRIPPLEMOUNT", App.Path + "\up02.bmp", "", 9, 1, 5
    Add "UPWEAPON", App.Path + "\up03.bmp", "", 9, 2, 5
    Add "UPSHIELD", App.Path + "\up04.bmp", "", 9, 2, 5
    Add "UPREPAIR", App.Path + "\up05.bmp", "", 9, 1, 5
    Add "UPORBITER", App.Path + "\up06.bmp", "", 9, 2, 5
    Add "UPLIFE", App.Path + "\up07.bmp", "", 9, 2, 5
    
    Add "HITMETER", App.Path + "\hitmeter.bmp", "", 11, 1, 1
    Add "SHIELDMETER", App.Path + "\shieldmeter.bmp", "", 14, 1, 1
    Add "METERBACK", App.Path + "\meterback.bmp", "", 1, 1, 1
    Add "METER", App.Path + "\meterback.bmp", "", 1, 1, 1
    Add "LIFES", App.Path + "\lifes.bmp", "", 10, 1, 1
    
    Add "IBUTTONS", App.Path + "\ibuttons.bmp", "ibuttons.wav", 1, 8, 1
    Add "IDISPLAY", App.Path + "\idisplay.bmp", "idisplay.wav", 1, 1, 1
    Add "IVIDEO", App.Path + "\ivideo.bmp", "idisplay.wav", 1, 1, 1
    Add "IDISPLAYBACK", App.Path + "\idisplay.bmp", "ichoice.wav", 1, 1, 1
    Add "IVIDEOBACK", App.Path + "\ivideoback.bmp", "", 1, 1, 1
    
    Add "NLSLOGO", App.Path + "\nlslogo.bmp", "", 6, 6, 8
    Add "IMENUBACK", App.Path + "\imenuback.bmp", "", 1, 1, 1
    Add "FANFARE", "", "fanfare.wav", 1, 1, 1
    
    Add "VPAUSE", App.Path + "\vpause.bmp", "", 6, 6, 20
    Add "VLEVEL", App.Path + "\vlevel.bmp", "", 6, 6, 20
    Add "VOVER", App.Path + "\vover.bmp", "", 6, 6, 20
    Add "VCAMPAIGN", App.Path + "\vcampaign.bmp", "", 6, 6, 20
    
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

    Dim L_oItem As cFrame
    For Each L_oItem In I_oCollection
        L_oItem.Terminate
    Next

    Set I_oCollection = Nothing
    
End Sub

