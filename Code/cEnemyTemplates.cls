VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnemyTemplates"
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

Public Sub Initialize()

    With Add("DEFAULT_BLOCKER1")
        Set .Frame = Parent.Frames("SHIP01")
        .Hitpoints = 25
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ExplosionSize = 1
        .ScoreValue = 300
    End With
        
    With Add("DEFAULT_BLOCKER2")
        Set .Frame = Parent.Frames("SHIP02")
        .Hitpoints = 50
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ExplosionSize = 1
        .ScoreValue = 500
    End With
    
    With Add("DEFAULT_BLOCKER3")
        Set .Frame = Parent.Frames("SHIP03")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 7
        .ExplosionSize = 2
        .ScoreValue = 1200
    End With
    
    With Add("DEFAULT_GUNSHIP1")
        Set .Frame = Parent.Frames("SHIP04")
        .Hitpoints = 140
        .Animated = False
        .Directed = False
        .Targetting = False
        .Shooting = True
        .Weapon = 8
        .ExplosionSize = 1
        .ScoreValue = 2000
    End With
    
    With Add("DEFAULT_GUNSHIP2")
        Set .Frame = Parent.Frames("SHIP05")
        .Hitpoints = 200
        .Animated = False
        .Directed = False
        .Targetting = False
        .Shooting = True
        .Weapon = 9
        .ExplosionSize = 2
        .ScoreValue = 3000
    End With
    
    With Add("DEFAULT_INTERCEPTOR1")
        Set .Frame = Parent.Frames("SHIP06")
        .Hitpoints = 120
        .Animated = False
        .Directed = True
        .Targetting = False
        .Weapon = 3
        .ExplosionSize = 1
        .ScoreValue = 4500
    End With
    
    With Add("DEFAULT_INTERCEPTOR2")
        Set .Frame = Parent.Frames("SHIP07")
        .Hitpoints = 220
        .Animated = False
        .Directed = True
        .Targetting = False
        .Weapon = 6
        .ExplosionSize = 2
        .ScoreValue = 6000
    End With
    
    With Add("DEFAULT_CAPITAL1")
        Set .Frame = Parent.Frames("SHIP08")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ExplosionSize = 0
        .ScoreValue = 0
        .NoCollide = True
        .NoHit = True
        .AddComponent "DEFAULT_TURRET", 0, 19
    End With
        
    With Add("DEFAULT_CAPITAL2")
        Set .Frame = Parent.Frames("SHIP09")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ExplosionSize = 0
        .ScoreValue = 0
        .NoCollide = True
        .NoHit = True
        .AddComponent "DEFAULT_TURRET", 0, -5
        .AddComponent "DEFAULT_TURRET", 0, 42
    End With
    
    With Add("DEFAULT_CAPITAL3")
        Set .Frame = Parent.Frames("SHIP10")
        .Hitpoints = 100
        .Animated = False
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ExplosionSize = 0
        .ScoreValue = 0
        .NoCollide = True
        .NoHit = True
        .AddComponent "DEFAULT_TURRET", -39, -25
        .AddComponent "DEFAULT_TURRET", 40, -25
        .AddComponent "DEFAULT_TURRET", -39, 19
        .AddComponent "DEFAULT_TURRET", 40, 19
    End With
        
    With Add("DEFAULT_BOSS")
        Set .Frame = Parent.Frames("SHIP11")
        .Hitpoints = 5000
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 10
        .ExplosionSize = 4
        .ScoreValue = 25000
        .AddComponent "DEFAULT_TURRET", -61, 42
        .AddComponent "DEFAULT_TURRET", 60, 42
    End With
      
    With Add("DEFAULT_TURRET")
        Set .Frame = Parent.Frames("SHIP0")
        .Hitpoints = 280
        .Animated = False
        .Directed = False
        .Targetting = True
        .Weapon = 2
        .ExplosionSize = 2
        .ScoreValue = 15000
    End With
    
    With Add("TWINMOUNT")
        Set .Frame = Parent.Frames("UPTWINMOUNT")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 1
    End With
    
    With Add("TRIPPLEMOUNT")
        Set .Frame = Parent.Frames("UPTRIPPLEMOUNT")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 2
    End With
    
    With Add("WEAPON")
        Set .Frame = Parent.Frames("UPWEAPON")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 3
    End With
    
    With Add("SHIELD")
        Set .Frame = Parent.Frames("UPSHIELD")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 4
    End With
    
    With Add("REPAIR")
        Set .Frame = Parent.Frames("UPREPAIR")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 5
    End With
    
    With Add("DRONE")
        Set .Frame = Parent.Frames("UPORBITER")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 6
    End With
        
    With Add("LIFE")
        Set .Frame = Parent.Frames("UPLIFE")
        .Hitpoints = 100
        .Animated = True
        .Directed = False
        .Targetting = False
        .NoHit = True
        .PowerUpStyle = 7
    End With
    
    With Add("DEFAULT_ROCK1")
        Set .Frame = Parent.Frames("ROCK1")
        .Hitpoints = 300
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ScoreValue = 500
    End With

    With Add("DEFAULT_ROCK2")
        Set .Frame = Parent.Frames("ROCK2")
        .Hitpoints = 500
        .Animated = True
        .Directed = False
        .Targetting = False
        .Weapon = 0
        .ScoreValue = 800
    End With

End Sub

Public Function Add(sName As String) As cEnemyTemplate

    On Error GoTo E_AlreadyThere
    Dim L_oNew As cEnemyTemplate
    Set L_oNew = New cEnemyTemplate
    Set L_oNew.Parent = Me
    I_oCollection.Add L_oNew, sName
    Set Add = L_oNew
    Set L_oNew = Nothing
E_AlreadyThere:

End Function

Public Property Get Item(sName As String) As cEnemyTemplate
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

