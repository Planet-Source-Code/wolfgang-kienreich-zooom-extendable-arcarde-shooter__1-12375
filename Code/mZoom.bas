Attribute VB_Name = "mZoom"
Public Const PIFACTOR = 0.01745329

Public Type tParsingResults
    sCommand As String
    sArgument() As String
End Type

Public Type JOYINFOEX
        dwSize As Long
        dwFlags As Long
        dwXpos As Long
        dwYpos As Long
        dwZpos As Long
        dwRpos As Long
        dwUpos As Long
        dwVpos As Long
        dwButtons As Long
        dwButtonNumber As Long
        dwPOV As Long
        dwReserved1 As Long
        dwReserved2 As Long
End Type

Public Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type
    
Public Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByVal pDest As Long, ByVal numBytes As Long, ByVal fillbyte As Byte)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function joyGetPosEx Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFOEX) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public G_nTranslucencyLookup(255, 100) As Byte

Public Function GetWaveFileFormat(ByVal sFileName As String) As WAVEFORMATEX
        
    Dim L_dWFX As WAVEFORMATEX
    Dim L_nPosition As Long
    Dim L_nWaveBytes() As Byte
    
    sFileName = App.Path + "\" + sFileName
    ReDim L_nWaveBytes(1 To FileLen(sFileName))
    Open sFileName For Binary As #1
    Get #1, , L_nWaveBytes
    Close #1
    L_nPosition = 1
    Do While Not (Chr(L_nWaveBytes(L_nPosition)) + Chr(L_nWaveBytes(L_nPosition + 1)) + Chr(L_nWaveBytes(L_nPosition + 2)) = "fmt")
        L_nPosition = L_nPosition + 1
    Loop
    CopyMemory VarPtr(L_dWFX), VarPtr(L_nWaveBytes(L_nPosition + 8)), Len(L_dWFX)
    
End Function

Public Function Parse(sLine As String) As tParsingResults

    Dim L_sComponent() As String
    Dim L_nPos As Integer
    ReDim L_sComponent(0)
    
    sLine = Trim(UCase(sLine))
    
    Do
        L_nPos = 0
        If L_nPos = 0 Then L_nPos = InStr(1, sLine, " ")
        If L_nPos = 0 Then L_nPos = InStr(1, sLine, vbTab)
        
        If L_nPos = 0 Then
        
            L_sComponent(UBound(L_sComponent)) = sLine
            Exit Do
            
        Else
        
            If L_nPos > 1 Then
            
                L_sComponent(UBound(L_sComponent)) = Left(sLine, L_nPos - 1)
                ReDim Preserve L_sComponent(UBound(L_sComponent) + 1)
                
            End If
            
            sLine = Mid(sLine, L_nPos + 1)
        
        End If
        
    Loop
    
    Parse.sCommand = L_sComponent(0)
    If UBound(L_sComponent) > 0 Then
        ReDim Parse.sArgument(UBound(L_sComponent) - 1)
        For L_nPos = 1 To UBound(L_sComponent)
            Parse.sArgument(L_nPos - 1) = L_sComponent(L_nPos)
        Next
    End If
    
End Function
