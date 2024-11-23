'Code VBScript
'My Password Provider Example
'A sequence of a set of English characters is randomly generated to match the standard 255 characters. It is saved to the user's directory and then used to convert input phrases.
'autor: alex;

'Change main settings and save this text as file.vbs

'To do: create algoritm pad to min len
'To do: now bad algoritm for uppercase

'not needed
	GenPassword "SGF2ZSBhIGdvb2QgbW9vZCBldmVyeW9uZQ", False, True
'Output:d9Wo2d7819*b?o?8?%wb207~*wu5p%w#2?
'dnNlIGJlc3BvbGV6bm8
';0)

Sub GenPassword(strPhrase, NoMes, boolUseSpecial)'(Optional strPhrase As String = "", Optional NoMes As Boolean = False, Optional boolUseSpecial = True) As String
'On Error GoTo Err_
'GenPassword = ""
'The name of the file My Password Provider settings
Dim strPPSFileName 'As String
    strPPSFileName = "MySettings-2024-11-22.pps"
Dim SettingsUL 'As String
	'This is nice bug-feature
	SettingsUL = "112377" 'Numbers of char in UPPER Case

'The set of characters
Dim strSimpleCharacters
    strSimpleCharacters = "abcdefghijklmnopqrstuvwxyz1234567890"
Dim strSpecialCharacters
    strSpecialCharacters = "%*)?@#$~"
Dim strAllCharacters
    strAllCharacters = strSimpleCharacters + strSpecialCharacters
   
    Dim fso ' As New FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim strUserDirectory
	
	Dim oShell
	Set oShell = CreateObject( "WScript.Shell" )
    strUserDirectory = oShell.ExpandEnvironmentStrings("%USERPROFILE%") 'Environ("USERPROFILE")
    
    strPPSFileName = strUserDirectory + "\" + strPPSFileName
    
    If Not NoMes Then
        strPPSFileName = InputBox("Enter the name of the file MyPasswordProvider settings", "MyPasswordProvider", strPPSFileName)
    End If


    Dim oPPSFile ' As Object
    Dim i ' As Integer
    If Not fso.FileExists(strPPSFileName) Then

        If Not NoMes Then
            If vbNo = MsgBox("File of settings not found. Create file?", vbYesNo, "My Password Provider") Then
                Exit Sub'Function
            End If
        End If
        
        fso.CreateTextFile strPPSFileName
        
        Set oPPSFile = fso.CreateTextFile(strPPSFileName)
        
        Randomize ' Initialize random-number generator. 
        
        Dim Char
        '
        For i = 0 To 255
            Char = " "
            While InStr(1, strSimpleCharacters, Char) = 0
                Char = Chr(255 * Rnd(9))
            Wend
            oPPSFile.Write Char
        Next
        For i = 0 To 255
            Char = " "
            While InStr(1, strAllCharacters, Char) = 0
                Char = Chr(255 * Rnd)
            Wend
            oPPSFile.Write Char
        Next
        
        oPPSFile.Close
   
    End If

'---------------------------------------------------
'!!! here must be code to pad to min len and encode with CryptoAPI
'---------------------------------------------------

    Set oPPSFile = fso.OpenTextFile(strPPSFileName)
   
    strAllCharacters = oPPSFile.ReadAll()
    oPPSFile.Close
    Set oPPSFile = Nothing
    Set fso = Nothing

    'The new set of characters
    Dim strNewSetOfCharacters
    
    If Not NoMes Then
        boolUseSpecial = (vbYes = MsgBox("Use Special Characters?", vbYesNo, "My Password Provider"))
    End If
    
    If boolUseSpecial Then
        strNewSetOfCharacters = Right(strAllCharacters, 256)
    Else
        strNewSetOfCharacters = Left(strAllCharacters, 256)
    End If
    
    If Len(strPhrase) = 0 Then
        strPhrase = InputBox("Enter phrase for password", "My Password Provider")
    End If
    
    Dim strPwd, intPos
    strPwd = ""
    For i = 1 To Len(strPhrase)
        Char = Mid(strNewSetOfCharacters, Asc(Mid(strPhrase, i, 1)), 1)
        If InStr(1, SettingsUL, CStr(i)) Then
            Char = UCase(Char)
        Else
            Char = LCase(Char)
        End If
        strPwd = strPwd + Char
    Next
    
    If Not NoMes Then
        'GenPassword = InputBox("Correct your password", "My Password Provider", strPwd)
		InputBox "Correct your password", "My Password Provider", strPwd
    Else
        'GenPassword = strPwd
    End If
    
'Exit_:
'    Exit Sub 'Function
'Err_:
'    MsgBox Err.Description, Err.Source
'    GoTo Exit_
End Sub
