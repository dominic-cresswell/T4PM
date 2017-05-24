Attribute VB_Name = "S_UserConfigCode"
Option Private Module


'----------------

Private Function CheckUserConfig() As Boolean
        
  CheckUserConfig = False
   ProgramPath$ = S_UserConfigCode.CheckProgramPath

' check the ...
    If FileExists(ProgramPath$ & "UserConfigFile") = False Then
            CheckUserConfig = False
    Else
            CheckUserConfig = True
    End If
     
    
End Function
    
    
Public Function CheckProgramPath() As String

' = = =
     ConfigFile$ = Environ("appdata")
        
    If Right(ConfigFile$, 1) <> "\" Then ConfigFile$ = ConfigFile$ & "\"

' .. set up the sub-folder
    If DirExists(ConfigFile$ & "T4PM") = False Then
        MkDir ConfigFile$ & "T4PM"
    End If
    
    CheckProgramPath = ConfigFile$ & "T4PM"
    If Right(CheckProgramPath, 1) <> "\" Then CheckProgramPath = CheckProgramPath & "\"
    
End Function



Sub CreateUserConfig(dummy$)

        DefaultConfig$ = ""
        DefaultConfig$ = DefaultConfig$ & "WorkingPath=" & Environ("userprofile") & "\" & vbCrLf
        DefaultConfig$ = DefaultConfig$ & "RememberLastProject=False" & vbCrLf
        DefaultConfig$ = DefaultConfig$ & "" & vbCrLf
     
   Call MakeTextFile(DefaultConfig$, ProgramPath$ & "UserConfigFile")


End Sub


Sub SetConfigSetting(inOption$, inParam$)

    If ProgramPath$ = "" Then ProgramPath$ = S_UserConfigCode.CheckProgramPath
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
 
    ' check for old
    
    If GetConfigSetting(inOption$) <> "" Then
        OldSetting$ = inOption$ + "=" + GetConfigSetting(inOption$)
        
        GetTxtData$ = Replace(GetTxtData$, OldSetting$, inOption$ + "=" + inParam$)
        

    Else
        GetTxtData$ = GetTxtData$ & inOption$ + "=" + inParam$ & vbCrLf
    End If
    
    
    Call MakeTextFile(GetTxtData$, ProgramPath$ & "UserConfigFile")
    

End Sub


Function GetConfigSetting(inOption$)
    
    ' ....
If ProgramPath$ = "" Then ProgramPath$ = S_UserConfigCode.CheckProgramPath
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
 
    GetTxtData$ = ReadTextFile(ProgramPath$ & "UserConfigFile")
  
       GetPoint = InStr(vbTextCompare, LCase(GetTxtData$), LCase(inOption$) & "=")
    If GetPoint = 0 Then Exit Function
    
    GetConfigSetting = Right(GetTxtData$, Len(GetTxtData$) - GetPoint - Len(inOption$))
    GetPoint = InStr(vbTextCompare, GetConfigSetting, vbCrLf)
    
    
'====
    If GetPoint > 0 Then
        GetConfigSetting = Left(GetConfigSetting, GetPoint - 1)
    End If
        
    
End Function


