Attribute VB_Name = "mDemo"
Option Explicit

Public Const MAX_PATH = 260
Public Const c1000 = 1000

Public Buffer       As String * MAX_PATH
Public Lang         As Long
Public SoundOn      As Boolean
'----------------------------------------------

'----------------------------------------------
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Public Function MakeLong(ByVal Low, ByVal High) As Long
    'This function combines 2 values into a long.
    'Assumes components are < 32767
    MakeLong = High * &H10000 + Low
End Function
Public Sub HoverSound()
    Const SYNC = 1
    If SoundOn Then
      sndPlaySound ByVal App.Path & "\Hover.wav", SYNC
    End If
End Sub
Public Sub SetControlCaptionStrings(frm As Form)

   On Error GoTo ProcedureError

   Dim ctl  As Control

   '-- set the form's caption
   If frm.Tag <> "" Then
      frm.Caption = GetResourceString(CInt(frm.Tag))
   End If
   '-- set the font
   '-- Set fnt = frm.Font
   '-- fnt.Name = GetResourceString(20)
   '-- fnt.Size = CInt(GetResourceString(21))

   '-- set the controls' captions using the Tag property
   For Each ctl In frm.Controls
      If ctl.Tag <> "" Then
         Select Case TypeName(ctl)
            Case "Menu", "Label", "CheckBox", "OptionButton", "ButtonEx"
               ctl.Caption = GetResourceString(Int(ctl.Tag))
         End Select
      End If
   Next

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox("mDeclare.SetControlCaptionStrings") = vbRetry Then Resume Next

End Sub

Public Function GetResourceStringFromFile(sModule As String, idString) As String

   Dim hModule As Long
   Dim nChars As Long
   Dim FreeLib As Boolean
   
   ' is module already mapped into this process.
   hModule = GetModuleHandle(sModule)
   If hModule = 0 Then ' load module
      hModule = LoadLibrary(sModule)
      FreeLib = True
   End If

   If hModule Then ' get resource idString.
      nChars = LoadString(hModule, idString, Buffer, MAX_PATH)
      If nChars Then
         GetResourceStringFromFile = left$(Buffer, nChars)
      End If
      FreeLibrary hModule
   End If
   
   ' unload library if we loaded it here.
   If FreeLib Then Call FreeLibrary(hModule)

End Function

Public Function GetResourceString(Num As Variant) As String
   On Error Resume Next
   Select Case Num
      Case 1000 To 1999 'Get from resource file (.Res)
         GetResourceString = LoadResString(Lang + Num)
      Case Else
         GetResourceString = GetResourceStringFromFile("Shell32.Dll", Num)
   End Select
End Function

Public Function ErrMsgBox(Msg As String) As Integer
   ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, Msg)
End Function


