Attribute VB_Name = "XPModule"
' This Adds support for WinXP Style Controls on WindowsXP. Only Executable file .exe can have XP Style,
'Simply compiling project from VB won't change Style.

'You can add XP Style to ur projects by copying 'ResWinXP.Res' file and this module
'in ur project and setting Start-up object to Sub Main from Project Properties.
'ResWinXP.Res is supposed to be biuld for every project specifically as it contains
'Product name n such but its not necessary (I think!) ,Even this Res is from another project of mine ;)

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
  'call InitCommonControls before we can use XP visual styles.
  On Error Resume Next
  Dim iccex As tagInitCommonControlsEx
  With iccex
      .lngSize = LenB(iccex)
      .lngICC = ICC_USEREX_CLASSES
  End With
  InitCommonControlsEx iccex
  
  'Now to Program...
  OGLWin.Show
End Sub

