VERSION 5.00
Begin VB.Form form1 
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1995
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   1995
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' coded by Jonahdrummer@hotpop.com
' www.jonsworldonline.com
' you may change any part of the code, all that is asked is if used give credit please, And Vote for me.

Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Private Sub Form_load()
Dim save1
Dim Text3
Dim Text4
form1.Visible = "false" 'hides the actual window from the user
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "cjr", "c:\windows\cjr.exe")  'make cjr start @ startup
On Error Resume Next
Open "C:\cjrcaption.txt" For Append As #1
Close #1
Open "C:\cjrcaption.txt" For Output As #1 'creates the caption file
Text3 = "Your Company Name"
Write #1, Text3
Close #1
Open "C:\cjrcaption.txt" For Input As #1 'input the information from the caption and text file for our message box
Input #1, Text3
Close #1
Open "C:\cjrtext.txt" For Append As #1
Close #1
Open "C:\cjrtext.txt" For Output As #1 'creates the text file
Text4 = "Message to User"
Write #1, Text4
Close #1
Open "C:\cjrtext.txt" For Input As #1
Input #1, Text4
Close #1
Open "C:\windows\system\cjmsgr.win" For Append As #1
Close #1
Open "C:\windows\system\cjmsgr.win" For Input As #1
Input #1, save1
Close #1
Text1.Text = Text3
Text2.Text = Text4
FileSystem.FileCopy App.Path & "\" & "cjr.exe", "C:\windows\cjr.exe" 'copy the exe to windows so it will restart at reboot and clean up our mess
FileSystem.FileCopy App.Path & "cjr.exe", "C:\windows\cjr.exe" 'used incase program ran from the root directory.
If Not save1 = "1" Then
Open "C:\windows\system\cjmsgr.win" For Output As #1 'create a mark file so it will only run once
save1 = "1"
Write #1, save1
Close #1
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeCaption", Text1.Text)
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeText", Text2.Text)
Else 'Now to clean up our mess
FileSystem.Kill "C:\cjrcaption.txt" 'delete the caption file
FileSystem.Kill "C:\cjrtext.txt" 'delete the text file
FileSystem.Kill "C:\windows\system\cjmsgr.win" 'delete our mark file
FileSystem.Kill "C:\windows\cjr.exe" 'delete the exe from the windows directory
Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeCaption") 'delete both registry strings
Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Winlogon", "LegalNoticeText")
Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "cjr") 'delete cjr from the runing at start
End If
End
End Sub
