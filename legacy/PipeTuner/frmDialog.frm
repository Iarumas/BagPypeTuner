VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDialog 
   Caption         =   "Dialog"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog gdlg_File 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Function File_Open(strExtension As String) As String

    Dim strDotExtension As String
    Dim strWildCardDotExt As String
        
    strDotExtension = "." & strExtension
    strWildCardDotExt = "*." & strExtension
    
    'open dialog and look for files with the file extension : *.strExt
    With gdlg_File
        .FileName = vbNullString
        .flags = cdlOFNOverwritePrompt
        .Filter = " (" & strWildCardDotExt & ")|" & strWildCardDotExt
        .ShowOpen
    End With
    
    'exit if no file is selected
    If gdlg_File.FileName = vbNullString Then
        MsgBox ("No File Selected !")
        Exit Function
    End If
    
    'exit if the file extension is not ".strExtension" (e.g. ".wav")
    If Right$(LCase$(gdlg_File.FileName), Len(strDotExtension)) <> strDotExtension Then
        MsgBox ("This is no " & strDotExtension & " File !")
        Exit Function
    End If
    
    File_Open = gdlg_File.FileName

End Function

Public Function File_Save(strExtension As String)
    
    Dim strDotExtension As String
    Dim strWildCardDotExt As String
        
    strDotExtension = "." & strExtension
    strWildCardDotExt = "*." & strExtension
    
    'open dialog and for files to save / file extension : *.strExtension
    With gdlg_File
        .FileName = vbNullString
        .flags = cdlOFNOverwritePrompt
        .Filter = " (" & strWildCardDotExt & ")|" & strWildCardDotExt
        .ShowSave
    End With

    'exit if no filename was entered
    If gdlg_File.FileName = vbNullString Then
        MsgBox ("No File Name Entered !")
        Exit Function
    End If
    
    'save file
    'if the file extension is .strExtension this is the correct filename
    ' else the extension has to be added
    If Right$(LCase$(gdlg_File.FileName), Len(strDotExtension)) = strDotExtension Then
            File_Save = gdlg_File.FileName
    Else
            File_Save = gdlg_File.FileName & strDotExtension
    End If
    
End Function

