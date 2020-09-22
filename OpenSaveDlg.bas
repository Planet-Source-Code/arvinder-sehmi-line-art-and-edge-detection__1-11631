Attribute VB_Name = "OpenSaveDlg"
'-------------------------------------------------------'
' This Code Was Taken From PSC                          '
' Thanks To: Brand-X Software For The Open_File Sub     '
' Minor Edits By Arvinder Sehmi & Creation Of Save_File '
'-------------------------------------------------------'
'Declare Api Calls
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Declare Types
Public Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
'Declare Variables
Public SaveFileDialog As OPENFILENAME
Public OpenFileDialog As OPENFILENAME
Private rv As Long
Private sv As Long
Public Function Open_File(hWnd As Long) As String
   rv& = GetOpenFileName(OpenFileDialog)
   If (rv&) Then
      Open_File = Trim$(OpenFileDialog.lpstrFile)
   Else
      Open_File = ""
   End If
End Function
Public Function Save_File(hWnd As Long) As String
   sv& = GetSaveFileName(SaveFileDialog)
   If (sv&) Then
      Save_File = Trim$(SaveFileDialog.lpstrFile)
   Else
      Save_File = ""
   End If
End Function
Private Sub InitSaveDlg()
  With SaveFileDialog
     .lStructSize = Len(SaveFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "Bmp Image File" + Chr$(0) + "*.Bmp"
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = App.Path
     .lpstrTitle = "Save Line Art Image..."
     .flags = 0
  End With
End Sub
Private Sub InitOpenDlg()
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "Image Files" + Chr$(0) + "*.bmp;*.jpg;*.pcx;*.gif"
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = App.Path
     .lpstrTitle = "Load Colour Image..."
     .flags = 0
   End With
End Sub
Public Sub InitDlgs()
 Call InitSaveDlg
 Call InitOpenDlg
End Sub
