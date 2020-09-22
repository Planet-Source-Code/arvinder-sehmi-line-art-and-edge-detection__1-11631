VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line Art And Edge Detection"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Load 
      Caption         =   "Load A Colour\Greyscale Image"
      Height          =   420
      Left            =   0
      TabIndex        =   27
      Top             =   3915
      Width           =   3435
   End
   Begin VB.CommandButton CancelDraw 
      Caption         =   "Cancel Draw"
      Height          =   510
      Left            =   6930
      TabIndex        =   22
      Top             =   5490
      Width           =   1860
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save Line Art Image"
      Height          =   420
      Left            =   3465
      TabIndex        =   21
      Top             =   3915
      Width           =   3435
   End
   Begin VB.CommandButton StartEdgeDetect 
      Caption         =   "Edge Detection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6930
      TabIndex        =   18
      Top             =   4950
      Width           =   1860
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options:"
      Height          =   4335
      Left            =   6930
      TabIndex        =   7
      Top             =   45
      Width           =   1815
      Begin VB.CheckBox uRed 
         Caption         =   "Red Values"
         Height          =   330
         Left            =   135
         TabIndex        =   12
         Top             =   495
         Value           =   1  'Checked
         Width           =   1590
      End
      Begin VB.CheckBox uGreen 
         Caption         =   "Green Values"
         Height          =   330
         Left            =   135
         TabIndex        =   11
         Top             =   810
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox uBlue 
         Caption         =   "Blue Values"
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   1125
         Width           =   1590
      End
      Begin VB.VScrollBar Tolerance 
         Height          =   1860
         LargeChange     =   25
         Left            =   1260
         Max             =   255
         TabIndex        =   9
         Top             =   2340
         Value           =   157
         Width           =   285
      End
      Begin VB.CheckBox Invert 
         Caption         =   "Invert Image"
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   1935
         Width           =   1410
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "The Current Tolerance Is:"
         Height          =   465
         Left            =   135
         TabIndex        =   17
         Top             =   3330
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Tolerance: --------------->"
         Height          =   420
         Left            =   180
         TabIndex        =   16
         Top             =   2385
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Use:"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "(When GreyScaling)"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   1485
         Width           =   1545
      End
      Begin VB.Label Tol 
         BackStyle       =   0  'Transparent
         Caption         =   "157"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   450
         TabIndex        =   13
         Top             =   3825
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tips"
      Height          =   1590
      Left            =   0
      TabIndex        =   3
      Top             =   4410
      Width           =   6900
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Me At: Arvinder@Sehmi.org.uk If You Need Help, Or Have A Question."
         ForeColor       =   &H8000000D&
         Height          =   870
         Left            =   4545
         TabIndex        =   33
         Top             =   540
         Width           =   2220
      End
      Begin VB.Label Tmp 
         BackStyle       =   0  'Transparent
         Height          =   510
         Left            =   5490
         TabIndex        =   20
         Top             =   495
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"MainFrm.frx":1472
         ForeColor       =   &H8000000C&
         Height          =   465
         Left            =   45
         TabIndex        =   6
         Top             =   225
         Width           =   6765
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ticking, Red, Green and Blue, Usually Gives a Good Result."
         ForeColor       =   &H80000014&
         Height          =   240
         Left            =   45
         TabIndex        =   5
         Top             =   630
         Width           =   6855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Experiment With All The Controls, And See What You Get."
         ForeColor       =   &H8000000C&
         Height          =   285
         Left            =   45
         TabIndex        =   4
         Top             =   855
         Width           =   6900
      End
      Begin VB.Label Label4 
         Caption         =   "If There Is Too Much Black In The Resulting Image, Then Decrease The Tolerance."
         ForeColor       =   &H80000014&
         Height          =   420
         Left            =   45
         TabIndex        =   32
         Top             =   1080
         Width           =   4200
      End
   End
   Begin VB.CommandButton StartLineArt 
      Caption         =   "Draw Line Art"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6930
      TabIndex        =   2
      Top             =   4410
      Width           =   1860
   End
   Begin VB.Frame Frame2 
      Caption         =   "Line Art Picture"
      Height          =   3840
      Left            =   3465
      TabIndex        =   1
      Top             =   45
      Width           =   3435
      Begin VB.VScrollBar DScrollV 
         Enabled         =   0   'False
         Height          =   3345
         LargeChange     =   500
         Left            =   3195
         SmallChange     =   150
         TabIndex        =   31
         Top             =   225
         Width           =   150
      End
      Begin VB.HScrollBar DScrollH 
         Enabled         =   0   'False
         Height          =   150
         LargeChange     =   500
         Left            =   90
         TabIndex        =   30
         Top             =   3600
         Width           =   3075
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   90
         ScaleHeight     =   3345
         ScaleWidth      =   3075
         TabIndex        =   28
         Top             =   225
         Width           =   3075
         Begin VB.PictureBox Dest 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   3345
            Left            =   0
            ScaleHeight     =   3345
            ScaleWidth      =   3075
            TabIndex        =   29
            Top             =   0
            Width           =   3075
         End
      End
      Begin VB.Label PercentDone 
         Alignment       =   2  'Center
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2655
         TabIndex        =   19
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Original Picture"
      Height          =   3840
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   3435
      Begin VB.HScrollBar SScrollH 
         Enabled         =   0   'False
         Height          =   150
         LargeChange     =   500
         Left            =   90
         TabIndex        =   26
         Top             =   3600
         Width           =   3075
      End
      Begin VB.VScrollBar SScrollV 
         Enabled         =   0   'False
         Height          =   3345
         LargeChange     =   500
         Left            =   3195
         SmallChange     =   150
         TabIndex        =   25
         Top             =   225
         Width           =   150
      End
      Begin VB.PictureBox SourceContainer 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   90
         ScaleHeight     =   3345
         ScaleWidth      =   3075
         TabIndex        =   23
         Top             =   225
         Width           =   3075
         Begin VB.PictureBox Source 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   3330
            Left            =   0
            ScaleHeight     =   3330
            ScaleWidth      =   3075
            TabIndex        =   24
            Top             =   0
            Width           =   3075
         End
      End
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------'
' Line Art Creation, And Edge Dection By Arivnder Sehmi. '
' Arvinder@Sehmi.org.uk                                  '
' September 23th 2000                                    '
'--------------------------------------------------------'
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Dim Cancel As Boolean ' has the cancel button been pressed?
Dim AveCol As Long ' Holds The Grey Colour Of A Pixel
Dim Saving As Boolean 'Is The App Saving A File?
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Do While Saving = True ' wait until the app has finished saving a file.
  DoEvents
 Loop
 Unload Me 'unload
 End       'end
End Sub

Private Sub Save_Click()
 Dim File As String ' holds the file name
 File = Save_File(Me.hWnd) 'show save dlg
 If Trim(File) = "" Then MsgBox "File Not Saved, Invalid Filename.", vbCritical, "Error": Exit Sub ' error in name
 Saving = True ' start saving
 Me.Caption = "Please Wait Saving....."
 Dest.Picture = Dest.Image 'set the picture to equal the image
 Tmp.Caption = File '-- get rid of any unwanted chars (ie chr13, or 0)
 File = Tmp.Caption '/
 If LCase(Right(File, 4) <> ".bmp") Then File = File & ".bmp" ' add the bmp on the file
 Call SavePicture(Dest.Picture, File) ' save the picture
 Saving = False ' no longer saving
 Me.Caption = "Line Art \ Edge Detection By Arvinder Sehmi"
End Sub

Private Sub Load_Click()
 Dim File As String ' holds the file name
 File = Open_File(Me.hWnd) 'show the open file dlg
 If Trim(File) = "" Then Exit Sub ' make sure the file is correct
 Source.Picture = LoadPicture(File) ' load the file
 Dest.Height = Source.Height '--setup sizes
 Dest.Width = Source.Width   '/
 
 Source.Top = 0: Source.Left = 0: Dest.Top = 0: Dest.Left = 0 'Re-set Image Positions
 SScrollV.Value = 0: SScrollH.Value = 0: DScrollV.Value = 0: DScrollH.Value = 0 'Re-set Image Positions
 
 'Re-set Scroll Bars
 If Source.Height <= SourceContainer.Height Then SScrollV.Enabled = False: DScrollV.Enabled = False: GoTo Next1:
 SScrollV.Enabled = True
 SScrollV.Max = Source.Height - SourceContainer.Height
 SScrollV.LargeChange = SScrollV.Max / 5
 DScrollV.Enabled = True
 DScrollV.Max = Source.Height - SourceContainer.Height
 DScrollV.LargeChange = DScrollV.Max / 5
 
Next1:
 
 'Re-set Scroll Bars
 If Source.Width <= SourceContainer.Width Then SScrollH.Enabled = False: DScrollH.Enabled = False: GoTo Next2:
 SScrollH.Enabled = True
 SScrollH.Max = Source.Width - SourceContainer.Width
 SScrollH.LargeChange = SScrollH.Max / 5
 DScrollH.Enabled = True
 DScrollH.Max = Source.Width - SourceContainer.Width
 DScrollH.LargeChange = DScrollH.Max / 5
 
Next2:

End Sub

Private Sub Form_Load()
 InitDlgs 'initalize save and open dialogs
 Dest.Height = Source.Height '--set up sizes
 Dest.Width = Source.Width   '/
End Sub
Sub CancelDraw_Click()
 Cancel = True ' Cancel The Draw
 StartLineArt.Enabled = True    '--Enable The Buttons
 StartEdgeDetect.Enabled = True '/
End Sub
Private Sub SScrollH_Change()    '\
 Source.Left = -(SScrollH.Value) ' \
End Sub                          ' |
Private Sub SScrollV_Change()    ' |
 Source.Top = -(SScrollV.Value)  ' |
End Sub                          ' |\------------Scroll The Image
Private Sub DScrollH_Change()    ' |/
 Dest.Left = -(DScrollH.Value)   ' |
End Sub                          ' |
Private Sub DScrollV_Change()    ' |
 Dest.Top = -(DScrollV.Value)    ' |
End Sub                          '/

Private Sub StartLineArt_Click()
Cancel = False
StartLineArt.Enabled = False
StartEdgeDetect.Enabled = False
Dim Total As Long 'store the total number of pixels in the image
Total = (Source.Width / Screen.TwipsPerPixelX) * (Source.Height / Screen.TwipsPerPixelY) 'get the total number of pixels in the image
Dest.Cls
For x = 0 To Source.Width / Screen.TwipsPerPixelX 'loop through the x-pixels
 For y = 0 To (Source.Height / Screen.TwipsPerPixelY) 'loop through the y-pixels
  AveCol = GreyFromColour(GetPixel(Source.hdc, x, y)) ' get the grey colour for a coloured pixel
  If AveCol > Tolerance.Value Then AveCol = 255 Else AveCol = 0 ' choose if it is to be black or white
  If Invert.Value = 1 Then AveCol = (255 - AveCol) 'invert the colours if nessesarry
  If AveCol = 0 Then SetPixel Dest.hdc, x, y, RGB(AveCol, AveCol, AveCol) ' draw the pixel if it is black
 Next y 'loop through the y-pixels
 DoEvents
 Dest.Refresh 'refresh
 PercentDone.Caption = Int(((x * y) / Total) * 100) & "%" 'calculate the percent done.
 If Cancel = True Then GoTo Finish:
Next x 'loop through the x-pixels
Finish:
StartLineArt.Enabled = True
StartEdgeDetect.Enabled = True
End Sub

Private Function GreyFromColour(LongCol As Long) As Integer
 ' Get The Red, Blue And Green Values Of A Colour From The Long Value
 Dim Blue As Double, Green As Double, Red As Double, GreenS As Double, BlueS As Double, Num As Integer
 Blue = Fix((LongCol / 256) / 256)
 Green = Fix((LongCol - ((Blue * 256) * 256)) / 256)
 Red = Fix(LongCol - ((Blue * 256) * 256) - (Green * 256))
 
 If uRed.Value = 1 Then GreyFromColour = GreyFromColour + Red: Num = Num + 1
 If uGreen.Value = 1 Then GreyFromColour = GreyFromColour + Green: Num = Num + 1
 If uBlue.Value = 1 Then GreyFromColour = GreyFromColour + Blue: Num = Num + 1
 GreyFromColour = GreyFromColour / Num ' average the colours, to get a grey colour.
End Function

Private Sub Tolerance_Change() '\
 Tol.Caption = Tolerance.Value ' \
End Sub                        ' |--Update the Tolerance.
Private Sub Tolerance_Scroll() ' |
 Tol.Caption = Tolerance.Value ' /
End Sub                        '/

Private Sub StartEdgeDetect_Click()
Cancel = False
StartLineArt.Enabled = False
StartEdgeDetect.Enabled = False
Dim Col As Long ' hold the colour of the pixel minus the colour of the pixel next to it
Dim Total As Long
Total = (Source.Width / Screen.TwipsPerPixelX) * (Source.Height / Screen.TwipsPerPixelY)
Dest.Cls
For x = 1 To Source.Width \ Screen.TwipsPerPixelX 'loop through the x-pixels
 For y = 1 To Source.Height \ Screen.TwipsPerPixelY 'loop through the y-pixels
  Col = Abs(GetPixel(Source.hdc, x, y) - GetPixel(Source.hdc, x, y - 1)) ' hold the colour of the pixel minus the colour of the pixel on the top of it
  If Col > (Tolerance.Value) ^ 3 Then Col = vbWhite Else Col = 0 ' choose if the colour is of high contrast
  If Invert.Value = 0 Then Col = (vbWhite - Col) ' check for an invert
  If Col = 0 Then SetPixel Dest.hdc, x, y, Col ' plot pixel
  Col = Abs(GetPixel(Source.hdc, x, y) - GetPixel(Source.hdc, x - 1, y)) ' hold the colour of the pixel minus the colour of the pixel on the left of it
  If Col > (Tolerance.Value) ^ 3 Then Col = vbWhite Else Col = 0 ' choose if the colour is of high contrast
  If Invert.Value = 0 Then Col = (vbWhite - Col) ' check for an invert
  If Col = 0 Then SetPixel Dest.hdc, x, y, Col ' plot pixel
 Next y 'loop through the y-pixels
 PercentDone.Caption = Int(((x * y) / Total) * 100) & "%" 'calculate the percent done.
 Dest.Refresh
 DoEvents
 If Cancel = True Then GoTo Finish:
Next x 'loop through the x-pixels
Finish:
StartLineArt.Enabled = True
StartEdgeDetect.Enabled = True
End Sub

Private Sub uGreen_Click() ' Make Sure At Least One Colour Is Selected
 If uBlue.Value = 0 _
 And uRed.Value = 0 _
 And uGreen.Value = 0 _
 Then uGreen.Value = 1
End Sub
Private Sub uRed_Click() ' Make Sure At Least One Colour Is Selected
 If uBlue.Value = 0 _
 And uRed.Value = 0 _
 And uGreen.Value = 0 _
 Then uRed.Value = 1
End Sub
Private Sub uBlue_Click() ' Make Sure At Least One Colour Is Selected
 If uBlue.Value = 0 _
 And uRed.Value = 0 _
 And uGreen.Value = 0 _
 Then uBlue.Value = 1
End Sub
