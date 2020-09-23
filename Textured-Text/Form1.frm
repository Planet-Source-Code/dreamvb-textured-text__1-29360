VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Textured Text Example"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtsave 
      Height          =   285
      Left            =   5385
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2925
      Width           =   3345
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   450
      Left            =   7170
      TabIndex        =   16
      Top             =   3420
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   450
      Left            =   5685
      TabIndex        =   15
      Top             =   3420
      Width           =   1440
   End
   Begin VB.ComboBox cbofontstyle 
      Height          =   315
      Left            =   7770
      TabIndex        =   14
      Top             =   1650
      Width           =   1035
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   390
      Pattern         =   "*.jpg;*.bmp;*.gif"
      TabIndex        =   11
      Top             =   2790
      Width           =   1980
   End
   Begin VB.ComboBox cbofontsize 
      Height          =   315
      Left            =   6255
      TabIndex        =   8
      Top             =   1635
      Width           =   675
   End
   Begin VB.ComboBox cbofont 
      Height          =   315
      Left            =   3390
      TabIndex        =   6
      Top             =   1635
      Width           =   1980
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Text            =   "Cool Textured Logo"
      Top             =   1635
      Width           =   1740
   End
   Begin VB.PictureBox picDst 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   330
      ScaleHeight     =   405
      ScaleWidth      =   390
      TabIndex        =   3
      Top             =   6045
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox PicRes 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   810
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make-Font"
      Height          =   450
      Left            =   4215
      TabIndex        =   1
      Top             =   3420
      Width           =   1440
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   105
      Picture         =   "Form1.frx":0C4F
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   577
      TabIndex        =   0
      Top             =   90
      Width           =   8655
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Save Image To:"
      Height          =   195
      Left            =   4215
      TabIndex        =   17
      Top             =   2970
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Font Style"
      Height          =   195
      Left            =   7005
      TabIndex        =   13
      Top             =   1695
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Preview Mode"
      Height          =   195
      Left            =   2715
      TabIndex        =   12
      Top             =   2445
      Width           =   1020
   End
   Begin VB.Image imgpreview 
      Height          =   1035
      Left            =   2430
      Stretch         =   -1  'True
      Top             =   2790
      Width           =   1545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   39
      X2              =   567
      Y1              =   147
      Y2              =   147
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   40
      X2              =   568
      Y1              =   146
      Y2              =   146
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Select a Texture"
      Height          =   195
      Left            =   735
      TabIndex        =   10
      Top             =   2445
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Text Display"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   1695
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Font Size"
      Height          =   195
      Left            =   5475
      TabIndex        =   7
      Top             =   1695
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fonts"
      Height          =   195
      Left            =   2925
      TabIndex        =   5
      Top             =   1695
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
    
End Function
Sub IntTitle()
  Dim X As Integer
  Dim Y As Integer
  Dim mRet As Long
      For Y = 0 To pic1.Height Step PicRes.Height
               For X = 0 To pic1.Width Step PicRes.Width
            BitBlt pic1.hDC, X, Y, PicRes.Width, PicRes.Height, PicRes.hDC, 0, 0, vbSrcCopy
          Next
      Next
      
End Sub

Private Sub cbofontstyle_Click()
    Select Case cbofontstyle.ListIndex
        Case 0
            picDst.FontBold = False
            picDst.FontItalic = False
            picDst.FontUnderline = False
        Case 1
            picDst.FontBold = True
            picDst.FontItalic = False
            picDst.FontUnderline = False
        Case 2
            picDst.FontBold = False
            picDst.FontItalic = True
            picDst.FontUnderline = False
        Case 3
            picDst.FontBold = False
            picDst.FontItalic = False
            picDst.FontUnderline = True
    End Select
    
End Sub

Private Sub Command1_Click()
    IntTitle
    picDst.Cls
    TMessage = Trim(txtMessage.Text)
    If Len(TMessage) <= 0 Then
        MsgBox "Please enter in some text first", vbInformation, "Error now Text entered"
        Exit Sub
    Else
        picDst.Width = pic1.Width
        picDst.Height = pic1.Height
        picDst.FontSize = Val(cbofontsize.Text)
        picDst.FontName = cbofont.Text
        picDst.CurrentX = (picDst.ScaleWidth - picDst.TextWidth(TMessage)) / 2
        picDst.CurrentY = (picDst.ScaleHeight - picDst.TextHeight(TMessage)) / 2
        picDst.Print TMessage
        BitBlt pic1.hDC, 0, 0, picDst.ScaleWidth, picDst.ScaleHeight, picDst.hDC, 0, 0, vbSrcPaint
        pic1.Refresh
        SavePicture pic1.Image, txtsave.Text
        MsgBox "New Image has been saved to " & txtsave.Text, vbInformation, "Work Saved...."
    End If
    
End Sub

Private Sub Command2_Click()
Dim mes As String
    mes = mes & "3D Texture Font VIsual Basic Example" & vbCrLf _
    & "By Ben Jones" & vbCrLf _
    & "Please Vote for meif you like this code"
    MsgBox mes, vbInformation, "About"
    
End Sub

Private Sub Command3_Click()
    Unload Form1
    End
    
End Sub

Private Sub File1_Click()
    imgpreview.Picture = LoadPicture(FixPath(File1.Path) & File1.FileName)
    PicRes.Picture = LoadPicture(FixPath(File1.Path) & File1.FileName)
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To Screen.FontCount - 1 Step 2
        cbofont.AddItem Screen.Fonts(i)
    Next
    i = 0
    For i = 12 To 72 Step 4
        cbofontsize.AddItem i
    Next
    
    cbofontstyle.AddItem "Normal"
    cbofontstyle.AddItem "Bold"
    cbofontstyle.AddItem "Italic"
    cbofontstyle.AddItem "Underline"
    cbofont.ListIndex = 0
    cbofontsize.ListIndex = 6
    cbofontstyle.ListIndex = 0
    
    File1.Path = FixPath(App.Path) & "textures\"
    txtsave.Text = FixPath(App.Path) & "Saved.bmp"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Set pic1 = Nothing
    Set picDst = Nothing
    Set PicRes = Nothing
    
End Sub
