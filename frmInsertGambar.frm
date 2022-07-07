VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInsertGambar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GAMBAR "
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7170
   Icon            =   "frmInsertGambar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton loadFolder 
      Caption         =   "..."
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   3000
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cb1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtNamaFile 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Frame frawarna 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pilih Warna"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton Command5 
         BackColor       =   &H00404040&
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF0000&
         Height          =   375
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdred 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame fraInsertGambar 
      Caption         =   "GAMBAR"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2145
         ScaleWidth      =   6825
         TabIndex        =   3
         Top             =   240
         Width           =   6855
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H000000FF&
         Caption         =   "RESET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmdWarna 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Pilih Warna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   4080
         Top             =   2400
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nama File (klik tombol disamping textbox)"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2760
      Width           =   3735
   End
End
Attribute VB_Name = "frmInsertGambar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type PointApi
   X As Double
   Y As Double
End Type
Public NoPendaftaran As String
Public FormPengirim As String
Private Point1 As PointApi
Private Point2 As PointApi
Private blnMouseDown As Boolean
Private warna As Variant
Private Sub cmdred_Click()
     warna = vbRed
    frawarna.Visible = False
    Shape1.BackColor = warna
End Sub

Private Sub cmdReset_Click()
    Picture1.BackColor = vbWhite
    blnMouseDown = False
    Point1.X = 0
    Point2.X = 0
    Point1.Y = 0
    Point2.Y = 0
End Sub

Private Sub cmdSimpan_Click()
    If txtNamaFile.Text = "" Then Call MsgBox("Pilih Nama Folder dan File terlebih Dahulu"): Exit Sub
    
    Call SavePicture(Picture1.Image, txtNamaFile.Text)
    Call MsgBox("Berhasil")
End Sub

Private Sub cmdtutup_Click()
    Set frmInsertGambar = Nothing
    Unload Me
End Sub

Private Sub cmdWarna_Click()
     frawarna.Visible = True
End Sub

Private Sub Command1_Click()
    Picture2.picture = Picture1.Image
End Sub

Private Sub Command3_Click()
    warna = vbBlue
    frawarna.Visible = False
    Shape1.BackColor = warna
End Sub

Private Sub Command4_Click()
    warna = vbYellow
    frawarna.Visible = False
    Shape1.BackColor = warna
End Sub

Private Sub Command5_Click()
    warna = vbBlack
    frawarna.Visible = False
    Shape1.BackColor = warna
End Sub

Private Sub Form_Load()
    'Call centerForm(Me, MDIUtama)
   ' Call LoadGambar(Picture1, "\\172.16.16.242\exeterbaru\image\IGD2.bmp")
End Sub

Private Sub picSignature_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 blnMouseDown = True
    If blnMouseDown = True Then
        Point1.X = X
        Point1.Y = Y
    End If
    picSignature.Line (X, Y)-(X, Y)
End Sub

Private Sub picSignature_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If blnMouseDown = True Then
        Point2 = Point1
        Point1.X = X
        Point1.Y = Y
    End If
    picSignature.Line (Point1.X, Point1.Y)-(Point2.X, Point2.Y)
End Sub

Private Sub picSignature_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  blnMouseDown = False
End Sub

Private Sub loadFolder_Click()
    cb1.ShowOpen
    txtNamaFile.Text = cb1.FileName
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnMouseDown = True
    If blnMouseDown = True Then
        Point1.X = X
        Point1.Y = Y
    End If
    Picture1.DrawWidth = 3
    Picture1.Line (X, Y)-(X, Y), warna
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnMouseDown = True Then
        Point2 = Point1
        Point1.X = X
        Point1.Y = Y
    End If
    Picture1.DrawWidth = 3
    Picture1.Line (Point1.X, Point1.Y)-(Point2.X, Point2.Y), warna
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     blnMouseDown = False
End Sub
Private Sub LoadGambar(picture As PictureBox, Path As String)
    picture.picture = LoadPicture(Path)
    
    picture.ScaleMode = 3
    picture.AutoRedraw = True
    picture.PaintPicture Picture1.picture, _
    0, 0, picture.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, picture.picture.Width / 26.46, _
    picture.picture.Height / 26.46
    
    picture.picture = Picture1.Image
  
End Sub

