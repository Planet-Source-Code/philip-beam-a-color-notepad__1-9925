VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Color Note  -  By Philip Beam  -  Unititled"
   ClientHeight    =   3375
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3125
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Main.frx":0000
      Top             =   250
      Width           =   5775
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Main.frx":013C
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Label5"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   0
      X2              =   5760
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Filename 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Status 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    PopupMenu mnuFile
End Sub

Private Sub Label2_Click()
    PopupMenu mnuEdit
End Sub

Private Sub Label3_Click()
    About.Visible = True
End Sub

Private Sub Label4_Click()
     PopupMenu mnuOptions
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetText Main.Text1.SelText
End Sub

Private Sub mnuEditCut_Click()
    Clipboard.SetText Main.Text1.SelText
    Main.Text1.SelText = ""
End Sub

Private Sub mnuEditDelete_Click()
    Main.Text1.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
    Main.Text1.SelText = Clipboard.GetText
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileNew_Click()
    Main.Text1.Text = ""
End Sub

Private Sub mnuFileOpen_Click()
    Main.CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    Main.CommonDialog1.ShowOpen
    Main.RichTextBox1.LoadFile (Main.CommonDialog1.Filename)
    Main.Text1.Text = Main.RichTextBox1.Text
    Main.Caption = "Color Note  -  By Philip Beam  -  " & Main.CommonDialog1.Filename & ""
End Sub

Private Sub mnuFileSave_Click()
    If Main.Status = "saved" Then
            Main.RichTextBox1.SaveFile (Main.Filename.Caption)
            Exit Sub
    End If
    Main.CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    Main.CommonDialog1.ShowSave
    Main.RichTextBox1.SaveFile (Main.CommonDialog1.Filename)
    Main.Caption = "Color Note  -  By Philip Beam  -  " & Main.CommonDialog1.Filename & ""
    Main.Status.Caption = "saved"
    Main.Filename.Caption = Main.CommonDialog1.Filename
End Sub

Private Sub mnuSaveAs_Click()
    Main.CommonDialog1.Filter = "Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*"
    Main.CommonDialog1.ShowSave
    Main.RichTextBox1.SaveFile (Main.CommonDialog1.Filename)
    Main.Caption = "Color Note  -  By Philip Beam  -  " & Main.CommonDialog1.Filename & ""
End Sub

Private Sub Text1_Change()
    Main.RichTextBox1.Text = Main.Text1.Text
End Sub
