VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Window State"
      Height          =   1215
      Left            =   1500
      TabIndex        =   11
      Top             =   480
      Width           =   1395
      Begin VB.OptionButton Option2 
         Caption         =   "Hidden"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Normal"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   14
         Top             =   420
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Minimized"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   13
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Maximized"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   12
         Top             =   900
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      Height          =   255
      Left            =   2940
      TabIndex        =   9
      Text            =   "c:\"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   2940
      TabIndex        =   6
      Text            =   "/C pause"
      Top             =   660
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Actions"
      Height          =   1215
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   1395
      Begin VB.OptionButton Option1 
         Caption         =   "Print"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   5
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Edit"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Open"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   315
      Left            =   2940
      TabIndex        =   1
      Top             =   1380
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Text            =   "command.com"
      Top             =   180
      Width           =   4455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   60
      TabIndex        =   18
      Top             =   2040
      Width           =   4515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Other Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3720
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   1800
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Send Me Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Working Directory"
      Height          =   195
      Index           =   2
      Left            =   2940
      TabIndex        =   10
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Parameters"
      Height          =   195
      Index           =   1
      Left            =   2940
      TabIndex        =   8
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Document/Program"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Action As String
Dim WinState As Integer
Private Sub Command1_Click()
    'log the command in the debug window
    Debug.Print Action & Text1.Text
    'log whether the command succeeded or not
    Debug.Print "Success: " & _
    ShellDocument(Text1.Text, Action, Text2.Text, Text3.Text, WinState)
End Sub

Private Sub Form_Load()
    Action = "Open"
    WinState = 1
    Label2.ToolTipText = "mailto:actoews@hotmail.com"
    Label3.ToolTipText = "If you liked this code, click here to vote for it, or to view my other psc submissions"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Caption = ""
End Sub

Private Sub Label2_Click()
    ShellDocument "mailto:actoews@hotmail.com"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Caption = Label2.ToolTipText
End Sub

Private Sub Label3_Click()
    ShellDocument "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=136063&strAuthorName=Alan%20Toews&txtMaxNumberOfEntriesPerPage=25"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Caption = Label3.ToolTipText
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.Caption = ""
End Sub

Private Sub Option1_Click(Index As Integer)
    Action = Option1(Index).Caption
End Sub

Private Sub Option2_Click(Index As Integer)
    WinState = Index
End Sub
