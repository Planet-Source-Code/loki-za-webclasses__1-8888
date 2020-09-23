VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form WebForm2 
   Caption         =   "WebTransform Export"
   ClientHeight    =   3645
   ClientLeft      =   2040
   ClientTop       =   2730
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4815
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Export"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Would you like to clear the variable?"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Please select the name of the variable of you would like to use to concatenate the strings"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Destination?"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "WebForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
End Sub

Private Sub Command2_Click()
    CommonDialog1.DefaultExt = ".trf"
    CommonDialog1.Filter = "Transformed File (*.trf)|*.trf|"
    CommonDialog1.Action = 1
    Text2.Text = CommonDialog1.FileTitle
End Sub

Private Sub Command3_Click()
Dim iFile As Long
Dim iLine As String
Dim counter As Long

    iFile = FreeFile

    Open Text2.Text For Output As #iFile
            If Check1.Value = "1" Then
                Print #iFile, Text3.Text & " = " & Chr(34) & Chr(34)
            End If
        For counter = 0 To WebForm1.List1.ListCount - 1
            Print #iFile, Text3.Text & " = " & Text3.Text & " & " & Chr(34) & WebForm1.List1.List(counter) & Chr(34)
        Next
    Close iFile
    WebForm2.Hide
    MsgBox ("Successfully exported " & Text2.Text)
End Sub

Private Sub Command4_Click()
    WebForm2.Hide
End Sub

Private Sub Form_Load()
    WebForm1.CommonDialog1.DefaultExt = ".trf"
End Sub

