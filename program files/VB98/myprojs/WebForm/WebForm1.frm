VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form WebForm1 
   Caption         =   "WebClass Formatter"
   ClientHeight    =   12270
   ClientLeft      =   195
   ClientTop       =   -1455
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   818
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   979
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Height          =   14295
      ItemData        =   "WebForm1.frx":0000
      Left            =   75
      List            =   "WebForm1.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   14415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu FileMenuOpen 
         Caption         =   "Open                                "
         Shortcut        =   ^O
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu FileMenuSave 
         Caption         =   "Save as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu FileMenuClose 
         Caption         =   "Close"
         Shortcut        =   ^C
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu FileMenuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu FileMenuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "WebForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FileMenu_Click()
    If List1.ListCount = 0 Then
        FileMenuSave.Enabled = False
        FileMenuClose.Enabled = False
        FileMenuExport.Enabled = False
    Else
        FileMenuSave.Enabled = True
        FileMenuClose.Enabled = True
        FileMenuExport.Enabled = True
    End If
End Sub

Private Sub FileMenuClose_Click()
    If List1.ListCount > 0 Then
        List1.Clear
    End If
End Sub

Private Sub FileMenuExport_Click()
    WebForm2.Show
End Sub

Private Sub FileMenuOpen_Click()
Dim iFile As Long
Dim iLine As String
Dim Str As String
Dim strtemp As String
Dim k As Long

    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "HTML (*.htm, *.html)|*.html;*.htm|ASP (*.asp)|*.asp|Untransformed File (*.urf)| *.urf|Transformed File (*.trf)| *.trf|"
    CommonDialog1.Action = 1
    
    If CommonDialog1.FileName <> "" Then
        iFile = FreeFile
        Open CommonDialog1.FileName For Input As #iFile
        List1.Clear

            Do While Not EOF(iFile)
                Input #iFile, iLine
                
            If Len(iLine) > 0 Then
                If InStr(iLine, Chr(34)) Then
                    iLine = Replace(iLine, Chr(34), Chr(39))
                End If
                    List1.AddItem (iLine)
            End If
            Loop
            
            Close iFile
    End If
End Sub

Private Sub FileMenuExit_Click()
    End
End Sub

Private Sub FileMenuSave_Click()
Dim counter As Long

    CommonDialog1.DefaultExt = ".trf"
    CommonDialog1.Filter = "Untransformed File (*.urf)|*.urf"
    CommonDialog1.Action = 2
    
    iFile = FreeFile

    Open CommonDialog1.FileName For Output As #iFile

        For counter = 0 To List1.ListCount - 1
            Print #iFile, List1.List(counter)
        Next
    Close iFile
End Sub

Private Sub Form_Resize()
    List1.Width = WebForm1.ScaleX(WebForm1.Width, 1, 3) - 18
    List1.Height = WebForm1.ScaleY(WebForm1.Height, 1, 3) - 30
End Sub

Private Sub List1_Click()

End Sub
