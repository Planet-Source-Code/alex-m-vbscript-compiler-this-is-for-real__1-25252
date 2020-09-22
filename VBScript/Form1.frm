VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "VBScript Compiler"
   ClientHeight    =   3255
   ClientLeft      =   3750
   ClientTop       =   3150
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
      End
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
      End
      Begin VB.Menu save 
         Caption         =   "&Save"
      End
      Begin VB.Menu ab 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu run 
      Caption         =   "&Run"
      Begin VB.Menu lastrun 
         Caption         =   "&Last Run"
      End
      Begin VB.Menu run2 
         Caption         =   "&Run"
      End
      Begin VB.Menu ac 
         Caption         =   "-"
      End
      Begin VB.Menu compile 
         Caption         =   "&Compile"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "&About VBScript Compiler..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show 'show about screen
End Sub

Private Sub compile_Click()
    CommonDialog1.Filter = "VBScript File (*.vbs)|*.vbs" 'Only save as .VBS files
    CommonDialog1.ShowSave 'Show the saving dialog
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1 'open chosen file
        Print #1, Text1.Text 'Save as a string to VBScript File
        Close #1 'close chosen file
    End If
    MsgBox "Compile Complete!", vbExclamation, "Complete!" 'tell user it is done
End Sub

Private Sub exit_Click()
End 'very simple, closes program :)
End Sub

Private Sub lastrun_Click()
'this simply looks for that "temporary file" that was created when you run a file in
'this program.  Very Simple!  If file is not there then it displays a message box  :)


On Error GoTo Otis 'If there are any errors, goto "Otis"
nResult = Shell("start.exe C:\Windows\Temp\tempscript.vbs", vbHide) 'Call Shell Command
Exit Sub 'make it exit the sub, everything went perfect!            'to open this file
                                                                   
Otis: 'this is "Otis"
MsgBox "Error:  Last run not found!", vbCritical, "Error!" 'Message Box displays error
End Sub

Private Sub new_Click()
Text1.Text = "" 'simple
End Sub

Private Sub open_Click()
Wrap$ = Chr$(13) + Chr$(10) ' make the wrap char thingy
    CommonDialog1.Filter = "VBSC File (*.VBSC)|*.VBSC" 'Filter only these files
    CommonDialog1.ShowOpen  'Show the open dialog
    If CommonDialog1.FileName <> "" Then
        Form1.MousePointer = 11 'cursor turns into an hour glass
        Open CommonDialog1.FileName For Input As #1
        On Error GoTo Giant: 'if file is to large
        Do Until EOF(1)
'!!!Putting text from file to text box!!!
            Line Input #1, LineOfText$
            AllText$ = AllText$ & LineOfText$ & Wrap$
        Loop
        
        Text1.Text = AllText$
        Text1.Enabled = True
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Fixit:
        Form1.MousePointer = 0 'Cursor is an arrow
        Close #1 'done with file, close it
    End If
    Exit Sub
Giant:
    MsgBox "Error: This file is too large to open!", vbCritical, "Error!" 'display error
    Resume Fixit: 'Goto Fixit
End Sub

Private Sub run2_Click()
 'runs what you have by saving a temporary file and opening it-Simple :)
        Open "C:\Windows\Temp\tempscript.vbs" For Output As #1
        Print #1, Text1.Text ' save file as string
        Close #1
nResult = Shell("start.exe C:\Windows\Temp\tempscript.vbs", vbHide) 'Call Shell Command
                                                                    'to run this file
End Sub

Private Sub save_Click()

    CommonDialog1.Filter = "VBSC File (*.VBSC)|*.VBSC" 'filter only these files
    CommonDialog1.ShowSave 'Show the saving dialog
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1 'open chosen file
        Print #1, Text1.Text 'Save as a string to file
        Close #1 'close chosen file
    End If

End Sub
