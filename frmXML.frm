VERSION 5.00
Begin VB.Form frmXML 
   Caption         =   "VB / ADO / XML - J. Brandon George"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   ">>>"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "<<<"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdXML 
      Caption         =   "Add to XML"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnext_Click()
On Error Resume Next
ModADO.Rs1.MoveNext
Text1.Text = ModADO.Rs1!info1
Text2.Text = ModADO.Rs1!num1
Text3.Text = ModADO.Rs1!Date


End Sub

Private Sub cmdprev_Click()
On Error Resume Next
ModADO.Rs1.MovePrevious
Text1.Text = ModADO.Rs1!info1
Text2.Text = ModADO.Rs1!num1
Text3.Text = ModADO.Rs1!Date

End Sub

Private Sub cmdxml_Click()
Dim info As String
Dim num As String
Dim datei As String

Dim cxml As New clsXML

Dim strFileName As String
    strFileName = App.Path & "\" & "axml.xml"
    
    
cxml.Initialize pavAUTO

cxml.OpenFromFile strFileName, True




info = ModADO.Rs1!info1
num = ModADO.Rs1!num1
datei = ModADO.Rs1!Date

cxml.InsertNode "/axml", "info", "" & info & "", , , norchild
cxml.InsertNode "/axml", "number", "" & num & "", , , norchild
cxml.InsertNode "/axml", "date", "" & datei & "", , , norchild

cxml.Save strFileName

MsgBox "data has been saved to XML file, please open the xml file named axml.xml.", vbInformation, "Done"

End Sub

Private Sub Form_Load()
ModADO.OpenDB
Text1.Text = ModADO.Rs1!info1
Text2.Text = ModADO.Rs1!num1
Text3.Text = ModADO.Rs1!Date



End Sub
