VERSION 5.00
Begin VB.Form frmDate 
   Caption         =   "Date Calculation"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Process"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================
'          N U M E R I C   V A L I D A T I O N
'=========================================================
Private Function VNumeric(strVerify As String) As Boolean
    VNumeric = CBool(IsNumeric(strVerify))  '~* BOOLEAN VARIABLE, TRUE = VALID; FALSE = INVALID.
End Function

'=========================================================
'            D A T E   C A L C U L A T I O N
'=========================================================
Private Sub cmdAdd_Click()
    Dim strDLen         As Long             '~* STRING LENGTH VARIABLE.
    Dim strDate         As String           '~* TXTDATE VARIABLE.
    Dim strCurr         As String           '~* CURRENT DATE (MM/DD/YYYY).
    
    strCurr = Date                          '~* DEFINING DATE.
    strDate = txtDate.Text                  '~* DEFINING TEXT BOX.
    strDLen = Len(txtDate.Text)             '~* DEFINING TEXT LENGTH.
    
    If txtDate.Text = "" Then               '~* EMPTY FIELD VALIDATION.
        MsgBox "Please enter a numeric value.", vbOKOnly, "Validation"
        Exit Sub
    ElseIf Not VNumeric(strDate) Then       '~* NUMERIC ONLY VALIDATION.
        MsgBox "Please enter a numeric value.", vbOKOnly, "Validation"
        Exit Sub
    ElseIf strDLen >= "6" Then              '~* LENGHT VALIDATION, WILL NOT ALLOW A NUMERIC VALUE EQUAL TO OR GREATER THAN 6 CHARACTERS.
        MsgBox "To circumvent an overflow, please enter a lower value.", vbOKOnly, "Validation"
        Exit Sub
    Else
        'INITIALIZING THE ADDITION OF THE VALID NUMERIC VALUE WITH THE CURRENT DATE.
        strTime = DateAdd("d", strDate, strCurr)
        'PROMPTING YOUR NEWLY FORMED DATE.
        MsgBox strTime, vbOKOnly, "Date Calculation"
    End If
End Sub
