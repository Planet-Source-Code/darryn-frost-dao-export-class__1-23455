VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DAO Export"
   ClientHeight    =   4965
   ClientLeft      =   2430
   ClientTop       =   1575
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Destination File or Database:"
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   6495
      Begin VB.TextBox txtDestPath 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   5055
      End
      Begin VB.CommandButton cmdSetDest 
         Caption         =   "Find..."
         Height          =   330
         Left            =   5520
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source Data"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin VB.ListBox lstFields 
         Height          =   1425
         Left            =   4200
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ListBox lstTables 
         Height          =   1425
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   1245
      End
      Begin VB.CommandButton cmdSetSource 
         Caption         =   "Find..."
         Height          =   330
         Left            =   5520
         TabIndex        =   4
         Top             =   600
         Width           =   765
      End
      Begin VB.TextBox txtSourcePath 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   4980
      End
      Begin VB.Label lblFields 
         Caption         =   "Fields"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Source Tables:"
         Height          =   225
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Database or file:"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog dlgDB 
      Left            =   360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "Export"
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   4440
      Width           =   1065
   End
   Begin VB.CommandButton butClose 
      Caption         =   "Close"
      Height          =   330
      Left            =   3840
      TabIndex        =   0
      Top             =   4440
      Width           =   1065
   End
End
Attribute VB_Name = "frmDemo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private m_Export As clsExport
Private m_Filename As String
Private m_ExpType As DataTypes
Private m_ExpTableName As String
Private dbSource As Database

Private Sub ListSourceTables()
Dim tDef As TableDef
Set dbSource = m_Export.Database
For Each tDef In dbSource.TableDefs
    If Left$(tDef.Name, 4) <> "MSys" Then
        lstTables.AddItem tDef.Name
    End If
Next
Set tDef = Nothing
End Sub

Private Sub lstTables_Click()
Dim tDef As DAO.TableDef
Dim tField As DAO.Field
lstFields.Clear
For Each tDef In dbSource.TableDefs
    If tDef.Name = lstTables.Text Then
        For Each tField In tDef.fields
        lstFields.AddItem tField.Name
        Next
    End If
Next
Set tDef = Nothing
Set tField = Nothing
cmdDoIt.Enabled = True
End Sub

Private Sub ResetForm()
    
    lstTables.Clear
    lstFields.Clear
    
    cmdDoIt.Enabled = False
End Sub

Private Sub butClose_Click()
End
End Sub


Private Sub cmdDoIt_Click()

Dim i As Integer
Dim strFromtable As String
Dim strFields As String

On Error Resume Next

For i = 0 To lstFields.ListCount - 1
    If lstFields.Selected(i) = True Then
        If strFields = "" Then
            strFields = lstFields.List(i)
        Else
            strFields = strFields & ", " & lstFields.List(i)
        End If
    End If
Next i

strFromtable = lstTables.Text

m_Export.ExportTable txtDestPath, strFromtable, m_Filename, strFields, m_ExpType

If Err.Number <> 0 Then
    MsgBox "Error# " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Exit Sub
End If

txtSourcePath.Text = ""
End Sub

Private Sub cmdSetSource_Click()


Dim pType As DataTypes
Dim pwd As String
Dim blnTryAgain As Boolean

blnTryAgain = False

ResetForm   'clear boxes

'Set filter string for Access and Dbase4
'Keep track of the index to set open database type later
m_Export.AccessFilterOn = True      'index 1
m_Export.DBase3FilterOn = False
m_Export.DBase4FilterOn = False
m_Export.DBase5FilterOn = False
m_Export.Excel30FilterOn = True     'index 2
m_Export.Excel40FilterOn = True     'index 3
m_Export.Excel50FilterOn = True     'index 4
m_Export.Excel80FilterOn = True     'index 5
m_Export.FoxPro20FilterOn = True     'index 6
m_Export.FoxPro25FilterOn = True     'index 7
m_Export.FoxPro26FilterOn = True     'index 8
m_Export.FoxPro30FilterOn = True     'index 9
m_Export.HtmlFilterOn = True         'index 10
m_Export.Paradox3XFilterOn = True     'index 11
m_Export.Paradox4XFilterOn = True     'index 12
m_Export.Paradox5XFilterOn = True     'index 13
m_Export.TextFilterOn = True         'index 14


With dlgDB
    .DialogTitle = "Find Source File"
    .Flags = FileOpenConstants.cdlOFNPathMustExist
    .Filter = m_Export.comDlgFilterString   'Combination filter set by setting the above filter properties
    .filename = ""
    .FilterIndex = 1
    .ShowOpen
End With

If Len(dlgDB.filename) > 0 Then
    txtSourcePath = dlgDB.filename
Else
    Exit Sub
End If

If dlgDB.FilterIndex = 1 Then
    pType = gT_ACCESS             'DataType Enumeration in clsExport
ElseIf dlgDB.FilterIndex = 2 Then
    pType = gt_EXCEL30
ElseIf dlgDB.FilterIndex = 3 Then
    pType = gt_EXCEL40
ElseIf dlgDB.FilterIndex = 4 Then
    pType = gt_EXCEL50
ElseIf dlgDB.FilterIndex = 5 Then
    pType = gt_EXCEL80
ElseIf dlgDB.FilterIndex = 6 Then
    pType = gT_FOXPRO20
ElseIf dlgDB.FilterIndex = 7 Then
    pType = gT_FOXPRO25
ElseIf dlgDB.FilterIndex = 8 Then
    pType = gT_FOXPRO26
ElseIf dlgDB.FilterIndex = 9 Then
    pType = gT_FOXPRO30
ElseIf dlgDB.FilterIndex = 10 Then
    pType = gT_HTML
ElseIf dlgDB.FilterIndex = 11 Then
    pType = gT_PARADOX3X
ElseIf dlgDB.FilterIndex = 12 Then
    pType = gT_PARADOX4X
ElseIf dlgDB.FilterIndex = 13 Then
    pType = gT_PARADOX5X
ElseIf dlgDB.FilterIndex = 14 Then
    pType = gT_TEXTFILE
End If
On Error Resume Next
tryagain:
'Open connection
m_Export.OpenConnection txtSourcePath, pType, False, pwd

If Err.Number <> 0 Then
  If Err.Number = 3031 Then
    If blnTryAgain = False Then
        blnTryAgain = True
        pwd = InputBox("This database requires a password to connect. Please enter the password.")
        GoTo tryagain
    Else
        blnTryAgain = False
        Exit Sub
    End If
  Else
    MsgBox "Error# " & Err.Number & vbCrLf & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
    Exit Sub
  End If
End If

ListSourceTables

End Sub

Private Sub cmdSetDest_Click()

'Set filter string for Access and Dbase4
'Keep track of the index to set open database type later

m_Export.AccessFilterOn = True      'index 1
m_Export.DBase3FilterOn = True      'index 3
m_Export.DBase4FilterOn = True      'index 4
m_Export.DBase5FilterOn = True      'index 5
m_Export.Excel30FilterOn = True     'index 6
m_Export.Excel40FilterOn = True     'index 7
m_Export.Excel50FilterOn = True     'index 8
m_Export.Excel80FilterOn = True     'index 9
m_Export.FoxPro20FilterOn = True    'index 10
m_Export.FoxPro25FilterOn = True    'index 11
m_Export.FoxPro26FilterOn = True    'index 12
m_Export.FoxPro30FilterOn = True    'index 13
m_Export.HtmlFilterOn = True        'index 14
m_Export.Paradox3XFilterOn = True   'index 15
m_Export.Paradox4XFilterOn = True   'index 16
m_Export.Paradox5XFilterOn = True   'index 17
m_Export.TextFilterOn = True        'index 18

With dlgDB
    .DialogTitle = "Set destination file"
    .Flags = FileOpenConstants.cdlOFNPathMustExist
    .Filter = m_Export.comDlgFilterString
    .filename = ""
    .FilterIndex = 18
    .ShowSave
End With

m_ExpType = dlgDB.FilterIndex

If Len(dlgDB.filename) > 0 Then
    txtDestPath = dlgDB.filename
    If m_ExpType = gT_ACCESS Or m_ExpType = gt_EXCEL30 _
            Or m_ExpType = gt_EXCEL40 Or m_ExpType = gt_EXCEL50 _
            Or m_ExpType = gt_EXCEL80 Then
            m_Filename = InputBox("Enter a name for your exported table")
    Else
        m_Filename = dlgDB.FileTitle
    End If
Else
    Exit Sub
End If
    
End Sub

Private Sub Form_Load()
Set m_Export = New clsExport
ResetForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set dbSource = Nothing
End Sub




