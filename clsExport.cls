VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This module is put together with ideas from lots of places, including microsoft, and source
'code samples from Able Consulting.
'
'USES DAO 3.51 Object Library
'
'You can export from any database to any Isam file directly. If you are exporting a table into
' a database, you can enter a new table name as a paramater.
'You can use the properties, or the parameters either way, and there is a filter function
'built in which, when you set all the filter-type properties, will return a filter string for
'use with the windows common dialog
'PROPERTIES:
'   A. (Various) FilterOn   Will include that particular filter in the
'                            "comDlgFilterString" property
'   B. DestDB               Database(or File) path to export INTO
'   C. SourceDB             Database(or File)path to get data from
'   D. Database             Allows access to opened Database for getting TableDefs, etc.
'   E. DataType             Property Using Public Enum for Setting connection and file properties
'   F. SourceTable          Table to get info from in open database
'   G. Desttable            Name of table to insert new data into
'   H. UserName             Not used, was going to use for databases that used user authentication
'   I. UserPwd              Not used, was going to use for databases that used user authentication
'
'METHODS
'   A. OpenConnection       Open database to get info from - Path to database,
'                                                            FileType as DataType enum,
'                                                            OpenExclusive as boolean,
'                                                            dataBase Password as string
'   B. Exporttable          Export file or table -  Path to insert into,
'                                                   Table to take info from
'                                                   TableName to insert into
'                                                   Fields as a comma-delimited string list
'                                                   DataType as DataType Enum
'
'Darryn Frost         i_jester69@hotmail.com          May 25th, 2001
'***********************************************************************************************
Option Explicit
Option Compare Binary

'private variables for all the data
Private m_WorkSpace As DAO.Workspace
Private m_Database As DAO.Database
Private m_DatabasePwd As String
Private m_UserName As String
Private m_UserPWD As String
Private m_DataType As DataTypes
Private m_ConString As String

Private m_DestDB As String
Private m_SourceTable As String
Private m_DestTable As String

Private m_AccessFilter As String
Private m_DBase3Filter As String
Private m_DBase4Filter As String
Private m_DBase5Filter As String
Private m_Excel30Filter As String
Private m_Excel40Filter As String
Private m_Excel50Filter As String
Private m_Excel80Filter As String
Private m_FoxPro20Filter As String
Private m_FoxPro25Filter As String
Private m_FoxPro26Filter As String
Private m_FoxPro30Filter As String
Private m_HtmlFilter As String
Private m_Paradox3XFilter As String
Private m_Paradox4XFilter As String
Private m_Paradox5XFilter As String
Private m_TextFilter As String

'data backend types used as the connect string
Const m_MSACCESS = "Microsoft Access"
Const m_DBASEIII = "Dbase III;"
Const m_DBASEIV = "Dbase IV;"
Const m_DBASE5 = "Dbase 5.0;"
Const m_EXCEL30 = "Excel 3.0;"
Const m_EXCEL40 = "Excel 4.0;"
Const m_EXCEL50 = "Excel 5.0;"
Const m_EXCEL80 = "Excel 5.0;"
Const m_FOXPRO20 = "FoxPro 2.0;"
Const m_FOXPRO25 = "FoxPro 2.5;"
Const m_FOXPRO26 = "FoxPro 2.6;"
Const m_FOXPRO30 = "FoxPro 3.0;"
Const m_HTML = "Html Export;"
Const m_PARADOX3X = "Paradox 3.X;"
Const m_PARADOX4X = "Paradox 4.X;"
Const m_PARADOX5X = "Paradox 5.X;"
Const m_TEXTFILES = "Text;"

'import/export data types
Public Enum DataTypes
  gT_ACCESS = 1
  gT_DBASEIII = 2
  gT_DBASEIV = 3
  gT_DBASEV = 4
  gt_EXCEL30 = 5
  gt_EXCEL50 = 6
  gt_EXCEL40 = 7
  gt_EXCEL80 = 8
  gT_FOXPRO20 = 9
  gT_FOXPRO25 = 10
  gT_FOXPRO26 = 11
  gT_FOXPRO30 = 12
  gT_HTML = 13
  gT_PARADOX3X = 14
  gT_PARADOX4X = 15
  gT_PARADOX5X = 16
  gT_TEXTFILE = 17
End Enum

Public Function ExportTable(Optional ByVal rsToDB As String, Optional ByVal rsFromTable As String, _
    Optional ByVal rsToTable As String, Optional ByVal fields As String, _
    Optional ByVal pDataType As DataTypes)

Dim i As Integer
Dim strSQL As String
Dim sConnect As String
Dim sDBName As String
Dim sFileName As String
Dim sNewTblName As String

On Error GoTo cmdExport_Error

'see if a datatype was passed in, if not use the property, then default(text)
If Not (pDataType > 0 And pDataType < 19) Then pDataType = m_DataType
If Not (pDataType > 0 And pDataType < 19) Then pDataType = gT_TEXTFILE

'see if a destination database was passed in, if not use the property,
'then raise error if no property
If rsToDB = "" Then rsToDB = m_DestDB
If rsToDB = "" Then Err.Raise vbObjectError + 101, "ExportTable", "No destination database supplied"

'see if a source tablename was passed in, if not use the property,
'then raise error if no property
If rsFromTable = "" Then rsFromTable = m_SourceTable
If rsFromTable = "" Then Err.Raise vbObjectError + 102, "ExportTable", "No source table supplied"

'see if a source tablename was passed in, if not use the property,
'then use the 'from' tablename
If rsToTable = "" Then rsToTable = rsFromTable

If fields = "" Then
    fields = "*"
End If

strSQL = "SELECT " & fields & " INTO"

sConnect = SetDataType(pDataType, rsToDB)

Dim strEXEC As String
strEXEC = strSQL & sConnect & "[" & rsToTable & "] FROM [" & StripOwner(rsFromTable) & "]"
Debug.Print strEXEC

m_Database.Execute strEXEC

Exit Function

cmdExport_Error:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function OpenConnection(ByVal pFilePath As String, pDataType As DataTypes, _
    Optional OpenExclusive As Boolean, Optional pPwd As String) As Boolean
On Error GoTo OpenError

Dim strConnect As String
Dim pDataSource As String
Dim pReadOnly As Boolean
pReadOnly = True
If pPwd = "" Then pPwd = m_DatabasePwd
If pPwd <> "" Then strConnect = ";pwd=" & pPwd

'set the database name for non Microsoft Access and Btrieve dbs that
  'came from the Common Dialog
  If pDataType <> gT_ACCESS And pDataType And pDataType <> gt_EXCEL30 _
            And pDataType <> gt_EXCEL40 And pDataType <> gt_EXCEL50 _
            And pDataType <> gt_EXCEL80 Then
     'need to strip off filename for these dbs
     pDataSource = StripFileName(pFilePath)
  Else
    pDataSource = pFilePath
  End If

Select Case pDataType
    Case gT_ACCESS
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, strConnect)
    Case gT_PARADOX3X
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_PARADOX3X & strConnect)
    Case gT_PARADOX4X
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_PARADOX4X & strConnect)
    Case gT_PARADOX5X
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_PARADOX5X & strConnect)
    Case gT_FOXPRO20
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_FOXPRO20 & strConnect)
    Case gT_FOXPRO25
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_FOXPRO25 & strConnect)
    Case gT_FOXPRO26
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_FOXPRO26 & strConnect)
    Case gT_FOXPRO30
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_FOXPRO30 & strConnect)
    Case gT_DBASEIII
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_DBASEIII & strConnect)
    Case gT_DBASEIV
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_DBASEIV & strConnect)
    Case gT_DBASEV
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_DBASE5 & strConnect)
    Case gt_EXCEL30
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_EXCEL30 & strConnect)
    Case gt_EXCEL40
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_EXCEL40 & strConnect)
    Case gt_EXCEL50
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_EXCEL50 & strConnect)
    Case gt_EXCEL80
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_EXCEL80 & strConnect)
    Case gT_TEXTFILE
      Set m_Database = m_WorkSpace.OpenDatabase(pDataSource, OpenExclusive, pReadOnly, m_TEXTFILES & strConnect)
End Select
  
Exit Function
  
OpenError:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function SetDataType(pDataType As DataTypes, rsToDB As String) As String
Dim txt As String

Select Case pDataType
    Case gT_TEXTFILE
      txt = "[Text;database=" & StripFileName(rsToDB) & "]."
    Case gt_EXCEL30
      txt = "[Excel 3.0;database=" & rsToDB & "]."
    Case gt_EXCEL40
      txt = "[Excel 4.0;database=" & rsToDB & "]."
    Case gt_EXCEL50
      txt = "[Excel 5.0;database=" & rsToDB & "]."
    Case gt_EXCEL80
      txt = "[Excel 8.0;database=" & rsToDB & "]."
    Case gT_HTML
      txt = "[Html Export;database=" & StripFileName(rsToDB) & "]."
    Case gT_ACCESS
      txt = "[;database=" & rsToDB & ";pwd=]."
    Case gT_PARADOX3X
      txt = "[Paradox 3.X;database=" & StripFileName(rsToDB) & "]."
    Case gT_PARADOX4X
      txt = "[Paradox 4.X;database=" & StripFileName(rsToDB) & "]."
    Case gT_PARADOX5X
      txt = "[Paradox 5.X;database=" & StripFileName(rsToDB) & "]."
    Case gT_FOXPRO20
      txt = "[FoxPro 2.0;database=" & StripFileName(rsToDB) & "]."
    Case gT_FOXPRO25
      txt = "[FoxPro 2.5;database=" & StripFileName(rsToDB) & "]."
    Case gT_FOXPRO26
      txt = "[FoxPro 2.6;database=" & StripFileName(rsToDB) & "]."
    Case gT_FOXPRO30
      txt = "[FoxPro 3.0;database=" & StripFileName(rsToDB) & "]."
    Case gT_DBASEIII
      txt = "[dBase III;database=" & StripFileName(rsToDB) & "]."
    Case gT_DBASEIV
      txt = "[dBase IV;database=" & StripFileName(rsToDB) & "]."
    Case gT_DBASEV
      txt = "[dBase 5.0;database=" & StripFileName(rsToDB) & "]."
  End Select
  
SetDataType = txt
  
End Function
'------------------------------------------------------------
'this function strips the file name from a path\file string
'------------------------------------------------------------
Private Function StripFileName(rsFileName As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = "\" Then
      Exit For
    End If
  Next

  StripFileName = Mid(rsFileName, 1, i - 1)

End Function

'------------------------------------------------------------
'strips the owner off of ODBC table names
'------------------------------------------------------------
Private Function StripOwner(ByVal rsTblName As String) As String

  If InStr(rsTblName, ".") > 0 Then
    rsTblName = Mid(rsTblName, InStr(rsTblName, ".") + 1, Len(rsTblName))
  End If
  StripOwner = rsTblName

End Function
'Strip the extension off filenames for databases that only need a table name
Private Function stripExtension(ByVal filename As String) As String
Dim stmp As String
Dim i As Integer
For i = Len(filename) To 1 Step -1
      If Mid(filename, i, 1) = "\" Then
        Exit For
      End If
    Next
stmp = Mid(filename, i + 1, Len(filename))
    'strip off the extension
    For i = 1 To Len(stmp)
      If Mid(stmp, i, 1) = "." Then
        Exit For
      End If
    Next
    stripExtension = Left(stmp, i - 1)

End Function
Public Property Get Database() As DAO.Database
Set Database = m_Database
End Property
Public Property Let dataType(ByVal ptNewValue As DataTypes)
m_DataType = ptNewValue
End Property
Public Property Let UserPassword(ByVal strNewValue As String)
m_UserPWD = strNewValue
End Property
Public Property Let UserName(ByVal strNewValue As String)
m_UserName = strNewValue
End Property
Public Property Let DatabasePassword(ByVal strNewValue As String)
m_DatabasePwd = strNewValue
End Property
Public Property Get DestDatabase() As String
DestDatabase = m_DestDB
End Property
Public Property Let DestDatabase(ByVal strNewValue As String)
m_DestDB = strNewValue
End Property
Public Property Let SourceTable(ByVal strNewValue As String)
m_SourceTable = strNewValue
End Property
Public Property Let DestTable(ByVal strNewValue As String)
m_DestTable = strNewValue
End Property
Public Property Let TextFilterOn(ByVal blnValue As Boolean)
m_TextFilter = "Delimited Text Files (*.txt)|*.txt"
If blnValue <> True Then m_TextFilter = ""
End Property
Public Property Let HtmlFilterOn(ByVal blnValue As Boolean)
m_HtmlFilter = "HTML Files (*.htm)|*.htm"
If blnValue <> True Then m_HtmlFilter = ""
End Property
Public Property Let AccessFilterOn(ByVal blnValue As Boolean)
m_AccessFilter = "Access Database Tables (*.mdb)|*.mdb"
If blnValue <> True Then m_AccessFilter = ""
End Property
Public Property Let DBase3FilterOn(ByVal blnValue As Boolean)
m_DBase3Filter = "DBase 3.0 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_DBase3Filter = ""
End Property
Public Property Let DBase4FilterOn(ByVal blnValue As Boolean)
m_DBase4Filter = "DBase 4.0 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_DBase4Filter = ""
End Property
Public Property Let DBase5FilterOn(ByVal blnValue As Boolean)
m_DBase5Filter = "DBase 5.0 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_DBase5Filter = ""
End Property
Public Property Let FoxPro20FilterOn(ByVal blnValue As Boolean)
m_FoxPro20Filter = "FoxPro 2.0 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_FoxPro20Filter = ""
End Property
Public Property Let FoxPro25FilterOn(ByVal blnValue As Boolean)
m_FoxPro25Filter = "FoxPro 2.5 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_FoxPro25Filter = ""
End Property
Public Property Let FoxPro26FilterOn(ByVal blnValue As Boolean)
m_FoxPro26Filter = "FoxPro 2.6 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_FoxPro26Filter = ""
End Property
Public Property Let FoxPro30FilterOn(ByVal blnValue As Boolean)
m_FoxPro30Filter = "FoxPro 3.0 Files (*.dbf)|*.dbf"
If blnValue <> True Then m_FoxPro30Filter = ""
End Property
Public Property Let Paradox3XFilterOn(ByVal blnValue As Boolean)
m_Paradox3XFilter = "Paradox 3X Files (*.db)|*.db"
If blnValue <> True Then m_Paradox3XFilter = ""
End Property
Public Property Let Paradox4XFilterOn(ByVal blnValue As Boolean)
m_Paradox4XFilter = "Paradox 4X Files (*.db)|*.db"
If blnValue <> True Then m_Paradox4XFilter = ""
End Property
Public Property Let Paradox5XFilterOn(ByVal blnValue As Boolean)
m_Paradox5XFilter = "Paradox 5X Files (*.db)|*.db"
If blnValue <> True Then m_Paradox5XFilter = ""
End Property
Public Property Let Excel30FilterOn(ByVal blnValue As Boolean)
m_Excel30Filter = "Excel 3.0 Files (*.xls)|*.xls"
If blnValue <> True Then m_Excel30Filter = ""
End Property
Public Property Let Excel40FilterOn(ByVal blnValue As Boolean)
m_Excel40Filter = "Excel 4.0 Files (*.xls)|*.xls"
If blnValue <> True Then m_Excel40Filter = ""
End Property
Public Property Let Excel50FilterOn(ByVal blnValue As Boolean)
m_Excel50Filter = "Excel 5.0 Files (*.xls)|*.xls"
If blnValue <> True Then m_Excel50Filter = ""
End Property
Public Property Let Excel80FilterOn(ByVal blnValue As Boolean)
m_Excel80Filter = "Excel 8.0 Files (*.xls)|*.xls"
If blnValue <> True Then m_Excel80Filter = ""
End Property
Public Property Get comDlgFilterString() As String
Dim txtTmp As String

txtTmp = m_AccessFilter
If txtTmp <> "" Then
    If m_DBase3Filter <> "" Then txtTmp = txtTmp & "|" & m_DBase3Filter
Else
    txtTmp = m_DBase3Filter
End If
If txtTmp <> "" Then
    If m_DBase4Filter <> "" Then txtTmp = txtTmp & "|" & m_DBase4Filter
Else
    txtTmp = m_DBase4Filter
End If
If txtTmp <> "" Then
    If m_DBase5Filter <> "" Then txtTmp = txtTmp & "|" & m_DBase5Filter
Else
    txtTmp = m_DBase5Filter
End If
If txtTmp <> "" Then
    If m_Excel30Filter <> "" Then txtTmp = txtTmp & "|" & m_Excel30Filter
Else
    txtTmp = m_Excel30Filter
End If
If txtTmp <> "" Then
    If m_Excel40Filter <> "" Then txtTmp = txtTmp & "|" & m_Excel40Filter
Else
    txtTmp = m_Excel40Filter
End If
If txtTmp <> "" Then
    If m_Excel50Filter <> "" Then txtTmp = txtTmp & "|" & m_Excel50Filter
Else
    txtTmp = m_Excel50Filter
End If
If txtTmp <> "" Then
    If m_Excel80Filter <> "" Then txtTmp = txtTmp & "|" & m_Excel80Filter
Else
    txtTmp = m_Excel80Filter
End If
If txtTmp <> "" Then
    If m_FoxPro20Filter <> "" Then txtTmp = txtTmp & "|" & m_FoxPro20Filter
Else
    txtTmp = m_FoxPro20Filter
End If
If txtTmp <> "" Then
    If m_FoxPro25Filter <> "" Then txtTmp = txtTmp & "|" & m_FoxPro25Filter
Else
    txtTmp = m_FoxPro25Filter
End If
If txtTmp <> "" Then
    If m_FoxPro26Filter <> "" Then txtTmp = txtTmp & "|" & m_FoxPro26Filter
Else
    txtTmp = m_FoxPro26Filter
End If
If txtTmp <> "" Then
    If m_FoxPro30Filter <> "" Then txtTmp = txtTmp & "|" & m_FoxPro30Filter
Else
    txtTmp = m_FoxPro30Filter
End If
If txtTmp <> "" Then
    If m_HtmlFilter <> "" Then txtTmp = txtTmp & "|" & m_HtmlFilter
Else
    txtTmp = m_HtmlFilter
End If
If txtTmp <> "" Then
    If m_Paradox3XFilter <> "" Then txtTmp = txtTmp & "|" & m_Paradox3XFilter
Else
    txtTmp = m_Paradox3XFilter
End If
If txtTmp <> "" Then
    If m_Paradox4XFilter <> "" Then txtTmp = txtTmp & "|" & m_Paradox4XFilter
Else
    txtTmp = m_Paradox4XFilter
End If
If txtTmp <> "" Then
    If m_Paradox5XFilter <> "" Then txtTmp = txtTmp & "|" & m_Paradox5XFilter
Else
    txtTmp = m_Paradox5XFilter
End If
If txtTmp <> "" Then
    If m_TextFilter <> "" Then txtTmp = txtTmp & "|" & m_TextFilter
Else
    txtTmp = m_TextFilter
End If


comDlgFilterString = txtTmp
End Property

Private Sub Class_Initialize()
Set m_WorkSpace = CreateWorkspace("", "admin", "")

'Turn on all the filters by default
TextFilterOn = True
HtmlFilterOn = True
AccessFilterOn = True
DBase3FilterOn = True
DBase4FilterOn = True
DBase5FilterOn = True
FoxPro20FilterOn = True
FoxPro25FilterOn = True
FoxPro26FilterOn = True
FoxPro30FilterOn = True
Paradox3XFilterOn = True
Paradox4XFilterOn = True
Paradox5XFilterOn = True
Excel30FilterOn = True
Excel40FilterOn = True
Excel50FilterOn = True
Excel80FilterOn = True
End Sub

Private Sub Class_Terminate()
Set m_WorkSpace = Nothing
Set m_Database = Nothing
End Sub
