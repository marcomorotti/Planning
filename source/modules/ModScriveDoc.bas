
'Attribute VB_Name = "bas_crystal_Document_QrySQL_FrmRptRecordSource_2Wrd_Create_qObjx"
Option Compare Database
Option Explicit
'Here is code to DOCUMENT the SQL stored for each QUERY. You can also document the source for FORMS and REPORTS.
'You can also create a query that lists the main object names and types in your database.
'
'make a new general module in your Access database
'copy this code and paste into Access
'make reference to
'Microsoft DAO Object Library or Microsoft Office ##.0 Access Database Engine Object Library
'(from the menu: Tools, References...)
'Debug, Compile and then Save
'click in first procedure, Run_Word_CreateDocumention_SQL, and press F5 to Run!
'A Word document showing the SQL for all your queries will be created.
'
'
'To use this module to document RecordSource for forms or reports, run one of the following procedures:
'Run_Word_CreateDocumention_Forms
'Run_Word_CreateDocumention_Reports
'
'
'To create a query from the MSysObjects table with a list of all the main object names and types in your database,
' run Run_Create_qObjex_byCrystal_Query

'
' This code was originally written by Crystal
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' NOTE: comment .Style assignment if using Access 2000 or below
'
' INCLUDES
' Word_CreateDocumention
' RunDocumentor
' MakeAccessDocTables
' Word_Write_SQL
' Word_Format_SQL
'
' QUICK LAUNCH
' Run_Word_CreateDocumention_SQL
' Run_Word_CreateDocumention_Forms
' Run_Word_CreateDocumention_Reports
' Run_Create_qObjex_byCrystal_Query
'
'CALLS
' xDoesExist
' xRSql
' xStartTime
' xEndTime
' xReportElapsedTime
' ... and more

' while developing, use early binding and reference:
' Microsoft Word Object Library
'
' Late binding is currently used,
' so it is not necessary to reference the Word Object Library
'
'~~~~~~~~~~~~~~~~~~
      'NEEDS reference to Microsoft DAO Library
      'or
      'Microsoft Office ##.0 Access Database Engine Object Library
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
Dim gStartTime As Date
'
'~~~~~~~~~~~~~~~~~~~~~ Run_Word_CreateDocumention
Sub Run_Word_CreateDocumention_SQL()

   'click HERE          press F5 to --> Write QUERY SQL statements to Word
   Word_CreateDocumention 1
   
End Sub
  
'~~~~~~~~~~~~~~~~~~~~~ Run_Word_CreateDocumention_Forms
Sub Run_Word_CreateDocumention_Forms()

   'click HERE          press F5 to --> Write FORM RecordSource(s) to Word
   Word_CreateDocumention 2
   
End Sub

'~~~~~~~~~~~~~~~~~~~~~ Run_Word_CreateDocumention_Reports
Sub Run_Word_CreateDocumention_Reports()

   'click HERE          press F5 to -->  Write REPORT RecordSource(s) to Word
   Word_CreateDocumention 3
   
End Sub

Sub Run_Create_qObjex_byCrystal_Query()
'110712 Crystal
   'click HERE          press F5 to -->  Create qObjex_byCrystal Query
   xCreate_qObjex
End Sub

'~~~~~~~~~~~~~~~~~~~~~ Word_CreateDocumention

Sub Word_CreateDocumention( _
   Optional pDocType As Integer = 1 _
   , Optional booAppVisible As Boolean = True _
   , Optional dbNameAnalyze As String = "" _
   )

   'CREATE WORD DOCUMENT with
   ' 1. SQL for All queries -- DEFAULT
   ' OR
   ' 2. RecordSource for All Forms
   ' OR
   ' 3. RecordSource for All Reports
    
   'crystal
   'strive4peace2010 at yahoo dot com
  
   
   'PARAMETERS
   '  pDocType
   '    1=Query, 2=Form, 3=Report
   '
   '  dbNameAnalyze is full path and name of
   '    database to analyze --: ie: "c:\data\MyDatabase.mdb"
   '    if missing, uses CurrentDb
   '    warning: no error checking on passed db name
   
   'NEEDS TABLES
   ' xl_Doc_DBs
   ' xl_Doc_Source
    
   'CALLS
   '   MakeAccessDocTables
   '   RunDocumentor
   '   Word_Write_SQL
   '   Word_Format_SQL
   '   xDeleteFile
   '   xReportElapsedTime
   '   xEndTime
    
   Dim appWord As Object
    
   Dim mPathFilename As String _
      , nDbID As Long _
      , sDocType As String _
      , s As String
   
   xStartTime
   
   Select Case pDocType
      Case 1: sDocType = "Query"
      Case 2: sDocType = "Form"
      Case 3: sDocType = "Report"
   End Select
   
   On Error GoTo Proc_Err
    
   ' optionally, you can collect a path
   ' to write to instead of the currentDb path
   
   ' you can  rename the Word Doc that is created
   ' of course <g>
   
   ' the current database will be used for the analysis tables
   ' since the optional parameter is not passed
   ' if tables to hold analysis do not yet exist, they will be made
    
   If Not MakeAccessDocTables( _
      "xl_Doc_DBs" _
      , "xl_Doc_Source" _
      ) Then
      
      'If MsgBox( _
         "Documention Tables already exist, do you want to replace what is there?" _
         , vbYesNo + vbDefaultButton2 _
         , "Replace Documentation Tables?" _
         ) Then
      'no code written for this -- data will be added
      
   End If
   
   '---------------- Do the analysis
   ' write all query info to table
   ' for the current database
   
   nDbID = RunDocumentor( _
      "xl_Doc_DBs" _
      , "xl_Doc_Source", , True, pDocType)
      
   
   If nDbID = 0 Then
      MsgBox "record creation not successful", , "aborting..."
      GoTo Proc_exit
   End If
   
   '---------------- Initialize Word
   Set appWord = CreateObject("Word.Application")
   '---
   
   ' make Word visible
   If booAppVisible Then appWord.Visible = True

   'make a new Word document
   appWord.Documents.Add

   'get the records we just analyzed
   
   s = "SELECT * FROM " & "xl_Doc_Source" _
      & " WHERE dbID = " & nDbID & ";"
      
   '---------------- Write results to Word
   If Not Word_Write_SQL(appWord, s, pDocType) Then
      MsgBox "Word document creation not successful", , "aborting..."
      GoTo Proc_exit
   End If
   
'   appWord.ActiveDocument.Save
   
   '---------------- Format the Word results
   Word_Format_SQL appWord
      
   'add title to top
   With appWord
      .Selection.HomeKey Unit:=6  'wdStory
      .Selection.SplitTable
      'wdLine=5
      .Selection.MoveUp Unit:=5, Count:=1
      .Selection.TypeText Text:= _
         "Documentation for " _
         & CurrentDb.name _
         & Chr(13) & Chr(10) & "VBA code written by Crystal ( strive4peace2010 at yahoo.com )"
      .Selection.TypeParagraph
   End With
   
   '---------------- Save and Close
   mPathFilename = CurrentProject.Path _
      & "\" & sDocType _
      & "_Documentation_" _
         & Format(Now(), "yymmdd_hnn_") _
         & Replace(Dir(CurrentDb.name), ".", "_") _
         & ".doc"
      
   'if the file that you want to save to already exists
   ' then delete it
         
   If Not xDeleteFile(mPathFilename, True) Then
      MsgBox "Cannot save file in Word to --> " _
         & vbCrLf & vbCrLf _
         & mPathFilename _
         & vbCrLf & vbCrLf _
         & "leaving code open -- you'll need to exit Access after you look at the file in Word " _
         , , "Aborting the save in Word"
      
   Else
      appWord.ActiveDocument.SaveAs mPathFilename
      appWord.ActiveDocument.Close False
      appWord.Quit
   End If
   
'~CL
   Set appWord = Nothing
'   DoEvents
   
   
   xReportElapsedTime "Done creating Documentation " _
      & vbCrLf _
      & " path --> " _
      & vbCrLf & vbCrLf _
      & CurrentProject.Path & "\" _
      & vbCrLf _
      & " filename --> " _
      & vbCrLf & vbCrLf _
      & mPathFilename
      
   
Proc_exit:
   On Error Resume Next
   If TypeName(appWord) <> "Nothing" Then
'         appWord.Quit
      Set appWord = Nothing
   End If
   xEndTime
   
   'EXIT
   Exit Sub

Proc_Err:

   MsgBox Err.Description, , _
        "ERROR " & Err.Number & "   RunDocumentor"
   Resume Proc_exit
   Resume
End Sub

'~~~~~~~~~~~~~~~~~~~~~ RunDocumentor

Function RunDocumentor( _
    ByVal pDbDocTable As String _
   , ByVal pQueryDocTable As String _
   , Optional pDbAnalName As String = "" _
   , Optional pSkipMssg As Boolean = False _
   , Optional pDocType As Integer = 1 _
   ) As Long

   'crystal
   'strive4peace2010 at yahoo dot com

   
   'PARAMETERS
   'pDbDocTable -- name of table in current db to write Db files that are analyzed
   'pQueryDocTable -- name of table in current db to write results to
   'pDbAnalName -- path and filename of database to analyze
   '             if not specified, currentDb is used
   'pSkipMssg -- if anything is passed, no msgbox will display at end
  
   'NOTE: pDbDocTable and pQueryDocTable
   'will be created
   'if they don't exist
   
   'RETURNS
   'value of DbID
   
   RunDocumentor = 0
   
   Dim dbCur As DAO.Database _
      , dbAnal As DAO.Database _
      , qdf As DAO.QueryDef _
      , frm As Form _
      , rpt As Report _
      , obj As Object _
      , R As DAO.Recordset
      
   Dim i As Integer _
      , s As String _
      , mDbPath As String _
      , mDbName As String _
      , nDbID As Long
      
   
   Set dbCur = DBEngine(0)(0) 'CurrentDB

   'close and delete the table
   'ignore error to close if table is not open
  
   On Error Resume Next
   'this is the only line that uses the currentdb regardless
   DoCmd.Close acTable, pQueryDocTable
   On Error GoTo Proc_Err

   ' document the queries
  
   If Len(Trim(pDbAnalName)) = 0 Then
      Set dbAnal = CurrentDb
      mDbPath = CurrentProject.Path
      mDbName = Dir(CurrentDb.name)
   Else
      mDbName = Dir(pDbAnalName)
      mDbPath = left(pDbAnalName _
         , Len(pDbAnalName) - Len(mDbName))
      If mDbName = "" Then
         MsgBox pDbAnalName & vbCrLf _
            & " is not a valid database"
         GoTo Proc_exit
      End If
      Set dbAnal = OpenDatabase(pDbAnalName)
   End If
  
   If right(mDbPath, 1) <> "\" Then mDbPath = mDbPath & "\"
  
   s = "INSERT INTO [" & pDbDocTable & "]" _
      & "( dbPath, dbName, aTypInt )" _
      & " SELECT '" & mDbPath & "'" _
      & ", '" & mDbName & "', " & pDocType & ";"
      
   xRSql s, "Append to " & pDbDocTable
   
   dbCur.TableDefs.Refresh
   DoEvents
   
   nDbID = Nz(DMax("DbID", pDbDocTable), 0)
   
   If nDbID = 0 Then
      MsgBox "Record was not written to " & pDbDocTable _
         , , "Aborting Query Documentation..."
      Exit Function
   End If
   
   'open table to document query SQL or form/report RowSource
   Set R = dbCur.OpenRecordset( _
      pQueryDocTable _
      , dbOpenDynaset)
   
   'count how many queries are found -- nothing yet
   i = 0
   
   Select Case pDocType
   
   '--------------------------------- Queries
   Case 1: 'QUERIES
      'loop through all the queries
      For Each qdf In dbAnal.QueryDefs
     
         i = i + 1
         
         'skip temporary queries
         If left(qdf.name, 1) = "~" Then _
            GoTo NextTableQ
            
         'add a new record
         'and fill the fields
          
         R.AddNew
         R!dbid = nDbID
         R!objNm = qdf.name
         '110712
         R.Update
         R.Bookmark = R.LastModified
'         On Error Resume Next
         R.Edit
         R!objSrc = qdf.sql
         R.Update
'         On Error GoTo Proc_Err
          
NextTableQ:
      Next qdf

       
   '--------------------------------- Forms
   Case 2: 'FORMS
      'loop through all the forms
      
      For Each obj In CurrentProject.AllForms
      
         DoCmd.OpenForm obj.name, acViewDesign
         
         Set frm = Forms(obj.name)
         
         'add a new record
         'and fill the fields
          
         R.AddNew
         R!dbid = nDbID
         R!objNm = obj.name
         R!objSrc = frm.RecordSource
                  
         R.Update
            
         DoCmd.Close acForm, obj.name, acSaveNo
   
      Next obj
       
    '--------------------------------- Reports
    Case 3: 'REPORTS
      'loop through all the reports
   
      For Each obj In CurrentProject.AllReports
      
         DoCmd.OpenReport obj.name, acViewDesign
         
         Set rpt = Reports(obj.name)
         
         'add a new record
         'and fill the fields
          
         R.AddNew
         R!dbid = nDbID
         R!objNm = obj.name
                        'fixed 120322
            R!objSrc = rpt.RecordSource
         
         R.Update
            
         DoCmd.Close acReport, obj.name, acSaveNo
   
      Next obj
         
   End Select
       
   'look at all the analysis
'   DoCmd.OpenTable pQueryDocTable
   
   If Not pSkipMssg Then
      MsgBox "Documented " & i & " queries" _
      , , "Done"
   End If
   
   RunDocumentor = nDbID
   
Proc_exit:
   On Error Resume Next
   'close and release object variables

   R.Close
   Set R = Nothing
   Set obj = Nothing
   Set frm = Nothing
   Set rpt = Nothing
   Set qdf = Nothing
   
   Set dbCur = Nothing
   
   If Len(Trim(pDbAnalName)) > 0 Then
      dbAnal.Close
   End If
   Set dbAnal = Nothing
   
   'EXIT
   Exit Function

Proc_Err:

   MsgBox Err.Description, , _
        "ERROR " & Err.Number & "   RunDocumentor"

   'press F8 to step through code and debug
   'remove next line after debugged
   Stop:    Resume
   Resume Proc_exit

End Function

'~~~~~~~~~~~~~~~~~~~~~ MakeAccessDocTables

Function MakeAccessDocTables( _
   ByVal pDbDocTable As String _
   , ByVal pQueryDocTable As String _
   , Optional ByVal pDocDbName As String = "" _
   ) As Boolean
'Crystal updated 110712
'strive4peace

   'PARAMETERS
   'pDbDocTable -- name of table in current db to write Db files that are analyzed
   'pQueryDocTable -- name of table in current db to write results to
   'pDocDbName -- path and filename of database to store analysis in
   '             if not specified, currentDb is used
   
   On Error GoTo Proc_Err
   
   MakeAccessDocTables = False
   
   Dim Db As DAO.Database _
      , tdf As DAO.TableDef _
      , fld As DAO.Field _
      , idx As DAO.Index

   If Len(Trim(pDocDbName)) > 0 Then
      Set Db = OpenDatabase(pDocDbName)
   Else
      Set Db = CurrentDb
   End If
   
   Db.TableDefs.Refresh

   ' if the table is already there
   ' we do not have to create it
     
   If Not xDoesExist(pDbDocTable) Then
   
      Set tdf = Db.CreateTableDef(pDbDocTable)
         
         'create autonumber PrimaryKey ID field
         Set fld = tdf.CreateField("DbID", dbLong)
         tdf.Fields.Append fld
         fld.Attributes = dbAutoIncrField
              
         'create field to store path of database
         'NOTE: if you have long paths, make this a memo
         Set fld = tdf.CreateField("dbPath", dbText, 255)
         fld.AllowZeroLength = True
         tdf.Fields.Append fld
         
         'create field to store name of database
         'Note: name limit is 100 characters -- you can increase this
         ' This is just the filename, not the full path
         Set fld = tdf.CreateField("dbName", dbText, 100)
         fld.AllowZeroLength = True
         tdf.Fields.Append fld
         
         'create field to store name of the analysis type
         ' 1=Query, 2=Form, 3=Report
         Set fld = tdf.CreateField("aTypInt", dbInteger)
         fld.DefaultValue = "1"
         tdf.Fields.Append fld
         
         'create field to store when the record was created
         Set fld = tdf.CreateField("dtmAdd", dbDate)
         fld.DefaultValue = "=Now()"
         tdf.Fields.Append fld
      
      'append the new table to the collection
      Db.TableDefs.Append tdf
      Db.TableDefs.Refresh
   
      'add primary key index
      With Db.TableDefs(pDbDocTable)
         Set idx = .CreateIndex("PrimaryKey")
         idx.Fields.Append idx.CreateField("DbID")
         idx.Primary = True
         .Indexes.Append idx
      End With
   End If
   
   ' if the table to store query documentation already exists,
   ' we don't have to make it
   If xDoesExist(pQueryDocTable) Then Exit Function
   
   Set tdf = Db.CreateTableDef(pQueryDocTable)
   
     'create dbID field to link to parent table
     Set fld = tdf.CreateField("DbID", dbLong)
     tdf.Fields.Append fld
     fld.DefaultValue = "=Null"
     
     'create autonumber PrimaryKey ID field
     Set fld = tdf.CreateField("objID", dbLong)
     tdf.Fields.Append fld
     fld.Attributes = dbAutoIncrField
          
     'create field to store query name
     Set fld = tdf.CreateField("objNm", dbText, 64)
     tdf.Fields.Append fld
         
     'create field to store SQL
     Set fld = tdf.CreateField("objSrc", dbMemo)
     fld.AllowZeroLength = True
     tdf.Fields.Append fld
   
     'create field to store when the record was created
     Set fld = tdf.CreateField("dtmAdd", dbDate)
     tdf.Fields.Append fld
     fld.DefaultValue = "=Now()"

   'append the new table to the collection
   Db.TableDefs.Append tdf
   Db.TableDefs.Refresh
   
   'add primary key index
   With Db.TableDefs(pQueryDocTable)
      Set idx = .CreateIndex("PrimaryKey")
      idx.Fields.Append idx.CreateField("objID")
      idx.Primary = True
      .Indexes.Append idx
   End With

   'make sure other processes see that
   'the new tables are there
   
   Db.TableDefs.Refresh
   DoEvents
   
   MakeAccessDocTables = True
   
Proc_exit:
   On Error Resume Next
   'close and release object variables
   Set fld = Nothing
   Set idx = Nothing
   Set tdf = Nothing
   
   If Len(Trim(pDocDbName)) > 0 Then Db.Close
   Set Db = Nothing
   
   'EXIT
   Exit Function

Proc_Err:

   MsgBox Err.Description, , _
        "ERROR " & Err.Number & "   MakeAccessDocTables"

   'press F8 to step through code and debug
   'remove next line after debugged
   Stop:    Resume
   Resume Proc_exit

End Function

'~~~~~~~~~~~~~~~~~~~~~ Word_Write_SQL
  
Function Word_Write_SQL( _
   appWord As Object _
   , pRecordset As String _
   , Optional pDocType As Integer = 1 _
   ) As Boolean
   
   '
   ' crystal
   ' strive4peace2010 at yahoo dot com
   ' 9/6/2006
   '
   'PARAMETERS
   ' appWord is the Word application object
   ' pRecordset is the name of a table or query or an SQL statement
   ' pDocType: 1=Query, 2=Form, 3=Report
   ' pAnalysisType is string corresponding to pDocType
    
   On Error GoTo Proc_Err
   
   Word_Write_SQL = False
   
   Dim rs As DAO.Recordset
   
   Dim tbl As Object
   'Dim tbl As Word.table
   
   Dim mNumRecords As Long _
      , mRow As Long _
      , mAnalysisType As String
   
   Select Case pDocType
      Case 1: mAnalysisType = "Query"
      Case 2: mAnalysisType = "Form"
      Case 3: mAnalysisType = "Report"
   End Select
   
   Set rs = CurrentDb.OpenRecordset( _
      pRecordset _
      , dbOpenSnapshot)
   
   If rs.EOF Then
      MsgBox "No records to write to Word" _
         , , "Aborting Query Documentation"
      GoTo Proc_exit
   End If
   
   rs.MoveLast
   mNumRecords = rs.RecordCount
   
   rs.MoveFirst
   
   With appWord.ActiveDocument
   
      With .PageSetup
         .TopMargin = 72 * (0.75) '72 points/inch
         .BottomMargin = 72 * (0.5)
         .LeftMargin = 72 * (0.75)
         .RightMargin = 72 * (0.5)
         .HeaderDistance = 72 * (0.5)
         .FooterDistance = 72 * (0.5)
      End With
      
      'Add a Word table with:
      '  Two columns
      '  one more row than record for the headings
      '
      Set tbl = appWord.ActiveDocument.Tables.Add( _
         Range:=appWord.Selection.Range _
         , NumRows:=(mNumRecords + 1) _
         , NumColumns:=2)
      
      'format Word table
      With tbl
         ' if using Access 2000 or below
         ' comment the .Style assignment
'         If .Style <> "Table Grid" Then
'            .Style = "Table Grid"
'         End If
         '------------------------------
         .ApplyStyleHeadingRows = True
         .ApplyStyleLastRow = True
         .ApplyStyleFirstColumn = True
         .ApplyStyleLastColumn = True
         
         With .Columns(1) 'Query Name
            .PreferredWidth = 72 * (2)
         End With
         
         With .Columns(2) 'SQL
            .PreferredWidth = 72 * (4.5)
         End With
         
         'header row
         With .Rows(1)
            With .Cells
               With .Shading
                  .Texture = 100 'wdTexture10Percent=100, wdTextureNone = 0
                  .ForegroundPatternColor = 15132390
                  .BackgroundPatternColor = 15132390 ' wdColorGray10= 15132390
               End With
               
               With .Borders(-1)  'wdBorderTop =-1
                  .LineStyle = 1  'wdLineStyleSingle=1
                  .LineWidth = 12 'wdLineWidth150pt=12
                  .Color = 0      'wdColorBlack = 0
               End With
               With .Borders(-2) 'wdBorderLeft = -2
                  .LineStyle = 1
                  .LineWidth = 12
                  .Color = 0
               End With
               With .Borders(-3) 'wdBorderBottom =-3
                  .LineStyle = 1
                  .LineWidth = 12
                  .Color = 0
               End With
               With .Borders(-4) 'wdBorderRight= -4
                  .LineStyle = 1
                  .LineWidth = 12
                  .Color = 0
               End With
               With .Borders(-6) 'wdBorderVertical = -6
                  .LineStyle = 1
                  .LineWidth = 12
                  .Color = 0
               End With
               
NextStatement:
            End With 'cells
         End With 'rows
         
         With .Range
            .Font.name = "Arial"
            .Font.Size = 8
         End With
         
         'heading rows
         .Cell(1, 1).Range.InsertAfter _
            mAnalysisType & " Name"
            
         .Cell(1, 2).Range.InsertAfter _
            IIf(pDocType = 1, "SQL", "RecordSource")
   
         mRow = 1
         
         Do While Not rs.EOF
            mRow = mRow + 1
                      
            ' --- Object Name
            If Len(Trim(Nz(rs!objNm, ""))) > 0 Then
               .Cell(mRow, 1).Range.InsertAfter rs!objNm
            End If
            
            ' --- SQL if Query, RecordSource if Form or Report
            If Len(Trim(Nz(rs!objSrc, ""))) > 0 Then
               .Cell(mRow, 2).Range.InsertAfter rs!objSrc
            End If
            
            rs.MoveNext
         Loop
   
      End With 'table
      
      '3 points before each paragraph
       
      .Content.ParagraphFormat.SpaceBefore = 3
      
   End With 'word application

   Word_Write_SQL = True

Proc_exit:
   On Error Resume Next
   
   rs.Close
   Set rs = Nothing
   
   'EXIT
   Exit Function

Proc_Err:
   If Err.Number = 462 Then
      MsgBox "Try closing the database and opening it again", , "ERROR 462"
      GoTo Proc_exit
   End If
  
   MsgBox Err.Description, , _
        "ERROR " & Err.Number & "   Word_Write_SQL"
  
   'press F8 to step through code and debug
   'remove next line after debugged
   Stop:    Resume
   Resume Proc_exit
   
End Function

'~~~~~~~~~~~~~~~~~~~~~ Word_Format_SQL
  
Sub Word_Format_SQL( _
   pWordApp As Object _
   )
   
   ' Crystal
   ' strive4peace2010 at yahoo dot com
   
   'PARAMETERS
   'pWordApp is object variable for Word
   'assumption:
   '    document to change is the ActiveDocument

   On Error Resume Next
   
    With pWordApp.ActiveDocument
    
      .DefaultTabStop = 72 * (0.2)

        With .Content.Find

            .Execute findText:=" SELECT " _
                     , replaceWith:=" ^lSELECT " _
                     , Replace:=2 'wdReplaceAll = 2

            .Execute findText:=" FROM " _
                     , replaceWith:=" ^lFROM " _
                     , Replace:=2

            .Execute findText:=" IN " _
                     , replaceWith:=" ^lIN " _
                     , Replace:=2

            .Execute findText:=" INTO " _
                     , replaceWith:=" ^lINTO " _
                     , Replace:=2

            .Execute findText:=" WHERE " _
                     , replaceWith:=" ^lWHERE " _
                     , Replace:=2

            .Execute findText:=" GROUP BY " _
                     , replaceWith:=" ^lGROUP BY " _
                     , Replace:=2

            .Execute findText:=" HAVING " _
                     , replaceWith:=" ^lHAVING " _
                     , Replace:=2

            .Execute findText:=" ORDER BY " _
                     , replaceWith:=" ^lORDER BY " _
                     , Replace:=2

            .Execute findText:=" SET " _
                     , replaceWith:=" ^lSET " _
                     , Replace:=2

            .Execute findText:=" ON " _
                     , replaceWith:=" ^l^t^tON " _
                     , Replace:=2

            .Execute findText:=" AND " _
                     , replaceWith:=" ^l^tAND " _
                     , Replace:=2

            .Execute findText:=" OR " _
                     , replaceWith:=" ^l^tOR " _
                     , Replace:=2

            .Execute findText:=" INNER " _
                     , replaceWith:=" ^l^tINNER " _
                     , Replace:=2

            .Execute findText:=" LEFT " _
                     , replaceWith:=" ^l^tLEFT " _
                     , Replace:=2

            .Execute findText:=" RIGHT " _
                     , replaceWith:=" ^l^tRIGHT " _
                     , Replace:=2

            .Execute findText:=", " _
                     , replaceWith:="^l^t, " _
                     , Replace:=2
                     
            'correct previous replacement if comma is at beginning of string literal
            .Execute findText:="'^l^t, " _
                     , replaceWith:="', " _
                     , Replace:=2
            .Execute findText:="""^l^t, " _
                     , replaceWith:=""", " _
                     , Replace:=2
         End With

    End With

End Sub


'================================================================= GENERAL
'procedures here are prefaced with "x" because they are normally
'found in my general libraries and are duplicated here

'------------------------------------ xDoesExist
Function xDoesExist( _
   TName As String _
   , Optional pDb _
   , Optional pTableOnly _
   ) As Boolean
   
   'return TRUE if table or query exists in current (or specified) database
   'example useage: call before Appending records to a table.  If not there, make the table
   ' If not xDoesExist("SummaryTable") then MakeTable "SummaryTable"
   xDoesExist = False
    
   Dim i As Integer
    
   Dim Db As Database
   
   If Not IsMissing(pDb) Then
      Set Db = pDb
   Else
      Set Db = CurrentDb
   End If
   
   For i = 0 To Db.TableDefs.Count - 1
      If Db.TableDefs(i).name = TName Then
         xDoesExist = True
         Exit Function
      End If
   Next i
   
   If Not IsMissing(pTableOnly) Then Exit Function
   
   For i = 0 To Db.QueryDefs.Count - 1
      If Db.QueryDefs(i).name = TName Then
         xDoesExist = True
         Exit Function
      End If
   Next i
   
End Function

'------------------------------------ xStartTime
  
Private Sub xStartTime( _
   Optional pMsg)
    
   On Error Resume Next
   gStartTime = Now()
   DoCmd.Hourglass True
   If IsMissing(pMsg) Then Exit Sub
   Debug.Print "--- START-------------" & pMsg & " ----- " & CStr(gStartTime)
End Sub

'------------------------------------ xReportElapsedTime
  
Private Sub xReportElapsedTime( _
   Optional pMessage As String _
   , Optional pTitle As String)

   On Error Resume Next
   Dim M As String, mEndTime As Date
   mEndTime = Now()
   DoCmd.Hourglass False
   If IsMissing(pMessage) Then
      M = ""
   Else
      M = pMessage & vbCrLf & "-------------" & vbCrLf
      Debug.Print "-------------" & pMessage & " ----- "
   End If
   SysCmd acSysCmdClearStatus
   M = M & "Start Time: " & Format(gStartTime, "hh:nn:ss") & vbCrLf _
      & "End Time: " & Format(mEndTime, "hh:nn:ss") & "     --> " _
      & "     Elapsed Time: " & Format((mEndTime - gStartTime) * 24 * 60, "0") & " minutes"
   MsgBox M, , IIf(IsMissing(pTitle), "Time to execute         ", pTitle)
    
End Sub

'------------------------------------ xEndTime
  
Private Sub xEndTime()
   DoCmd.Hourglass False
   SysCmd acSysCmdClearStatus
   Debug.Print "End " & Format(Now(), "h:nn")
End Sub

'------------------------------------ xRSql
  
Sub xRSql( _
   pSql _
   , Optional pMsg _
   , Optional IsAggregateUpdate As Boolean)

   On Error GoTo Proc_Err
   
   Dim mTime As Date
   mTime = Now()
   
   If Not IsMissing(pMsg) Then
      SysCmd acSysCmdSetStatus, pMsg & "..."
   End If
   Debug.Print pSql
   If Not IsMissing(IsAggregateUpdate) Then
      If IsAggregateUpdate Then
         DoCmd.Echo False
         DoCmd.SetWarnings False
         DoCmd.RunSQL pSql
         DoCmd.Echo True
         DoCmd.SetWarnings True
      Else
         CurrentDb.Execute pSql
      End If
   Else
         CurrentDb.Execute pSql
   End If
   Debug.Print " --- " & Format((Now() - mTime) * 24 * 60 * 60, "#,##0") & " seconds" & " --- "

Proc_exit:
   Exit Sub
    
Proc_Err:
   DoCmd.Echo True
   DoCmd.SetWarnings True
   Resume Proc_exit
   'to see errors in the SQL, look at the debug window
   'if the timing line is missing, SQL did not execute
End Sub

'------------------------------------------ xDeleteFile

Function xDeleteFile( _
   pPathFilename As String _
   , Optional pSayMsg As Boolean = True _
   ) _
   As Boolean
   
   ' CALLS
   '  xWaitMinutes 1
   
   xDeleteFile = False
   
   On Error Resume Next
   Kill pPathFilename
   On Error GoTo Proc_Err
   
   If pSayMsg Then
      SysCmd acSysCmdSetStatus, "Waiting for " & pPathFilename & " to get erased ..."
   End If
   
   Do While Len(Dir(pPathFilename)) > 0
      xWaitMinutes 1
      Kill pPathFilename
      DoEvents
   Loop
    
   If Not Len(Dir(pPathFilename)) > 0 Then
      ' file has been deleted
      xDeleteFile = True
   Else
      If pSayMsg Then
         If MsgBox("Cannot delete --> " _
            & pPathFilename _
            & vbCrLf & vbCrLf _
            & "Do you want to close/delete the file yourself?" _
            , vbYesNo, "Aborting...") = vbNo Then GoTo Proc_exit
      Else
      
               
      End If
      
   End If
   
Proc_exit:
   If pSayMsg Then SysCmd acSysCmdClearStatus

   
   Exit Function
   
Proc_Err:

   'file not found
   If Err.Number = 53 Then
      xDeleteFile = True ' because file is not there
      GoTo Proc_exit
   End If
   
   MsgBox Err.Description, , "ERROR " & Err.Number & "   xDeleteFile"
   
   GoTo Proc_exit
    
   Resume
   
End Function

'------------------------------------------ xWaitMinutes
'this is not efficient -- but we don't need efficiency here
Function xWaitMinutes(pNum As Integer) As Boolean
   xWaitMinutes = False
   Dim mTime As Date, mNumber As Double
   mTime = Now()
   Do
      mNumber = (2 ^ 20 + 34) Mod 36
   Loop Until DateDiff("n", mTime, Now()) = pNum
   xWaitMinutes = True
End Function

'------------------------------------------ xCreate_qObjex
Sub xCreate_qObjex(Optional pBooOpen As Boolean = True)
'110712 Crystal strive4peace

' Create a query with object names and types
   'if the query already exists, it will be overwritten
   'if you made changes you want to keep, change the name.
   'For instance, add something to the end like --> _yymmdd
   Dim strSQL As String _
      , qName As String
      
   qName = "qObjex_byCrystal"
   
   strSQL = "SELECT GetObjectType([Type]) AS ObjectType, MSysObjects.Name, [Type] as Type_" _
      & " FROM MSysObjects " _
      & " WHERE (((Left([Name], 1)) <> ""~"") And ((Left([Name], 4)) <> ""MSys""))" _
      & " ORDER BY GetObjectType([Type]), MSysObjects.Name;"
    xMakeQuery strSQL, qName

   Application.RefreshDatabaseWindow
   'default pBooOpen = true -->
   'assume if the user is making the query, they want to open it.
   If pBooOpen Then DoCmd.OpenQuery qName
   
End Sub

'------------------------------------------ GetObjectType
Function GetObjectType(pType) As String
   Select Case pType
   Case 1: GetObjectType = "Table"
   Case 5: GetObjectType = "Query"
   Case -32768: GetObjectType = "Form"
   Case -32764: GetObjectType = "Report"
   Case -32766: GetObjectType = "Macro"
   Case -32761: GetObjectType = "Module"
   Case Else: GetObjectType = pType
   End Select
End Function

'------------------------------------------ xMakeQuery
Sub xMakeQuery( _
   ByVal pSql As String, _
   ByVal qName As String)

   'modified 3-30-08
   'crystal
   'strive4peace2009 at yahoo dot com

   On Error GoTo Proc_Err

Debug.Print pSql

   'if query already exists, update the SQL
   'if not, create the query
   
    If Nz(DLookup("[Name]", "MSysObjects", _
        "[Name]='" & qName _
        & "' And [Type]=5"), "") = "" Then
        CurrentDb.CreateQueryDef qName, pSql
    Else
       'if query is open, close it
       On Error Resume Next
       DoCmd.Close acQuery, qName, acSaveNo
       On Error GoTo Proc_Err
       CurrentDb.QueryDefs(qName).sql = pSql
    End If
   
Proc_exit:
   CurrentDb.QueryDefs.Refresh
   Application.RefreshDatabaseWindow
   DoEvents
   Exit Sub
   
Proc_Err:
   MsgBox Err.Description, , _
     "ERROR " & Err.Number & "  xMakeQuery"
    
   Resume Proc_exit

   'if you want to single-step code to find error, CTRL-Break at MsgBox
   'then set this to be the next statement
   Resume
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~