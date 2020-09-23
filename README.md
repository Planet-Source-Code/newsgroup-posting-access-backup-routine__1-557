<div align="center">

## ACCESS BACKUP ROUTINE


</div>

### Description

A short routine that backups the tables from an open Access database

George Kinney <kinneyg@logan.net>
 
### More Info
 
Right now it is basic, it assumes that the tables to backup are in the

local database (easily changed, just haven't had a chance to do it.),

and just exports EVERYTHING that isn't filtered out.

A number of improvements can (and will eventually) be built in so that it

can address attached tables, multiple backups, backup logging, etc. These

are all things I need to add anyways for a current project, and when they

are done, I'll b eposting them to.

Apologies are in order to a few of the people I sent code. The function

relied on a couple of outside functions, not included in the post, and

also contained a lot extraneous junk (you who work with large projects full

time know how this stuff accumulates, those who don't, well you'll find

out.). To these people, I'm sorry for that, and hope you don't take me to

be a complete idiot. (3am is a bad time to reply to mail!)

I don't claim to be a programming guru, but I think this example could

benefit some people. I recieved a lot of help from others early on, so I

intend to give what I can as I can so others can hopefully benefit from me.

'Just call BackupDatabase() with the name of the backup file

'you want to create, and sit back.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-access-backup-routine__1-557/archive/master.zip)

### API Declarations

Const modulename = "MBackup"


### Source Code

```
Function BackupDataBase (filename$) As Integer
'**********************************************************************************
'* PROCEDURE: BackupDataBase
'* ARGS:   filename$ -- name of new DataBase, defaults to current Dir
'* RETURNS:  TRUE/FALSE
'* CREATED:  7/95
'* REVISED:  8/2/95 GDK Changed to use the App's dir.
'* Comments  Creates newDataBase, and exports ALL existing tables in the
'*       Current database to it.
'* ToDo:   Backup current backup before writing over it. (part of backup
'*       archive system)
'*       Add new backup logging stuff to this function.(Date, location, etc.)
'**********************************************************************************
On Error GoTo BackupDataBase_Err
  Dim newDB As Database, oldDB As Database, oldTable As TableDef
  Dim tempname As String, path As String, intIndex As Integer, numTables As Integer
  Dim intIndex2 As Integer, errorFlag As Integer
  'backup defaults to current directory,...
  path = GetApplicationDir() & filename$
  'replace above line with this one to pass a full path to this function
  'path = filename$
  'If database already exists, delete it.
  If MB_FileExists(path) Then
    Kill path
  End If
  'create new file
  Set newDB = DBEngine.workspaces(0).CreateDatabase(path, DB_LANG_GENERAL)
  newDB.Close
  Set oldDB = DBEngine(0)(0)
  'Get number of tables and their names
  numTables = oldDB.tabledefs.count - 1
  'Actually export all the tables in the list.
  For intIndex = 0 To numTables
    tempname = oldDB.tabledefs(intIndex).name
    If ValidTableFilter(tempname) Then
      DoCmd TransferDatabase A_EXPORT, "Microsoft Access", path, A_TABLE, tempname, tempname
    End If
  Next intIndex
  BackupDataBase = True
BackupDataBase_Exit:
  If errorFlag Then
    BackupDataBase = False
    'if we errored out, then destroy the backup, (less risk of using incorrect file).
    If MB_FileExists(path) Then
      Kill path
    End If
  Else
    BackupDataBase = True
  End If
  Exit Function
BackupDataBase_Err:
  MsgBox "Backup Failed! Error: " & Error$, 16, "FUNCTION: BackupDataBase( " & filename$ & " )"
  errorFlag = True
  Resume BackupDataBase_Exit
End Function
Function GetApplicationDir () As String
'***************************************************************************
'* PROCEDURE: GetApplicationDir
'* ARGS:   NONE
'* RETURNS:  App's dir
'* CREATED:  8/2/95 GDK
'* REVISED:
'* Comments  Retrieves App's directory, (actually the current MDB's dir.)
'***************************************************************************
  Dim d As Database, path As String, i%
  Set d = DBEngine(0)(0)
    path = d.name
  d.Close
  For i% = Len(path) To 0 Step -1
    If Mid$(path, i%, 1) = "\" Then
      path = Left$(path, i%)
      Exit For
    End If
  Next i%
  GetApplicationDir = path
End Function
'*************************************************************
'* FUNCTION: MB_FileExists
'* ARGUMENTS: strFilename  -- name of file to look for
'* RETURNS:  TRUE/FALSE   -- TRUE = File Exists
'* CREATED:  8/95 GDK Initial Code
'* CHANGED:  N/A
'*************************************************************
Function MB_FileExists (strFileName As String) As Integer
'
'Check to see if file strFileName exists
'
  If Len(Dir$(strFileName)) Then
    MB_FileExists = True
  End If
End Function
'***************************************************************
'* FUNCTION: ValidTableFilter
'* ARGUMENTS: tablename$ -- table to OK for export
'* RETURNS:  TRUE/FALSE -- TRUE = OK to export
'* PURPOSE:  Screen out invalid tables by testing them here.
'* CREATED:  2/97 GDK Initial code
'* CHANGES:  N/A
'***************************************************************
Function ValidTableFilter (tablename$) As Integer
On Error GoTo ValidTableFilter_Error:
  If Left$(tablename$, 4) = "MSys" Then
    Exit Function
  End If
  If tablename$ = "" Then
    Exit Function
  End If
  'Add test functions above this line.
  ValidTableFilter = True
ValidTableFilter_Exit:
  Exit Function
ValidTableFilter_Error:
  MsgBox Error, 16, "FUNCTION: ValidTableFilter( " & tablename$ & ")"
  Resume ValidTableFilter_Exit
End Function
```

