# Script Get fields list

I've developed a .vbs script that will scan a MS Access database, loop for each tables and, for each of them, will get the list of fields.

For each fields, a lot of information's will be retrieved like, not exhaustive, his name, size, type, ... and also the shortest and longest value size (for text and memo fields). For instance, if a text field is found, the script will retrieve his size (f.i. 255 chars max) and will examine all records in the table for retrieving, for that field, the smallest size (f.i. 10) and the greatest one (f.i. 50). So, if the max size is 50 and the size has been set to 255, perhaps the MS Access developer can safely modify the max size from 255 to 50.

To make the script to run:

1. Copy/paste the source code below, one by one, and save it to a text file (with Notepad). The first file to create will be `access_get_fields_list.vbs`, the second will be `access_get_fields_list.cmd` (see after),
2. Before saving the `access_get_fields_list.cmd` be sure to edit the file and mention the full filename of your database (see after),
3. You're ready, from your File Explorer, just double-clic on the `access_get_fields_list.cmd`, the analyse script will be executed and Excel will be opened at the end.

*If everything goes fine, you'll see a DOS window and after a few seconds (depending on the size and complexity of the database), you'll have the report in Excel, automatically opened.*

## Prepare files

### access_get_fields_list.vbs

Copy the source code below in the clipboard, start Notepad and paste the lines. Save the file onto your hard disk and give `access_get_fields_list.vbs` as filename. You can then quit Notepad.

```vbnet
' ====================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' Open a database, get the list of tables and for each of them,
' get the list of fields and a few properties like the type, the size,
' the shortest and longest value size (for text and memo fields)
'
' The output will be something like :
' Database;TableName;FieldName;FieldType;FieldSize;ShortestSize;LongestSize;Position;Occurences;
' C:\Temp\db1.accdb;Bistel;RefDate;Date/Time;8;;1;1
' C:\Temp\db1.accdb;Bistel;BudgetType;Byte;1;;2;1
' C:\Temp\db1.accdb;Bistel;OrganicDivision;Text (fixed width);2;;3;1
' C:\Temp\db1.accdb;Bistel;Program;Text (fixed width);1;;4;1
' C:\Temp\db1.accdb;Bistel;Published;Yes/No;1;;5;1
' C:\Temp\db1.accdb;Bistel;DescriptionDutch;Text;50;10;48;6;1
' C:\Temp\db1.accdb;Bistel;DescriptionFrench;Text;50;0;50;7;1
' C:\Temp\db1.accdb;Bistel;Article;Text;6;6;6;8;1
' C:\Temp\db1.accdb;departements;bud;Text;255;2;2;1;1
'
' Changes
' =======
'
' March 2018 - Make this script stand alone by including MS Access and
'             MS Excel classes in the script
'
' ====================================================================

Option Explicit

Class clsMSAccess

    Private oApplication
    Private bVerbose

    Private sDatabaseName
    Private sDelim

    Public Property Let verbose(bYesNo)
        bVerbose = bYesNo
    End Property

    Public Property Let DatabaseName(ByVal sFileName)
        sDatabaseName = sFileName
    End Property

    ' Define the delimiter to use for the CSV file (; or , or ...)
    Public Property Let CSVDelimiter(ByVal sDelimiter)
        sDelim = sDelimiter
    End Property

    Private Sub Class_Initialize()

        bVerbose = False
        sDatabaseName = ""
        sDelim = ";"

        Set oApplication = Nothing

    End Sub

    Private Sub Class_Terminate()
        If Not (oApplication Is Nothing) Then
            oApplication.Quit
            Set oApplication = Nothing
        End If
    End Sub

    ' Verify that databases mentionned in the arrDBNames are well
    ' present and accessible to the user. Return false otherwise
    Private Function CheckIfFilesExists(ByRef arrDBNames)

        Dim objFSO
        Dim bReturn
        Dim i, iMin, iMax

        iMin = LBound(arrDBNames)
        iMax = UBound(arrDBNames)
        bReturn = True

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        iMin = LBound(arrDBNames)
        iMax = UBound(arrDBNames)

        For i = iMin To iMax

            If Not (objFSO.FileExists(arrDBNames(I))) Then
                bReturn = False
                wScript.echo "ERROR - clsMSAccess::CheckIfFilesExists - " & _
                    "File " & arrDBNames(I) & " not found " & _
                    "(clsMSAccess::CheckIfFilesExists)"
            End if

        Next

        CheckIfFilesExists = bReturn

    End function

    ' -----------------------------------------------------------
    ' FieldTypeName
    ' by Allen Browne, allen@allenbrowne.com. Updated June 2006.
    ' copied from http://allenbrowne.com/func-06.html
    ' (No license information found at that URL.)
    ' -----------------------------------------------------------
    Private Function GetFieldTypeName(FieldType, FieldAttributes)

        Dim sReturn

        Select Case CLng(FieldType)
            Case 1: sReturn = "Yes/No"                    ' dbBoolean
            Case 2: sReturn = "Byte"                    ' dbByte
            Case 3: sReturn = "Integer"                    ' dbInteger
            Case 4                                        ' dbLong
                If (FieldAttributes And 17) = 0 Then    ' dbAutoIncrField
                    sReturn = "Long Integer"
                Else
                    sReturn = "AutoNumber"
                End If
            Case 5: sReturn = "Currency"                ' dbCurrency
            Case 6: sReturn = "Single"                    ' dbSingle
            Case 7: sReturn = "Double"                     ' dbDouble
            Case 8: sReturn = "Date/Time"                ' dbDate
            Case 9: sReturn = "Binary"                     ' dbBinary
            Case 10                                         ' dbText
                If (FieldAttributes And 1) = 0 Then     ' dbFixedField
                    sReturn = "Text"
                Else
                    sReturn = "Text (fixed width)"        ' (no interface)
                End If
            Case 11: sReturn = "OLE Object"                 ' dbLongBinary
            Case 12                                         ' dbMemo
                If (FieldAttributes And 32768) = 0 Then ' dbHyperlinkField
                    sReturn = "Memo"
                Else
                    sReturn = "Hyperlink"
                End If
            Case 15: sReturn = "GUID"                     'dbGUID
            'Attached tables only: cannot create these in JET.
            Case 16: sReturn = "Big Integer"            ' dbBigInt
            Case 17: sReturn = "VarBinary"                ' dbVarBinary
            Case 18: sReturn = "Char"                    ' dbChar
            Case 19: sReturn = "Numeric"                ' dbNumeric
            Case 20: sReturn = "Decimal"                ' dbDecimal
            Case 21: sReturn = "Float"                    ' dbFloat
            Case 22: sReturn = "Time"                     ' dbTime
            Case 23: sReturn = "Time Stamp"                  ' dbTimeStamp
            Case Else: sReturn = "Field type " & fld.Type & " unknown"
        End Select

        GetFieldTypeName = sReturn

    End Function

    ' -----------------------------------------------------------
    ' Open the database
    ' -----------------------------------------------------------
    Public Sub OpenDatabase()

        If (oApplication is Nothing) then
            Set oApplication = CreateObject("Access.Application")
            oApplication.Visible = True
        End if

        If (Right(sDatabaseName,4) = ".adp") Then
            oApplication.OpenAccessProject sDatabaseName
        Else
            oApplication.OpenCurrentDatabase sDatabaseName
        End If

    End Sub

    ' -----------------------------------------------------------
    ' Close the database
    ' -----------------------------------------------------------
    Public Sub CloseDatabase()

        If Not (oApplication is Nothing) then
            oApplication.CloseCurrentDatabase
        End if

    End Sub

    ' -----------------------------------------------------------
    '
    ' Scan one or severall MS Access databases, retrieve the list
    ' of tables in these DBs and get the list of fields plus some
    ' properties like the size and, for text fields, the shortest size
    ' and the longest one.
    '
    ' @arrDBNames : array - Contains the list of databases to scan
    '
    ' Example =
    '
    '    arr(0) = "c:\temp\db1.accdb"
    '    arr(1) = "c:\temp\db2.accdb"
    '    arr(2) = "c:\temp\db3.accdb"
    '
    '    wScript.echo GetFieldsList(arr)
    '
    ' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#getfieldslist
    '
    ' -----------------------------------------------------------
    Public Function GetFieldsList(ByRef arrDBNames)

        Dim i, iMin, iMax, sShortest, sLargest, wPos, wRow, wFieldsCount
        Dim sSQL, sReturn, sTableName, sType, sFormulaOccurences, sFormula
        Dim objTable, objField, rs

        If bVerbose Then
            wScript.echo vbCrLf & "=== clsMSAccess::GetFieldsList ===" & vbCrLf
        End If

        If IsArray(arrDBNames) Then

            ' Before starting, just verify that files exists
            ' If no, show an error message and stop
            If CheckIfFilesExists(arrDBNames) Then

                ' Ok, database(s) are well present, we can start
                sReturn = "Filename;TableName;FieldName;FieldType; " & _
                    "FieldSize;ShortestSize;LongestSize;Position;Occurences" & vbCrLf

                sFormulaOccurences = "=COUNTIFS($B$2:$B$@COUNT@,B@ROW@,$C$2:$C$@COUNT@,C@ROW@)"

                wRow = 1
                iMin = LBound(arrDBNames)
                iMax = UBound(arrDBNames)

                For i = iMin To iMax

                    If bVerbose Then
                        wScript.echo "Process " & arrDBNames(I) & " " & _
                            "(clsMSAccess::GetFieldsList)"
                    End If

                    sDatabaseName = arrDBNames(I)
                    Call OpenDatabase()

                    oApplication.CurrentDB.TableDefs.Refresh

                    For each objTable In oApplication.CurrentDB.TableDefs
                        sTableName = objTable.Name

                        wPos = 0

                        ' Ignore system and temporary tables
                        If (lcase(Left(sTableName, 4))<>"msys") And (Left(sTableName, 1) <> "~") Then

                            If bVerbose Then
                                wScript.echo "    Get list of fields of [" & _
                                    sTableName & "]"
                            End If

                            ' Get the number of fields in the table
                            wFieldsCount = objTable.Fields.Count

                            For Each objField In objTable.Fields

                                wPos = wPos + 1
                                wRow = wRow + 1

                                If bVerbose Then
                                    wScript.echo "      " & wPos & "/" & _
                                        wFieldsCount & " - " & _
                                        "Field [" & _
                                        objField.Name & "]"
                                End If

                                sShortest = ""
                                sLargest = ""

                                sType = GetFieldTypeName(objField.Type, objField.Attributes)

                                If (sType = "Text") Or (sType = "Memo") Then

                                    sSQL = "SELECT " & _
                                        "Min(Len([" & objField.Name & "])) As Min, " & _
                                        "Max(Len([" & objField.Name & "])) As Max " & _
                                        "FROM [" & sTableName & "]"

                                    Set rs = oApplication.CurrentDB.OpenRecordset(sSQL, 4)
                                    sShortest = rs.Fields("Min").Value
                                    sLargest = rs.Fields("Max").Value
                                     rs.Close
                                     Set rs = Nothing

                                End If

                                sFormula = replace(sFormulaOccurences, "@ROW@", wRow)

                                sReturn = sReturn & _
                                    arrDBNames(I) & sDelim & _
                                    sTableName & sDelim & _
                                    objField.Name & sDelim & _
                                    sType & sDelim & _
                                    objField.Size & sDelim & _
                                    sShortest & sDelim & _
                                    sLargest & sDelim & _
                                    wPos & sDelim & _
                                    sFormula & vbCrLf

                            Next

                        End if

                    Next

                    Call CloseDatabase

                Next

                sReturn = Replace(sReturn, "@COUNT@", wRow)

            End IF

        Else

            wScript.echo "ERROR - clsMSAccess::GetFieldsList - " & _
                "You must provide an array with filenames. " & _
                "(clsMSAccess::GetFieldsList)"

        End If

        GetFieldsList = sReturn

    End Function

End Class

Class clsMSExcel

    Private oApplication
    Private sFileName
    Private bVerbose, bEnableEvents, bDisplayAlerts

    Private bAppHasBeenStarted

    Public Property Let verbose(bYesNo)
        bVerbose = bYesNo
    End Property

    Public Property Let EnableEvents(bYesNo)
        bEnableEvents = bYesNo

        If Not (oApplication Is Nothing) Then
            oApplication.EnableEvents = bYesNo
        End if
    End Property

    Public Property Let FileName(ByVal sName)
        sFileName = sName
    End Property

    Public Property Get FileName
        FileName = sFileName
    End Property

    Private Sub Class_Initialize()
        bVerbose = False
        bAppHasBeenStarted = False
        bEnableEvents = False
        bDisplayAlerts = False
        Set oApplication = Nothing
    End Sub

    Private Sub Class_Terminate()
        Set oApplication = Nothing
    End Sub

    ' --------------------------------------------------------
    ' Initialize the oApplication object variable : get a pointer
    ' to the current Excel.exe app if already in memory or start
    ' a new instance.
    '
    ' If a new instance has been started, initialize the variable
    ' bAppHasBeenStarted to True so the rest of the script knows
    ' that Excel should then be closed by the script.
    ' --------------------------------------------------------
    Public Function Instantiate()

        If (oApplication Is Nothing) Then

            On error Resume Next

            Set oApplication = GetObject(,"Excel.Application")

            If (Err.number <> 0) or (oApplication Is Nothing) Then
                Set oApplication = CreateObject("Excel.Application")
                ' Remember that Excel has been started by
                ' this script ==> should be released
                bAppHasBeenStarted = True
            End If

            oApplication.EnableEvents = bEnableEvents
            oApplication.DisplayAlerts = bDisplayAlerts

            Err.clear

            On error Goto 0

        End If

        ' Return True if the application was created right
        ' now
        Instantiate = bAppHasBeenStarted

    End Function

    ' --------------------------------------------------------
    ' Be sure Excel is visible
    ' --------------------------------------------------------
    Public Sub MakeVisible

        Dim objShell

        If Not (oApplication Is Nothing) Then

            With oApplication

                .Application.ScreenUpdating = True
                .Application.Visible = True
                .Application.DisplayFullScreen = False

                .WindowState = -4137 ' xlMaximized

            End With

            Set objShell = CreateObject("WScript.Shell")
            objShell.appActivate oApplication.Caption
            Set objShell = Nothing

        End If

    End Sub

    Public Sub Quit()
        If not (oApplication Is Nothing) Then
            oApplication.Quit
        End If
    End Sub

    ' --------------------------------------------------------
    ' Open a CSV file, correctly manage the split into columns,
    ' add a title, rename the tab
    '
    ' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md#opencsv
    ' --------------------------------------------------------
    Public Sub OpenCSV(sTitle, sSheetCaption)

        Dim objFSO
        Dim wCol

        If bVerbose AND (sFileName = "") Then
            wScript.echo "Error, you need to initialize the " & _
                "filename first", " (clsMSExcel::OpenCSV)"
            Exit sub
        End If

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        If (objFSO.FileExists(sFileName)) Then

            If bVerbose Then
                wScript.echo "Open " & sFileName & _
                    " (clsMSExcel::OpenCSV)"
            End If

            If (oApplication Is Nothing) Then
                Call Instantiate()
            End If

            ' 1 =  xlDelimited
            ' Delimiter is ";"
            oApplication.Workbooks.OpenText sFileName,,,1,,,,,,,True,";"

            ' If a title has been specified,
            ' add quickly a small templating
            If (Trim(sTitle) <> "") Then

                With oApplication.ActiveSheet

                    ' Get the number of colunms in the file
                    wCol = .Range("A1").CurrentRegion.Columns.Count

                    .Range("1:3").insert
                    .Range("A2").Value = Trim(sTitle)

                    With .Range(.Cells(2, 1), .Cells(2, wCol))
                        ' 7 = xlCenterAcrossSelection
                        .HorizontalAlignment = 7
                        .font.bold = True
                        .font.size = 14
                    End with

                    .Cells(4,1).AutoFilter

                    .Columns.EntireColumn.AutoFit

                    .Cells(5,1).Select

                End with

                oApplication.ActiveWindow.DisplayGridLines = False
                oApplication.ActiveWindow.FreezePanes = true

            End If

            If (Trim(sSheetCaption) <> "") Then
                oApplication.ActiveSheet.Name = sSheetCaption
            End If

        End If

    End Sub

End Class

Sub ShowHelp()

    wScript.echo " =========================================="
    wScript.echo " = Scan for fields in MS Access databases ="
    wScript.echo " =========================================="
    wScript.echo ""
    wScript.echo " Please specify the name of the database to scan; f.i. : "
    wScript.echo " " & Wscript.ScriptName & " 'C:\Temp\db1.accdb'"
    wScript.echo ""

    wScript.echo "To get more info, please read https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#getfieldslist"
    wScript.echo ""

    wScript.quit

End sub

Dim cMSAccess, cMSExcel
Dim arrDBNames(0)
Dim sFieldsList, sFileName, sFile
Dim objFSO, objFile, oShell

    ' Get the first argument (f.i. "C:\Temp\db1.accdb")
    If (wScript.Arguments.Count = 0) Then

        Call ShowHelp

    Else

        ' Get the path specified on the command line
        sFile = Wscript.Arguments.Item(0)

        Set cMSAccess = New clsMSAccess

        cMSAccess.Verbose = True

        arrDBNames(0) = sFile

        ' Get the list of fields for each table in the
        ' specified databases
        sFieldsList = cMSAccess.GetFieldsList(arrDBNames)

        Set cMSAccess = Nothing

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        ' Finally, output the list into a flatfile and open it
        sFileName = objFSO.GetSpecialFolder(2) & "\output.csv"

        Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
        objFile.Write sFieldsList
        objFile.Close
        Set objFile = Nothing

        Set cMSExcel = New clsMSExcel
        cMSExcel.FileName = sFileName
        cMSExcel.Verbose = True
        cMSExcel.OpenCSV sFile & " - Field lists", "fields"
        Call cMSExcel.MakeVisible
        Set cMSExcel = Nothing

    End if
```

### access_get_fields_list.cmd

Copy the source code below in the clipboard, start Notepad and paste the lines.

**Change the file name and mention the fullname of the database to analyse** (For instance, `c:\databases\my_db.accdb`)

Save the file onto your hard disk and give `access_get_fields_list.cmd` as filename. You can then quit Notepad.

```vbnet
cscript access_get_fields_list.vbs "C:\temp\my_db.accdb" //nologo
```

## Understand the report

Once the process is finished, Excel will be automatically fired with something like this:

![MSAccess Get fields list](./images/get_fields_list.png)

* Filename: The MS Access filename (absolute)
* TableName: The name of the table
* FieldName: The name of the field found in that table
* FieldType: The data type (integer, string, date, ...)
* FieldSize: The maximum size defined in the table (f.i. 255 means that this field can contains up to 255 characters)
* ShortestSize: When the table contains records, the ShortestSize info is "what is the smaller information stored in that field?" (example: if the field is a firstname, size 255 but the shortest firstname is `Paul`, then `ShortestSize` will be set to 4)
* LongestSize: When the table contains records, the `LongestSize` info is "what is the biggest information stored in that field?" (example: if the field is a firstname, size 255 but the longest firstname is `Christophe`, then `LongestSize` will be set to 10)
* Position: The position of that field in the structure of the table (is the first defined field, the second, ...)
* Occurences: How many time, that specific `FieldName` is found in the entire database. If you've a lot of tables, perhaps the field called `CustomerID` is used in the `Customers` table and in the `Orders` table too so `Occurences` will be set to 2 in this case.

In a context of optimization:

* Be sure to not have too big fieldsize. By default, MS Access suggest a size of 255 for text fields but for name and firstname a size of 40 characters is enough.
* Check the `LongestSize` property: if you see f.i. a size of 4 this means that you're probably storing a code (a ZipCode f.i. is max 4 digits in Belgium). If the `FieldSize` is set to 50, you know you can reduce that size to 4.
