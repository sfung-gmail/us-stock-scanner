Option Explicit

Const msoFalse = 0
Const msoTrue = -1

Const ppAutoSizeNone = 0
Const msoAutoSizeNone = 0

Const MAX_CHARS_PER_SLIDE = 140
Const MAX_LINES_PER_SLIDE = 4

Const MAX_SEARCH_YEAR = 2100
Const MIN_SEARCH_YEAR = 2010

Const FILE_ATTRIBUTE_REPARSE_POINT = 1024

Dim fso, shell, rootFolder, autorunFolder, copyFolder
Dim outputFolder, folderName

Dim masterTemplate, refreshScript
Dim readingTemplate, scripturesTemplate, callResponseTemplate, songShortcut
Dim readingTextFile, scripturesTextFile, callResponseTextFile
Dim readingOutputFile, scripturesOutputFile, callResponseOutputFile

Dim songsInput, songNames, songName
Dim songRootPath, songFilePath
Dim missingSongs, i

Dim ppt, errText

On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If Err.Number <> 0 Or fso Is Nothing Or shell Is Nothing Then
    MsgBox "Unable to start Windows scripting components.", _
           vbCritical, "Create Worship Folder"
    WScript.Quit 1
End If

Err.Clear
On Error GoTo 0

rootFolder = fso.GetParentFolderName(WScript.ScriptFullName)
autorunFolder = fso.BuildPath(rootFolder, "autorun")
copyFolder = fso.BuildPath(autorunFolder, "worship-files")

readingTextFile = CNRead() & ".txt"
scripturesTextFile = CNScriptures() & ".txt"
callResponseTextFile = CNCallResponse() & ".txt"

readingOutputFile = CNRead() & ".pptx"
scripturesOutputFile = CNScriptures() & ".pptx"
callResponseOutputFile = CNCallResponse() & ".pptx"

readingTemplate = fso.BuildPath(autorunFolder, CNRead() & "-template.pptx")
scripturesTemplate = fso.BuildPath(autorunFolder, CNScriptures() & "-template.pptx")
callResponseTemplate = fso.BuildPath(autorunFolder, CNCallResponse() & "-template.pptx")

masterTemplate = fso.BuildPath(copyFolder, "master.pptx")
refreshScript = fso.BuildPath(copyFolder, "refresh-master.vbs")

songShortcut = fso.BuildPath(autorunFolder, "songs-folder.lnk")

' ---------------------------------------------------------------
' Validate required structure before asking the user for input.
' ---------------------------------------------------------------

If Not fso.FolderExists(autorunFolder) Then
    Fail MsgAutoRunMissing()
End If

If Not fso.FolderExists(copyFolder) Then
    Fail MsgCopyFolderMissing()
End If

If Not fso.FileExists(masterTemplate) Then
    Fail MsgFileMissing("autorun\worship-files\master.pptx")
End If

If Not fso.FileExists(refreshScript) Then
    Fail MsgFileMissing("autorun\worship-files\refresh-master.vbs")
End If

If Not fso.FileExists(readingTemplate) Then
    Fail MsgFileMissing("autorun\" & CNRead() & "-template.pptx")
End If

If Not fso.FileExists(scripturesTemplate) Then
    Fail MsgFileMissing("autorun\" & CNScriptures() & "-template.pptx")
End If

If Not fso.FileExists(callResponseTemplate) Then
    Fail MsgFileMissing("autorun\" & CNCallResponse() & "-template.pptx")
End If

If Not fso.FileExists(songShortcut) Then
    Fail MsgFileMissing("autorun\songs-folder.lnk")
End If

If Not fso.FileExists(fso.BuildPath(rootFolder, readingTextFile)) Then
    Fail MsgFileMissing(readingTextFile)
End If

If Not fso.FileExists(fso.BuildPath(rootFolder, scripturesTextFile)) Then
    Fail MsgFileMissing(scripturesTextFile)
End If

If Not fso.FileExists(fso.BuildPath(rootFolder, callResponseTextFile)) Then
    Fail MsgFileMissing(callResponseTextFile)
End If

' ---------------------------------------------------------------
' Ask for new folder name.
' ---------------------------------------------------------------

folderName = InputBox(MsgEnterFolderName(), MsgAppTitle())
folderName = Trim(folderName)

If folderName = "" Then
    WScript.Quit 0
End If

If Not IsValidFolderName(folderName) Then
    Fail MsgInvalidFolderName()
End If

outputFolder = fso.BuildPath(rootFolder, folderName)

If fso.FolderExists(outputFolder) Then
    Fail MsgFolderExists() & vbCrLf & outputFolder & vbCrLf & vbCrLf & _
         MsgNoChanges()
End If

' ---------------------------------------------------------------
' Ask for requested songs.
' ---------------------------------------------------------------

songsInput = InputBox( _
    MsgEnterSongs() & vbCrLf & MsgSongExample(), _
    MsgAppTitle() _
)

If Len(songsInput) = 0 Then
    WScript.Quit 0
End If

songNames = Split(songsInput, ",")

songRootPath = GetShortcutTarget(songShortcut)

If songRootPath = "" Or Not fso.FolderExists(songRootPath) Then
    Fail MsgSongRootMissing()
End If

' ---------------------------------------------------------------
' Start PowerPoint.
' ---------------------------------------------------------------

On Error Resume Next

Set ppt = GetObject(, "PowerPoint.Application")

If Err.Number <> 0 Then
    Err.Clear
    Set ppt = CreateObject("PowerPoint.Application")
End If

If Err.Number <> 0 Or ppt Is Nothing Then
    errText = Err.Description
    Err.Clear
    On Error GoTo 0

    Fail MsgCannotStartPowerPoint() & vbCrLf & errText
End If

On Error GoTo 0

' ---------------------------------------------------------------
' Create output folder and copy direct fixed files.
' ---------------------------------------------------------------

errText = ""

On Error Resume Next

fso.CreateFolder outputFolder

If Err.Number <> 0 Then
    errText = Err.Description
    Err.Clear
Else
    CopyDirectFiles copyFolder, outputFolder, errText
End If

On Error GoTo 0

If errText <> "" Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCopyFixedFiles() & errText
End If

' ---------------------------------------------------------------
' Create call-and-response PPTX file.
' Output naming:
'   02_宣召及啟應.pptx
'
' Template Selection Pane names required:
'   Slide 1: CALL_SCRIPTURE_1
'   Slide 2: CALL_1 and RESPONSE_1
' ---------------------------------------------------------------

If Not CreateCallResponsePpt( _
    ppt, _
    callResponseTemplate, _
    fso.BuildPath(rootFolder, callResponseTextFile), _
    fso.BuildPath(outputFolder, "02_" & callResponseOutputFile), _
    errText _
) Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCreateFile("02_" & callResponseOutputFile) & _
         vbCrLf & errText
End If

' ---------------------------------------------------------------
' Create scripture PPTX files.
' Output naming:
'   30_經訓.pptx
'   40_讀經.pptx
' ---------------------------------------------------------------

If Not CreateScripturePpt( _
    ppt, _
    readingTemplate, _
    fso.BuildPath(rootFolder, readingTextFile), _
    fso.BuildPath(outputFolder, "40_" & readingOutputFile), _
    errText _
) Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCreateFile("40_" & readingOutputFile) & _
         vbCrLf & errText
End If

If Not CreateScripturePpt( _
    ppt, _
    scripturesTemplate, _
    fso.BuildPath(rootFolder, scripturesTextFile), _
    fso.BuildPath(outputFolder, "30_" & scripturesOutputFile), _
    errText _
) Then
    CleanupFolder outputFolder
    SafeQuitPowerPoint ppt

    Fail MsgCannotCreateFile("30_" & scripturesOutputFile) & _
         vbCrLf & errText
End If

' ---------------------------------------------------------------
' Copy song PPTX files.
' Output naming:
'   10_<song title>.pptx
'
' Missing songs receive:
'   (找不到) <song title>.txt
' ---------------------------------------------------------------

missingSongs = ""

For i = 0 To UBound(songNames)
    songName = Trim(songNames(i))

    If songName <> "" Then
        songFilePath = FindSongPpt(songRootPath, songName & ".pptx")

        If songFilePath <> "" Then
            errText = ""

            On Error Resume Next

            fso.CopyFile _
                songFilePath, _
                fso.BuildPath(outputFolder, "10_" & songName & ".pptx"), _
                True

            If Err.Number <> 0 Then
                errText = Err.Description
                Err.Clear
            End If

            On Error GoTo 0

            If errText <> "" Then
                CleanupFolder outputFolder
                SafeQuitPowerPoint ppt

                Fail MsgCannotCopySong() & errText
            End If

        Else
            If Not CreateMissingSongMarker( _
                outputFolder, songName, errText _
            ) Then
                CleanupFolder outputFolder
                SafeQuitPowerPoint ppt

                Fail MsgCannotCreateMarker() & errText
            End If

            If missingSongs <> "" Then
                missingSongs = missingSongs & vbCrLf
            End If

            missingSongs = missingSongs & songName
        End If
    End If
Next

RemoveTxtFilesWithMatchingPptx outputFolder
SafeQuitPowerPoint ppt
Set ppt = Nothing

' ---------------------------------------------------------------
' Refresh master after all PPTX files are in the output folder.
' ---------------------------------------------------------------

errText = ""

On Error Resume Next

shell.Run _
    "wscript.exe """ & _
    fso.BuildPath(outputFolder, "refresh-master.vbs") & _
    """", _
    0, _
    True

If Err.Number <> 0 Then
    errText = Err.Description
    Err.Clear
End If

On Error GoTo 0

If errText <> "" Then
    MsgBox MsgCreatedRefreshFailed() & vbCrLf & _
           errText & vbCrLf & vbCrLf & _
           MsgRunRefreshManually(), _
           vbExclamation, MsgAppTitle()

ElseIf missingSongs <> "" Then
    MsgBox MsgDone() & vbCrLf & vbCrLf & _
           MsgCreatedFolder() & vbCrLf & _
           outputFolder & vbCrLf & vbCrLf & _
           MsgMissingSongMarkers() & vbCrLf & _
           missingSongs, _
           vbExclamation, MsgAppTitle()

Else
    MsgBox MsgDone() & vbCrLf & vbCrLf & _
           MsgCreatedFolder() & vbCrLf & _
           outputFolder, _
           vbInformation, MsgAppTitle()
End If

' ===============================================================
' Song search
' ===============================================================

Function FindSongPpt(ByVal worshipDataRoot, ByVal requestedFileName)

    Dim songFolder
    Dim yearValue, yearFolder
    Dim foundPath

    FindSongPpt = ""

    ' First priority:
    ' 崇拜用資料\詩歌\<requested file>
    songFolder = fso.BuildPath(worshipDataRoot, CNSongs())

    If fso.FileExists(fso.BuildPath(songFolder, requestedFileName)) Then
        FindSongPpt = fso.BuildPath(songFolder, requestedFileName)
        Exit Function
    End If

    ' Second priority:
    ' Searches all ordinary subfolders within each year.
    For yearValue = MinNumber(Year(Date), MAX_SEARCH_YEAR) _
                    To MIN_SEARCH_YEAR Step -1

        yearFolder = fso.BuildPath(worshipDataRoot, CStr(yearValue))

        If fso.FolderExists(yearFolder) Then
            foundPath = FindFileRecursive(yearFolder, requestedFileName)

            If foundPath <> "" Then
                FindSongPpt = foundPath
                Exit Function
            End If
        End If
    Next
End Function

Function FindFileRecursive(ByVal folderPath, ByVal requestedFileName)

    Dim folderObject
    Dim fileObject
    Dim subFolderObject
    Dim foundPath
    Dim attributes

    FindFileRecursive = ""

    On Error Resume Next

    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    ' Check direct files in this folder first.
    For Each fileObject In folderObject.Files
        If StrComp(fileObject.Name, requestedFileName, 1) = 0 Then
            FindFileRecursive = fileObject.Path
            Exit Function
        End If
    Next

    ' Search child folders.
    ' Skip reparse points, junctions, and symbolic links.
    For Each subFolderObject In folderObject.SubFolders

        On Error Resume Next
        attributes = subFolderObject.Attributes

        If Err.Number = 0 Then
            If (attributes And FILE_ATTRIBUTE_REPARSE_POINT) = 0 Then
                On Error GoTo 0

                foundPath = FindFileRecursive( _
                    subFolderObject.Path, _
                    requestedFileName _
                )

                If foundPath <> "" Then
                    FindFileRecursive = foundPath
                    Exit Function
                End If

            Else
                Err.Clear
                On Error GoTo 0
            End If

        Else
            Err.Clear
            On Error GoTo 0
        End If
    Next
End Function

' ===============================================================
' Copy fixed files
' ===============================================================

Sub CopyDirectFiles(ByVal sourceFolder, ByVal destinationFolder, ByRef errorMessage)

    Dim fileObject

    errorMessage = ""

    On Error Resume Next

    For Each fileObject In fso.GetFolder(sourceFolder).Files
        fso.CopyFile _
            fileObject.Path, _
            fso.BuildPath(destinationFolder, fileObject.Name), _
            True

        If Err.Number <> 0 Then
            errorMessage = Err.Description
            Err.Clear
            Exit For
        End If
    Next

    On Error GoTo 0
End Sub

' ===============================================================
' Scripture PPTX generation
' ===============================================================

Function CreateScripturePpt( _
    ByVal pptApp, _
    ByVal templatePath, _
    ByVal textPath, _
    ByVal destinationPath, _
    ByRef errorMessage _
)

    Dim title
    Dim lines, lineCount
    Dim pageStart(), pageEnd(), pageCount
    Dim pres, pageIndex, slideObject
    Dim duplicateRange, pageText

    CreateScripturePpt = False
    errorMessage = ""

    If Not ReadUtf8ScriptureFile( _
        textPath, title, lines, lineCount, errorMessage _
    ) Then
        Exit Function
    End If

    If Not MakePageBreaks( _
        lines, lineCount, _
        pageStart, pageEnd, pageCount, _
        errorMessage _
    ) Then
        Exit Function
    End If

    On Error Resume Next

    fso.CopyFile templatePath, destinationPath, True

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set pres = pptApp.Presentations.Open( _
        destinationPath, _
        msoFalse, _
        msoFalse, _
        msoFalse _
    )

    If Err.Number <> 0 Or pres Is Nothing Then
        errorMessage = MsgCannotOpenTemplateCopy() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    If pres.Slides.Count <> 1 Then
        pres.Close
        errorMessage = MsgTemplateSlideCount()
        Exit Function
    End If

    For pageIndex = 2 To pageCount
        Set duplicateRange = pres.Slides(1).Duplicate
        duplicateRange(1).MoveTo pres.Slides.Count
    Next

    For pageIndex = 1 To pageCount
        Set slideObject = pres.Slides(pageIndex)

        pageText = JoinLines( _
            lines, _
            pageStart(pageIndex), _
            pageEnd(pageIndex) _
        )

        ' Replace every exact {{SCRIPTURE_TITLE}} placeholder.
        If ReplaceAllExactText( _
            slideObject, _
            "{{SCRIPTURE_TITLE}}", _
            title _
        ) = 0 Then
            pres.Close
            errorMessage = MsgTemplateMissingTitle()
            Exit Function
        End If

        ' Replace every exact {{SCRIPTURE_TEXT}} placeholder.
        If ReplaceAllExactText( _
            slideObject, _
            "{{SCRIPTURE_TEXT}}", _
            pageText _
        ) = 0 Then
            pres.Close
            errorMessage = MsgTemplateMissingText()
            Exit Function
        End If
    Next

    On Error Resume Next

    pres.Save

    If Err.Number <> 0 Then
        errorMessage = MsgCannotSavePpt() & Err.Description
        Err.Clear
        pres.Close
        On Error GoTo 0
        Exit Function
    End If

    pres.Close

    On Error GoTo 0

    CreateScripturePpt = True
End Function

Function ReplaceAllExactText( _
    ByVal slideObject, _
    ByVal placeholder, _
    ByVal replacement _
)

    Dim i, shapeObject
    Dim replacementCount
    Dim isScriptureText

    replacementCount = 0
    isScriptureText = (placeholder = "{{SCRIPTURE_TEXT}}")

    For i = 1 To slideObject.Shapes.Count
        On Error Resume Next

        Set shapeObject = slideObject.Shapes(i)

        If shapeObject.HasTextFrame = msoTrue Then
            If Trim(shapeObject.TextFrame.TextRange.Text) = placeholder Then

                If isScriptureText Then
                    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
                    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
                    shapeObject.TextFrame.VerticalAnchor = 1
                End If

                shapeObject.TextFrame.TextRange.Text = replacement

                If isScriptureText Then
                    ApplyScriptureHangingIndent shapeObject

                    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
                    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
                    shapeObject.TextFrame.VerticalAnchor = 1
                End If

                replacementCount = replacementCount + 1
            End If
        End If

        Err.Clear
        On Error GoTo 0
    Next

    ReplaceAllExactText = replacementCount
End Function

' ===============================================================
' Scripture hanging indent
' ===============================================================

Sub ApplyScriptureHangingIndent(ByVal shapeObject)

    Const HANGING_INDENT_POINTS = 55

    Dim paragraphIndex
    Dim paragraphCount
    Dim paragraphRange

    On Error Resume Next

    paragraphCount = _
        shapeObject.TextFrame.TextRange.Paragraphs.Count

    For paragraphIndex = 1 To paragraphCount

        Set paragraphRange = _
            shapeObject.TextFrame.TextRange.Paragraphs( _
                paragraphIndex, _
                1 _
            )

        ' Required: no bullets in scripture paragraphs.
        paragraphRange.ParagraphFormat.Bullet.Visible = msoFalse

        ' PowerPoint's safe indent method:
        ' positive value shifts the complete paragraph right,
        ' negative value brings only the first line back left.
        paragraphRange.ParagraphFormat.Indent _
            HANGING_INDENT_POINTS, _
            -HANGING_INDENT_POINTS
    Next

    Err.Clear
    On Error GoTo 0
End Sub

Function MakePageBreaks( _
    ByRef lines, _
    ByVal lineCount, _
    ByRef pageStart, _
    ByRef pageEnd, _
    ByRef pageCount, _
    ByRef errorMessage _
)

    Dim currentLine
    Dim pageFirst
    Dim pageChars
    Dim lineLength

    MakePageBreaks = False
    errorMessage = ""
    pageCount = 0

    ReDim pageStart(1)
    ReDim pageEnd(1)

    If lineCount = 0 Then
        pageCount = 1
        pageStart(1) = 0
        pageEnd(1) = -1

        MakePageBreaks = True
        Exit Function
    End If

    currentLine = 0

    Do While currentLine < lineCount
        pageFirst = currentLine
        pageChars = 0

        Do While currentLine < lineCount
            lineLength = Len(lines(currentLine))

            If currentLine > pageFirst Then
                lineLength = lineLength + 1
            End If

            If currentLine > pageFirst Then
                If pageChars + lineLength > MAX_CHARS_PER_SLIDE Or _
                   currentLine - pageFirst >= MAX_LINES_PER_SLIDE Then
                    Exit Do
                End If
            End If

            pageChars = pageChars + lineLength
            currentLine = currentLine + 1
        Loop

        ' Protect against a single unusually long line.
        If currentLine = pageFirst Then
            currentLine = currentLine + 1
        End If

        pageCount = pageCount + 1

        ReDim Preserve pageStart(pageCount)
        ReDim Preserve pageEnd(pageCount)

        pageStart(pageCount) = pageFirst
        pageEnd(pageCount) = currentLine - 1
    Loop

    MakePageBreaks = True
End Function

Function ReadUtf8ScriptureFile( _
    ByVal filePath, _
    ByRef title, _
    ByRef bodyLines, _
    ByRef lineCount, _
    ByRef errorMessage _
)

    Dim stream
    Dim text
    Dim rawLines
    Dim i
    Dim lineValue

    ReadUtf8ScriptureFile = False

    errorMessage = ""
    title = ""
    lineCount = 0

    ReDim bodyLines(0)

    On Error Resume Next

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        errorMessage = MsgCannotReadText() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    text = Replace(text, ChrW(&HFEFF), "")
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    rawLines = Split(text, vbLf)

    title = Trim(rawLines(0))

    If title = "" Then
        errorMessage = MsgEmptyTitle()
        Exit Function
    End If

    ReDim bodyLines(0)

    For i = 1 To UBound(rawLines)
        lineValue = rawLines(i)

        If Trim(lineValue) <> "" Then
            ReDim Preserve bodyLines(lineCount)
            bodyLines(lineCount) = lineValue
            lineCount = lineCount + 1
        End If
    Next

    ReadUtf8ScriptureFile = True
End Function

Function JoinLines(ByRef values, ByVal firstIndex, ByVal lastIndex)

    Dim i, result

    result = ""

    If lastIndex < firstIndex Then
        JoinLines = ""
        Exit Function
    End If

    For i = firstIndex To lastIndex
        If result <> "" Then
            result = result & vbCrLf
        End If

        result = result & values(i)
    Next

    JoinLines = result
End Function

' ===============================================================
' Call-and-response PPTX generation
' ===============================================================

' ===============================================================
' PART 1
' Replace these two existing functions:
'   CreateCallResponsePpt
'   SetTextInNamedShape
' ===============================================================

Function CreateCallResponsePpt( _
    ByVal pptApp, _
    ByVal templatePath, _
    ByVal textPath, _
    ByVal destinationPath, _
    ByRef errorMessage _
)

    Dim callScripture
    Dim callLines(), responseLines(), pairCount
    Dim hasCallScripture, hasCallResponse
    Dim pres, pairIndex, slideObject
    Dim duplicateRange
    Dim detailError

    CreateCallResponsePpt = False
    errorMessage = ""

    If Not ReadCallResponseFile( _
        textPath, _
        callScripture, _
        hasCallScripture, _
        callLines, _
        responseLines, _
        pairCount, _
        hasCallResponse, _
        errorMessage _
    ) Then
        Exit Function
    End If

    On Error Resume Next

    fso.CopyFile templatePath, destinationPath, True

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Set pres = pptApp.Presentations.Open( _
        destinationPath, _
        msoFalse, _
        msoFalse, _
        msoFalse _
    )

    If Err.Number <> 0 Or pres Is Nothing Then
        errorMessage = MsgCannotOpenTemplateCopy() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    If pres.Slides.Count <> 2 Then
        pres.Close
        errorMessage = MsgCallTemplateSlideCount()
        Exit Function
    End If

    ' If [啟應] exists, create one 啟應 slide per complete pair.
    ' Duplicate the original slide 2 before deleting any slide.
    If hasCallResponse Then
        For pairIndex = 2 To pairCount
            Set duplicateRange = pres.Slides(2).Duplicate
            duplicateRange(1).MoveTo pres.Slides.Count
        Next
    End If

    ' Write [宣召] text only if the section has usable content.
    ' Its empty lines are already preserved by ReadCallResponseFile.
    If hasCallScripture Then
        detailError = ""

        If Not SetTextInNamedShape( _
            pres.Slides(1), _
            "CALL_SCRIPTURE_1", _
            callScripture, _
            detailError _
        ) Then
            pres.Close
            errorMessage = MsgTemplateMissingNamedShape( _
                "CALL_SCRIPTURE_1" _
            ) & vbCrLf & detailError
            Exit Function
        End If
    End If

    ' Write each 啟 / 應 pair.
    If hasCallResponse Then
        For pairIndex = 1 To pairCount
            Set slideObject = pres.Slides(pairIndex + 1)
            detailError = ""

            If Not SetTextInNamedShape( _
                slideObject, _
                "CALL_1", _
                callLines(pairIndex), _
                detailError _
            ) Then
                pres.Close
                errorMessage = MsgTemplateMissingNamedShape( _
                    "CALL_1" _
                ) & vbCrLf & detailError
                Exit Function
            End If

            detailError = ""

            If Not SetTextInNamedShape( _
                slideObject, _
                "RESPONSE_1", _
                responseLines(pairIndex), _
                detailError _
            ) Then
                pres.Close
                errorMessage = MsgTemplateMissingNamedShape( _
                    "RESPONSE_1" _
                ) & vbCrLf & detailError
                Exit Function
            End If
        Next
    End If

    ' Remove unused template slide(s).
    '
    ' If only [啟應] exists:
    ' - template slide 1 is deleted;
    ' - original template slide 2 becomes the first output slide.
    '
    ' If only [宣召] exists:
    ' - template slide 2 is deleted.
    If Not hasCallResponse Then
        pres.Slides(2).Delete
    End If

    If Not hasCallScripture Then
        pres.Slides(1).Delete
    End If

    On Error Resume Next

    pres.Save

    If Err.Number <> 0 Then
        errorMessage = MsgCannotSavePpt() & Err.Description
        Err.Clear
        pres.Close
        On Error GoTo 0
        Exit Function
    End If

    pres.Close

    On Error GoTo 0

    CreateCallResponsePpt = True
End Function


' Gets a text box by its exact PowerPoint Selection Pane name and
' replaces all its text. It does not search for placeholder strings.
Function SetTextInNamedShape( _
    ByVal slideObject, _
    ByVal shapeName, _
    ByVal replacement, _
    ByRef errorMessage _
)

    Dim shapeObject

    SetTextInNamedShape = False
    errorMessage = ""

    On Error Resume Next

    Set shapeObject = slideObject.Shapes(shapeName)

    If Err.Number <> 0 Or shapeObject Is Nothing Then
        errorMessage = MsgNamedShapeNotFound(shapeName)
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    If shapeObject.HasTextFrame <> msoTrue Then
        errorMessage = MsgNamedShapeNoText(shapeName)
        On Error GoTo 0
        Exit Function
    End If

    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
    shapeObject.TextFrame.VerticalAnchor = 1

    shapeObject.TextFrame.TextRange.Text = replacement

    shapeObject.TextFrame.AutoSize = ppAutoSizeNone
    shapeObject.TextFrame2.AutoSize = msoAutoSizeNone
    shapeObject.TextFrame.VerticalAnchor = 1

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    SetTextInNamedShape = True
End Function

' ===============================================================
' PART 2
' Replace the existing ReadCallResponseFile function with this.
'
' Also add the four MsgCall... functions at the bottom of this part
' if they do not already exist in your VBS.
' ===============================================================

' ===============================================================
' Replace ReadCallResponseFile with this version
' ===============================================================

Function ReadCallResponseFile( _
    ByVal filePath, _
    ByRef callScripture, _
    ByRef hasCallScripture, _
    ByRef callLines, _
    ByRef responseLines, _
    ByRef pairCount, _
    ByRef hasCallResponse, _
    ByRef errorMessage _
)

    Dim stream
    Dim text
    Dim rawLines
    Dim i
    Dim rawLine
    Dim trimmedLine
    Dim section
    Dim sawCallSection, sawResponseSection
    Dim pendingCall, hasPendingCall
    Dim pendingBlankLine

    ReadCallResponseFile = False

    errorMessage = ""
    callScripture = ""
    hasCallScripture = False
    pairCount = 0
    hasCallResponse = False

    section = 0
    sawCallSection = False
    sawResponseSection = False
    pendingCall = ""
    hasPendingCall = False
    pendingBlankLine = False

    ReDim callLines(0)
    ReDim responseLines(0)

    On Error Resume Next

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    text = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        errorMessage = MsgCannotReadText() & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0

    text = Replace(text, ChrW(&HFEFF), "")
    text = Replace(text, vbTab, "")
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)

    rawLines = Split(text, vbLf)

    ' section:
    ' 0 = before / outside sections
    ' 1 = [宣召]
    ' 2 = [啟應]
    For i = 0 To UBound(rawLines)

        rawLine = rawLines(i)
        trimmedLine = Trim(rawLine)

        If trimmedLine = CNCallSection() Then

            If sawCallSection Then
                errorMessage = MsgCallSectionRepeated(CNCallSection())
                Exit Function
            End If

            section = 1
            sawCallSection = True
            pendingBlankLine = False

        ElseIf trimmedLine = CNResponseSection() Then

            If sawResponseSection Then
                errorMessage = MsgCallSectionRepeated(CNResponseSection())
                Exit Function
            End If

            section = 2
            sawResponseSection = True

        ElseIf section = 1 Then

            ' Keep empty lines in [宣召].
            ' A blank source line is remembered, then converted into
            ' TWO line breaks before the next non-empty line. This
            ' produces one visible empty paragraph in PowerPoint.
            If trimmedLine = "" Then

                If callScripture <> "" Then
                    pendingBlankLine = True
                End If

            Else

                If callScripture <> "" Then

                    If pendingBlankLine Then
                        ' One normal paragraph break + one empty paragraph.
                        callScripture = callScripture & _
                                        vbCrLf & vbCrLf
                    Else
                        callScripture = callScripture & vbCrLf
                    End If
                End If

                callScripture = callScripture & Trim(rawLine)
                pendingBlankLine = False
            End If

        ElseIf section = 2 Then

            ' Ignore blank separator lines in [啟應].
            If trimmedLine <> "" Then

                If Left(trimmedLine, 2) = CNCallPrefix() Or _
                   Left(trimmedLine, 2) = CNCallPrefixHalf() Then

                    If hasPendingCall Then
                        errorMessage = MsgCallPairMismatch()
                        Exit Function
                    End If

                    pendingCall = Trim(Mid(trimmedLine, 3))

                    If pendingCall = "" Then
                        errorMessage = MsgCallEmptyLine(CNCallPrefix())
                        Exit Function
                    End If

                    hasPendingCall = True

                ElseIf Left(trimmedLine, 2) = CNResponsePrefix() Or _
                       Left(trimmedLine, 2) = CNResponsePrefixHalf() Then

                    If Not hasPendingCall Then
                        errorMessage = MsgCallPairMismatch()
                        Exit Function
                    End If

                    If Trim(Mid(trimmedLine, 3)) = "" Then
                        errorMessage = MsgCallEmptyLine(CNResponsePrefix())
                        Exit Function
                    End If

                    pairCount = pairCount + 1

                    ReDim Preserve callLines(pairCount)
                    ReDim Preserve responseLines(pairCount)

                    callLines(pairCount) = pendingCall
                    responseLines(pairCount) = Trim(Mid(trimmedLine, 3))

                    pendingCall = ""
                    hasPendingCall = False

                Else
                    errorMessage = MsgCallBadLine()
                    Exit Function
                End If
            End If
        End If
    Next

    hasCallScripture = (Trim(callScripture) <> "")

    If hasPendingCall Then
        errorMessage = MsgCallPairMismatch()
        Exit Function
    End If

    hasCallResponse = (pairCount > 0)

    If Not hasCallScripture And Not hasCallResponse Then
        errorMessage = MsgCallNoUsableSection()
        Exit Function
    End If

    ReadCallResponseFile = True
End Function

Function MsgCallNoUsableSection()
    MsgCallNoUsableSection = CNCallResponse() & ".txt " & _
                             ChrW(&H5167) & ChrW(&H6C92) & ChrW(&H6709) & _
                             ChrW(&H53EF) & ChrW(&H7528) & ChrW(&H5167) & _
                             ChrW(&H5BB9) & ChrW(&H3002) & _
                             ChrW(&H8ACB) & ChrW(&H586B) & ChrW(&H5BEB) & _
                             " " & CNCallSection() & " " & _
                             ChrW(&H6216) & " " & CNResponseSection() & _
                             ChrW(&H3002)
End Function


Function MsgCallSectionRepeated(ByVal sectionName)
    MsgCallSectionRepeated = sectionName & " " & _
                             ChrW(&H6BB5) & ChrW(&H843D) & _
                             ChrW(&H91CD) & ChrW(&H8907) & ChrW(&H51FA) & _
                             ChrW(&H73FE) & ChrW(&H3002)
End Function


Function MsgCallEmptyLine(ByVal prefix)
    MsgCallEmptyLine = prefix & " " & _
                       ChrW(&H5F8C) & ChrW(&H5FC5) & ChrW(&H9808) & _
                       ChrW(&H586B) & ChrW(&H5BEB) & ChrW(&H6587) & _
                       ChrW(&H5B57) & ChrW(&H3002)
End Function


Function MsgCallBadLine()
    MsgCallBadLine = CNResponseSection() & " " & _
                     ChrW(&H6BB5) & ChrW(&H843D) & _
                     ChrW(&H5167) & ChrW(&HFF0C) & _
                     ChrW(&H6BCF) & ChrW(&H884C) & _
                     ChrW(&H5FC5) & ChrW(&H9808) & _
                     ChrW(&H4EE5) & " " & _
                     CNCallPrefix() & " " & _
                     ChrW(&H6216) & " " & _
                     CNResponsePrefix() & " " & _
                     ChrW(&H958B) & ChrW(&H982D) & ChrW(&H3002)
End Function

' ===============================================================
' Missing-song marker
' ===============================================================

Function CreateMissingSongMarker( _
    ByVal destinationFolder, _
    ByVal songTitle, _
    ByRef errorMessage _
)

    Dim markerFile
    Dim textStream

    CreateMissingSongMarker = False
    errorMessage = ""

    markerFile = fso.BuildPath( _
        destinationFolder, _
        MsgMissingPrefix() & songTitle & ".txt" _
    )

    On Error Resume Next

    Set textStream = fso.CreateTextFile(markerFile, True, True)

    If Err.Number <> 0 Then
        errorMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    textStream.WriteLine songTitle
    textStream.Close

    On Error GoTo 0

    CreateMissingSongMarker = True
End Function

' ===============================================================
' Path and validation helpers
' ===============================================================

Function GetShortcutTarget(ByVal shortcutPath)

    Dim shortcutObject

    On Error Resume Next

    Set shortcutObject = shell.CreateShortcut(shortcutPath)
    GetShortcutTarget = shortcutObject.TargetPath

    If Err.Number <> 0 Then
        Err.Clear
        GetShortcutTarget = ""
    End If

    On Error GoTo 0
End Function

Function IsValidFolderName(ByVal value)

    Dim invalidCharacters, i

    invalidCharacters = "\/:*?""<>|"

    IsValidFolderName = True

    If value = "" Or value = "." Or value = ".." Then
        IsValidFolderName = False
        Exit Function
    End If

    For i = 1 To Len(invalidCharacters)
        If InStr(value, Mid(invalidCharacters, i, 1)) > 0 Then
            IsValidFolderName = False
            Exit Function
        End If
    Next
End Function

Function MinNumber(ByVal firstValue, ByVal secondValue)

    If firstValue < secondValue Then
        MinNumber = firstValue
    Else
        MinNumber = secondValue
    End If
End Function

Sub SafeQuitPowerPoint(ByVal pptApp)

    On Error Resume Next

    If Not pptApp Is Nothing Then
        pptApp.Quit
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Sub CleanupFolder(ByVal path)

    On Error Resume Next

    If fso.FolderExists(path) Then
        fso.DeleteFolder path, True
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Sub Fail(ByVal text)

    MsgBox text, vbCritical, MsgAppTitle()
    WScript.Quit 1
End Sub

' ===============================================================
' Remove copied TXT files when a same-name PPTX exists
' ===============================================================

Sub RemoveTxtFilesWithMatchingPptx(ByVal folderPath)

    Dim folderObject
    Dim fileObject
    Dim pptxPath

    On Error Resume Next

    Set folderObject = fso.GetFolder(folderPath)

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    On Error GoTo 0

    ' The folder is newly created, so deletion is safe:
    ' only files such as 02_宣召及啟應.txt are removed when
    ' 02_宣召及啟應.pptx is present in that same folder.
    For Each fileObject In folderObject.Files

        If LCase(fso.GetExtensionName(fileObject.Name)) = "txt" Then

            pptxPath = fso.BuildPath( _
                folderPath, _
                fso.GetBaseName(fileObject.Name) & ".pptx" _
            )

            On Error Resume Next

            If fso.FileExists(pptxPath) Then
                fso.DeleteFile fileObject.Path, True
            End If

            Err.Clear
            On Error GoTo 0
        End If
    Next
End Sub

' ===============================================================
' Chinese filenames and Chinese messages
' ===============================================================

Function CNRead()
    CNRead = ChrW(&H8B80) & ChrW(&H7D93)
End Function

Function CNScriptures()
    CNScriptures = ChrW(&H7D93) & ChrW(&H8A13)
End Function

Function CNSongs()
    CNSongs = ChrW(&H8A69) & ChrW(&H6B4C)
End Function

Function CNCallResponse()
    CNCallResponse = ChrW(&H5BA3) & ChrW(&H53EC) & _
                     ChrW(&H53CA) & ChrW(&H555F) & _
                     ChrW(&H61C9)
End Function

Function CNCallSection()
    CNCallSection = "[" & ChrW(&H5BA3) & ChrW(&H53EC) & "]"
End Function

Function CNResponseSection()
    CNResponseSection = "[" & ChrW(&H555F) & ChrW(&H61C9) & "]"
End Function

Function CNCallPrefix()
    CNCallPrefix = ChrW(&H555F) & ChrW(&HFF1A)
End Function

Function CNResponsePrefix()
    CNResponsePrefix = ChrW(&H61C9) & ChrW(&HFF1A)
End Function

Function CNCallPrefixHalf()
    CNCallPrefixHalf = ChrW(&H555F) & ":"
End Function

Function CNResponsePrefixHalf()
    CNResponsePrefixHalf = ChrW(&H61C9) & ":"
End Function

Function MsgAppTitle()
    MsgAppTitle = ChrW(&H5EFA) & ChrW(&H7ACB) & _
                  ChrW(&H5D07) & ChrW(&H62DC) & _
                  ChrW(&H8CC7) & ChrW(&H6599) & _
                  ChrW(&H593E)
End Function

Function MsgAutoRunMissing()
    MsgAutoRunMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                        " autorun " & _
                        ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                        ChrW(&H3002)
End Function

Function MsgCopyFolderMissing()
    MsgCopyFolderMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                           " autorun\worship-files " & _
                           ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                           ChrW(&H3002)
End Function

Function MsgFileMissing(ByVal f)
    MsgFileMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                     ChrW(&H6A94) & ChrW(&H6848) & ChrW(&HFF1A) & f
End Function

Function MsgEnterFolderName()
    MsgEnterFolderName = ChrW(&H8ACB) & ChrW(&H8F38) & ChrW(&H5165) & _
                         ChrW(&H65B0) & ChrW(&H7684) & _
                         ChrW(&H5D07) & ChrW(&H62DC) & _
                         ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                         ChrW(&H540D) & ChrW(&H7A31) & ChrW(&HFF1A)
End Function

Function MsgInvalidFolderName()
    MsgInvalidFolderName = ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                           ChrW(&H540D) & ChrW(&H7A31) & _
                           ChrW(&H7121) & ChrW(&H6548) & ChrW(&H3002) & _
                           ChrW(&H4E0D) & ChrW(&H53EF) & ChrW(&H4F7F) & _
                           ChrW(&H7528) & ChrW(&HFF1A) & _
                           "\ / : * ? "" < > |"
End Function

Function MsgFolderExists()
    MsgFolderExists = ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                      ChrW(&H5DF2) & ChrW(&H7D93) & ChrW(&H5B58) & _
                      ChrW(&H5728) & ChrW(&HFF1A)
End Function

Function MsgNoChanges()
    MsgNoChanges = ChrW(&H6C92) & ChrW(&H6709) & _
                   ChrW(&H4F5C) & ChrW(&H51FA) & _
                   ChrW(&H4EFB) & ChrW(&H4F55) & _
                   ChrW(&H8B8A) & ChrW(&H66F4) & _
                   ChrW(&H3002)
End Function

Function MsgEnterSongs()
    MsgEnterSongs = ChrW(&H8ACB) & ChrW(&H8F38) & ChrW(&H5165) & _
                    ChrW(&H8A69) & ChrW(&H6B4C) & _
                    ChrW(&H540D) & ChrW(&H7A31) & _
                    ChrW(&HFF0C) & ChrW(&H4EE5) & _
                    ChrW(&H9017) & ChrW(&H865F) & _
                    ChrW(&H5206) & ChrW(&H9694) & _
                    ChrW(&HFF1A)
End Function

Function MsgSongExample()
    MsgSongExample = ChrW(&H4F8B) & ChrW(&H5982) & ChrW(&HFF1A) & _
                     ChrW(&H8A69) & ChrW(&H6B4C) & "A," & _
                     ChrW(&H8A69) & ChrW(&H6B4C) & "B," & _
                     ChrW(&H8A69) & ChrW(&H6B4C) & "C"
End Function

Function MsgSongRootMissing()
    MsgSongRootMissing = ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                         " songs-folder.lnk " & _
                         ChrW(&H6240) & ChrW(&H6307) & ChrW(&H5B9A) & _
                         ChrW(&H7684) & ChrW(&H5D07) & ChrW(&H62DC) & _
                         ChrW(&H7528) & ChrW(&H8CC7) & ChrW(&H6599) & _
                         ChrW(&H6839) & ChrW(&H76EE) & ChrW(&H9304) & _
                         ChrW(&H3002)
End Function

Function MsgCannotStartPowerPoint()
    MsgCannotStartPowerPoint = ChrW(&H7121) & ChrW(&H6CD5) & _
                               ChrW(&H555F) & ChrW(&H52D5) & _
                               " Microsoft PowerPoint" & ChrW(&H3002)
End Function

Function MsgCannotCopyFixedFiles()
    MsgCannotCopyFixedFiles = ChrW(&H7121) & ChrW(&H6CD5) & _
                              ChrW(&H8907) & ChrW(&H88FD) & _
                              " worship-files " & _
                              ChrW(&H5167) & ChrW(&H7684) & _
                              ChrW(&H6A94) & ChrW(&H6848) & _
                              ChrW(&HFF1A)
End Function

Function MsgCannotCreateFile(ByVal f)
    MsgCannotCreateFile = ChrW(&H7121) & ChrW(&H6CD5) & _
                          ChrW(&H5EFA) & ChrW(&H7ACB) & _
                          " " & f & ChrW(&HFF1A)
End Function

Function MsgCannotCopySong()
    MsgCannotCopySong = ChrW(&H7121) & ChrW(&H6CD5) & _
                        ChrW(&H8907) & ChrW(&H88FD) & _
                        ChrW(&H8A69) & ChrW(&H6B4C) & _
                        " PPTX" & ChrW(&HFF1A)
End Function

Function MsgCannotCreateMarker()
    MsgCannotCreateMarker = ChrW(&H7121) & ChrW(&H6CD5) & _
                            ChrW(&H5EFA) & ChrW(&H7ACB) & _
                            ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                            ChrW(&H8A69) & ChrW(&H6B4C) & _
                            ChrW(&H7684) & ChrW(&H63D0) & ChrW(&H793A) & _
                            ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                            ChrW(&HFF1A)
End Function

Function MsgCreatedRefreshFailed()
    MsgCreatedRefreshFailed = ChrW(&H5D07) & ChrW(&H62DC) & _
                              ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                              ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & _
                              ChrW(&HFF0C) & ChrW(&H4F46) & _
                              ChrW(&H7121) & ChrW(&H6CD5) & _
                              ChrW(&H81EA) & ChrW(&H52D5) & _
                              ChrW(&H66F4) & ChrW(&H65B0) & _
                              " master" & ChrW(&HFF1A)
End Function

Function MsgRunRefreshManually()
    MsgRunRefreshManually = ChrW(&H8ACB) & ChrW(&H5728) & _
                            ChrW(&H65B0) & ChrW(&H8CC7) & _
                            ChrW(&H6599) & ChrW(&H593E) & _
                            ChrW(&H5167) & ChrW(&H624B) & _
                            ChrW(&H52D5) & ChrW(&H57F7) & _
                            ChrW(&H884C) & _
                            " refresh-master.vbs" & ChrW(&H3002)
End Function

Function MsgDone()
    MsgDone = ChrW(&H5B8C) & ChrW(&H6210) & ChrW(&H3002)
End Function

Function MsgCreatedFolder()
    MsgCreatedFolder = ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & _
                       ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & _
                       ChrW(&HFF1A)
End Function

Function MsgMissingSongMarkers()
    MsgMissingSongMarkers = ChrW(&H4EE5) & ChrW(&H4E0B) & _
                            ChrW(&H8A69) & ChrW(&H6B4C) & _
                            ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                            ChrW(&HFF0C) & ChrW(&H5DF2) & _
                            ChrW(&H5EFA) & ChrW(&H7ACB) & _
                            ChrW(&H63D0) & ChrW(&H793A) & _
                            ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                            ChrW(&HFF1A)
End Function

Function MsgMissingPrefix()
    MsgMissingPrefix = "(" & _
                       ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                       ") "
End Function

Function MsgCannotOpenTemplateCopy()
    MsgCannotOpenTemplateCopy = ChrW(&H7121) & ChrW(&H6CD5) & _
                                ChrW(&H958B) & ChrW(&H555F) & _
                                ChrW(&H8907) & ChrW(&H88FD) & _
                                ChrW(&H5F8C) & ChrW(&H7684) & _
                                " template" & ChrW(&HFF1A)
End Function

Function MsgTemplateSlideCount()
    MsgTemplateSlideCount = "template " & _
                            ChrW(&H5FC5) & ChrW(&H9808) & _
                            ChrW(&H53EA) & ChrW(&H5305) & ChrW(&H542B) & _
                            " 1 " & ChrW(&H5F35) & _
                            ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & _
                            ChrW(&H3002)
End Function

Function MsgTemplateMissingTitle()
    MsgTemplateMissingTitle = "template " & _
                              ChrW(&H7F3A) & ChrW(&H5C11) & _
                              " {{SCRIPTURE_TITLE}}" & ChrW(&H3002)
End Function

Function MsgTemplateMissingText()
    MsgTemplateMissingText = "template " & _
                             ChrW(&H7F3A) & ChrW(&H5C11) & _
                             " {{SCRIPTURE_TEXT}}" & ChrW(&H3002)
End Function

Function MsgCannotSavePpt()
    MsgCannotSavePpt = ChrW(&H7121) & ChrW(&H6CD5) & _
                       ChrW(&H5132) & ChrW(&H5B58) & _
                       ChrW(&H8F38) & ChrW(&H51FA) & ChrW(&H7684) & _
                       " PPTX" & ChrW(&HFF1A)
End Function

Function MsgCannotReadText()
    MsgCannotReadText = ChrW(&H7121) & ChrW(&H6CD5) & _
                        ChrW(&H8B80) & ChrW(&H53D6) & _
                        " UTF-8 " & _
                        ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                        ChrW(&HFF1A)
End Function

Function MsgEmptyTitle()
    MsgEmptyTitle = ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H6A94) & _
                    ChrW(&H7B2C) & ChrW(&H4E00) & ChrW(&H884C) & _
                    ChrW(&H5FC5) & ChrW(&H9808) & ChrW(&H662F) & _
                    ChrW(&H7D93) & ChrW(&H6587) & _
                    ChrW(&H6A19) & ChrW(&H984C) & ChrW(&H3002)
End Function

Function MsgCallTemplateSlideCount()
    MsgCallTemplateSlideCount = CNCallResponse() & " template " & _
                                ChrW(&H5FC5) & ChrW(&H9808) & _
                                ChrW(&H53EA) & ChrW(&H5305) & ChrW(&H542B) & _
                                " 2 " & ChrW(&H5F35) & _
                                ChrW(&H6295) & ChrW(&H5F71) & ChrW(&H7247) & _
                                ChrW(&H3002)
End Function

Function MsgTemplateMissingNamedShape(ByVal shapeName)
    MsgTemplateMissingNamedShape = "template " & _
                                   ChrW(&H7F3A) & ChrW(&H5C11) & _
                                   " Selection Pane " & _
                                   ChrW(&H540D) & ChrW(&H7A31) & " " & _
                                   shapeName & ChrW(&H3002)
End Function

Function MsgNamedShapeNotFound(ByVal shapeName)
    MsgNamedShapeNotFound = "Selection Pane " & _
                            ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & _
                            " " & shapeName & ChrW(&H3002)
End Function

Function MsgNamedShapeNoText(ByVal shapeName)
    MsgNamedShapeNoText = "Selection Pane " & _
                          ChrW(&H540D) & ChrW(&H7A31) & " " & _
                          shapeName & " " & _
                          ChrW(&H4E0D) & ChrW(&H662F) & _
                          ChrW(&H6587) & ChrW(&H5B57) & ChrW(&H65B9) & _
                          ChrW(&H584A) & ChrW(&H3002)
End Function

Function MsgCallSectionMissing()
    MsgCallSectionMissing = CNCallResponse() & ".txt " & _
                            ChrW(&H7F3A) & ChrW(&H5C11) & " " & _
                            CNCallSection() & " " & _
                            ChrW(&H6216) & " " & _
                            CNResponseSection() & " " & _
                            ChrW(&H6BB5) & ChrW(&H843D) & ChrW(&H3002)
End Function

Function MsgCallEmptyScripture()
    MsgCallEmptyScripture = CNCallSection() & " " & _
                            ChrW(&H6BB5) & ChrW(&H843D) & _
                            ChrW(&H6C92) & ChrW(&H6709) & _
                            ChrW(&H5167) & ChrW(&H5BB9) & ChrW(&H3002)
End Function

Function MsgCallNoPairs()
    MsgCallNoPairs = CNResponseSection() & " " & _
                     ChrW(&H6BB5) & ChrW(&H843D) & _
                     ChrW(&H6C92) & ChrW(&H6709) & _
                     ChrW(&H4EFB) & ChrW(&H4F55) & " " & _
                     CNCallPrefix() & "/" & CNResponsePrefix() & " " & _
                     ChrW(&H914D) & ChrW(&H5C0D) & ChrW(&H3002)
End Function

Function MsgCallPairMismatch()
    MsgCallPairMismatch = ChrW(&H6BCF) & ChrW(&H53E5) & " " & _
                          CNCallPrefix() & " " & _
                          ChrW(&H4E4B) & ChrW(&H5F8C) & _
                          ChrW(&H5FC5) & ChrW(&H9808) & _
                          ChrW(&H8DDF) & ChrW(&H96A8) & _
                          ChrW(&H4E00) & ChrW(&H53E5) & " " & _
                          CNResponsePrefix() & ChrW(&H3002)
End Function
