Attribute VB_Name = "CatalogMacros"
Option Explicit

' Strider Media Catalog — Photos + Videos
' Entry / VideoEntry: A = labels, B = values, C+ = click-to-pick options
'
' Shortcuts (Control+Shift — not ⌘; avoid Excel built-ins like L/A/M/F/Y):
'   Control+Shift+Q  LoadFromFinderSelection
'   Control+Shift+I  LoadFromPreviewSelection   (front Preview document)
'   Control+Shift+W  AppendPhotoRow
'   Control+Shift+E  LoadVideoFromFinderSelection
'   Control+Shift+U  LoadVideoFromPreviewSelection
'   Control+Shift+R  AppendVideoRow
'   Control+Shift+O  OpenSelectedInPreview
'   Control+Shift+J  RevealSelectedInFinder
'   Control+Shift+D  TrashFrontPreviewDocument  (confirm, then Trash)
'   Control+Shift+K  AddCatalogField            (new click-pick tag)
'   Control+Shift+S  AppendPhotoSequence        (all frames of Entry SequenceID)
' Re-import this module, then run AssignMacroShortcuts (or reopen the workbook).
' Re-import alone does not rebind OnKey while the book stays open — Auto_Open only runs on open.
' If click-to-pick stops working, run FixEnableEvents, then re-paste Entry/VideoEntry SelectionChange.

Private Const ENTRY_SHEET As String = "Entry"
Private Const VIDEO_ENTRY_SHEET As String = "VideoEntry"
Private Const PHOTOS_SHEET As String = "Photos"
Private Const VIDEOS_SHEET As String = "Videos"
Private Const LISTS_SHEET As String = "Lists"
Private Const CAPTIONS_SHEET As String = "Captions"
Private Const HEADER_ROW As Long = 1
' Captions sheet: row1 title, row2 prompt, row3 hint, row4 headers, row5+ data
Private Const CAPTIONS_HEADER_ROW As Long = 4
Private Const LISTS_HEADER_ROW As Long = 2
Private Const ENTRY_FIRST_FIELD_ROW As Long = 9
Private Const ENTRY_OPTION_FIRST_COL As Long = 3
' Default caption provider: google = Gemini Flash (cheap + good enough). Use "anthropic" for Claude.
Private Const DEFAULT_CAPTION_PROVIDER As String = "google"
' Guard so caption-edit dialog does not re-enter on SelectionChange.
Private gCaptionEditorOpen As Boolean

Public Sub Auto_Open()
    Application.EnableEvents = True
    ' AssignMacroShortcuts also re-applies UserInterfaceOnly sheet protection
    ' (that flag does not persist across reopen; without it macros get 1004).
    AssignMacroShortcuts
End Sub

Public Sub Auto_Close()
    ClearMacroShortcuts
End Sub

' Run if click-to-pick / sheet events seem dead.
Public Sub FixEnableEvents()
    Application.EnableEvents = True
End Sub

Public Sub AssignMacroShortcuts()
    On Error Resume Next
    ' Clear previous / conflicting bindings
    Application.OnKey "^+l"
    Application.OnKey "^+a"
    Application.OnKey "^+m"
    Application.OnKey "^+p"
    Application.OnKey "^+f"
    Application.OnKey "^+y"
    Application.OnKey "^+q"
    Application.OnKey "^+i"
    Application.OnKey "^+w"
    Application.OnKey "^+e"
    Application.OnKey "^+u"
    Application.OnKey "^+r"
    Application.OnKey "^+o"
    Application.OnKey "^+j"
    Application.OnKey "^+d"
    Application.OnKey "^+y"
    Application.OnKey "^+k"
    Application.OnKey "^+s"

    Application.OnKey "^+q", "LoadFromFinderSelection"
    Application.OnKey "^+i", "LoadFromPreviewSelection"
    Application.OnKey "^+w", "AppendPhotoRow"
    Application.OnKey "^+e", "LoadVideoFromFinderSelection"
    Application.OnKey "^+u", "LoadVideoFromPreviewSelection"
    Application.OnKey "^+r", "AppendVideoRow"
    Application.OnKey "^+o", "OpenSelectedInPreview"
    Application.OnKey "^+j", "RevealSelectedInFinder"
    Application.OnKey "^+d", "TrashFrontPreviewDocument"
    Application.OnKey "^+y", "CaptionSelectedPhotos"
    Application.OnKey "^+k", "AddCatalogField"
    Application.OnKey "^+s", "AppendPhotoSequence"
    On Error GoTo 0
    EnsureRawAndPixelFields
    ReapplyUiOnlyProtection
End Sub

' One-shot / bootstrap: HaveRaw? (after Path), VertPix/HorPix (after FileSize) on Entry + Photos.
Public Sub EnsureRawAndPixelFields()
    Dim yn As Variant
    Dim none As Variant
    Dim structureOn As Boolean

    yn = Array("Y", "N")
    none = Array()

    On Error Resume Next
    structureOn = ThisWorkbook.ProtectStructure
    UnprotectWorkbookStructure
    UnprotectSheet Worksheets(ENTRY_SHEET)
    UnprotectSheet Worksheets(PHOTOS_SHEET)
    UnprotectSheet Worksheets(LISTS_SHEET)
    On Error GoTo 0

    On Error Resume Next
    If ListsHeaderCol("HaveRaw?") = 0 Then InsertListsColumn "HaveRaw?", yn
    InsertEntryFieldRow ENTRY_SHEET, "HaveRaw?", yn, "Path"
    InsertEntryFieldRow ENTRY_SHEET, "VertPix", none, "FileSize"
    InsertEntryFieldRow ENTRY_SHEET, "HorPix", none, "VertPix"
    InsertCatalogColumn PHOTOS_SHEET, "HaveRaw?", "Path"
    InsertCatalogColumn PHOTOS_SHEET, "VertPix", "FileSize"
    InsertCatalogColumn PHOTOS_SHEET, "HorPix", "VertPix"
    InsertHaveRawMetadataPreview ENTRY_SHEET
    InsertMetadataPreviewRow ENTRY_SHEET, "VertPix"
    InsertMetadataPreviewRow ENTRY_SHEET, "HorPix"
    On Error GoTo 0

    If structureOn Then ProtectWorkbookStructure
End Sub

Private Sub InsertHaveRawMetadataPreview(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim fieldRow As Long
    Dim valueCol As Long
    Dim label As String
    Dim r As Long
    Dim sectionStart As Long
    Dim insertAt As Long
    Dim col As String

    Set ws = Worksheets(sheetName)
    If Not FindFormField(sheetName, "HaveRaw?", fieldRow, valueCol) Then Exit Sub
    label = "Keyword rawYes/rawNo:"
    For r = 1 To 200
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), label, vbTextCompare) = 0 Then Exit Sub
    Next r

    sectionStart = 0
    For r = 1 To 200
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Metadata written to file on Append", vbTextCompare) = 0 Then
            sectionStart = r
            Exit For
        End If
    Next r
    If sectionStart = 0 Then Exit Sub

    insertAt = 0
    For r = sectionStart + 1 To sectionStart + 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Location → Locale", vbTextCompare) = 0 Or _
           Left$(Trim$(CStr(ws.Cells(r, 1).Value)), 12) = "Description " Then
            insertAt = r
            Exit For
        End If
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit For
    Next r
    If insertAt = 0 Then insertAt = sectionStart + 1

    ws.Rows(insertAt).Insert Shift:=xlDown
    ws.Cells(insertAt, 1).Value = label
    ws.Cells(insertAt, 1).Font.Bold = True
    col = ColLetter(valueCol)
    ws.Cells(insertAt, 2).Formula = _
        "=IF($" & col & "$" & fieldRow & "="""",""(none)"",IF(OR(UPPER($" & col & "$" & fieldRow & ")=""Y"",UPPER($" & col & "$" & fieldRow & ")=""YES""),""rawYes"",""rawNo""))"
End Sub

Public Sub ClearMacroShortcuts()
    On Error Resume Next
    Application.OnKey "^+l"
    Application.OnKey "^+a"
    Application.OnKey "^+m"
    Application.OnKey "^+p"
    Application.OnKey "^+f"
    Application.OnKey "^+y"
    Application.OnKey "^+q"
    Application.OnKey "^+i"
    Application.OnKey "^+w"
    Application.OnKey "^+e"
    Application.OnKey "^+u"
    Application.OnKey "^+r"
    Application.OnKey "^+o"
    Application.OnKey "^+j"
    Application.OnKey "^+d"
    Application.OnKey "^+k"
    Application.OnKey "^+s"
    On Error GoTo 0
End Sub

' Add a click-pick tag: prompts for name + options; updates Lists, Entry, VideoEntry, Photos, Videos.
' Existing catalog rows are left intact (new column empty).
Public Sub AddCatalogField()
    Dim fieldName As String
    Dim optionsRaw As String
    Dim afterRaw As String
    Dim afterName As String
    Dim opts() As String
    Dim parts() As String
    Dim i As Long
    Dim nOpts As Long
    Dim piece As String

    fieldName = Trim$(CStr(Application.InputBox( _
        Prompt:="New tag field name (e.g. Substrate):", _
        Title:="Add Catalog Field", Type:=2)))
    If fieldName = "False" Or Len(fieldName) = 0 Then Exit Sub

    If FormRowFor(ENTRY_SHEET, fieldName) > 0 Then
        MsgBox "Entry already has a field named """ & fieldName & """.", vbExclamation
        Exit Sub
    End If
    If HeaderCol(Worksheets(PHOTOS_SHEET), fieldName) > 0 Then
        MsgBox "Photos already has a column named """ & fieldName & """.", vbExclamation
        Exit Sub
    End If
    If ListsHeaderCol(fieldName) > 0 Then
        MsgBox "Lists already has a column named """ & fieldName & """.", vbExclamation
        Exit Sub
    End If

    optionsRaw = Trim$(CStr(Application.InputBox( _
        Prompt:="Click-pick options, separated by commas" & vbCrLf & _
                "(e.g. mud, sand, rock, vegetation):", _
        Title:="Add Catalog Field", Type:=2)))
    If optionsRaw = "False" Or Len(optionsRaw) = 0 Then Exit Sub

    parts = Split(optionsRaw, ",")
    nOpts = 0
    ReDim opts(0 To 0)
    For i = LBound(parts) To UBound(parts)
        piece = Trim$(parts(i))
        If Len(piece) > 0 Then
            ReDim Preserve opts(0 To nOpts)
            opts(nOpts) = piece
            nOpts = nOpts + 1
        End If
    Next i
    If nOpts = 0 Then
        MsgBox "Enter at least one option.", vbExclamation
        Exit Sub
    End If

    afterRaw = Trim$(CStr(Application.InputBox( _
        Prompt:="Insert after which existing field?" & vbCrLf & _
                "(Leave blank to insert above notes)", _
        Title:="Add Catalog Field", Default:="", Type:=2)))
    If afterRaw = "False" Then Exit Sub
    afterName = Trim$(afterRaw)

    If Len(afterName) > 0 Then
        If FormRowFor(ENTRY_SHEET, afterName) = 0 Then
            MsgBox "Entry has no field named """ & afterName & """.", vbExclamation
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    On Error GoTo failSoft

    InsertListsColumn fieldName, opts
    InsertEntryFieldRow ENTRY_SHEET, fieldName, opts, afterName
    InsertMetadataPreviewRow ENTRY_SHEET, fieldName
    On Error Resume Next
    InsertEntryFieldRow VIDEO_ENTRY_SHEET, fieldName, opts, afterName
    InsertMetadataPreviewRow VIDEO_ENTRY_SHEET, fieldName
    On Error GoTo failSoft
    InsertCatalogColumn PHOTOS_SHEET, fieldName, afterName
    On Error Resume Next
    InsertCatalogColumn VIDEOS_SHEET, fieldName, afterName
    On Error GoTo failSoft

    Application.ScreenUpdating = True
    MsgBox "Added tag """ & fieldName & """ to Lists, Entry, VideoEntry, Photos, Videos," & vbCrLf & _
           "and the ""Metadata written to file on Append"" preview." & vbCrLf & vbCrLf & _
           "Save the workbook. Existing rows keep a blank value for the new column.", vbInformation
    Exit Sub

failSoft:
    Application.ScreenUpdating = True
    MsgBox "Could not finish adding the field: " & Err.Description, vbExclamation
End Sub

Private Function ListsHeaderCol(ByVal headerName As String) As Long
    Dim ws As Worksheet
    Dim c As Long
    On Error Resume Next
    Set ws = Worksheets(LISTS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        ListsHeaderCol = 0
        Exit Function
    End If
    For c = 1 To 64
        If StrComp(Trim$(CStr(ws.Cells(LISTS_HEADER_ROW, c).Value)), headerName, vbTextCompare) = 0 Then
            ListsHeaderCol = c
            Exit Function
        End If
        If Len(Trim$(CStr(ws.Cells(LISTS_HEADER_ROW, c).Value))) = 0 And c > 1 Then Exit For
    Next c
    ListsHeaderCol = 0
End Function

Private Sub InsertListsColumn(ByVal fieldName As String, opts As Variant)
    Dim ws As Worksheet
    Dim col As Long
    Dim c As Long
    Dim i As Long

    Set ws = Worksheets(LISTS_SHEET)
    col = 1
    For c = 1 To 64
        If Len(Trim$(CStr(ws.Cells(LISTS_HEADER_ROW, c).Value))) > 0 Then
            col = c + 1
        End If
    Next c
    ws.Cells(LISTS_HEADER_ROW, col).Value = fieldName
    ws.Cells(LISTS_HEADER_ROW, col).Font.Bold = True
    For i = LBound(opts) To UBound(opts)
        ws.Cells(LISTS_HEADER_ROW + 1 + (i - LBound(opts)), col).Value = opts(i)
    Next i
    On Error Resume Next
    ws.Range(ws.Cells(LISTS_HEADER_ROW, col), ws.Cells(200, col)).Locked = False
    On Error GoTo 0
End Sub

Private Sub InsertEntryFieldRow(ByVal sheetName As String, ByVal fieldName As String, _
                                opts As Variant, ByVal afterName As String)
    Dim ws As Worksheet
    Dim insertAt As Long
    Dim notesRow As Long
    Dim i As Long

    Set ws = Worksheets(sheetName)
    If FormRowFor(sheetName, fieldName) > 0 Then Exit Sub

    If Len(afterName) > 0 Then
        insertAt = FormRowFor(sheetName, afterName) + 1
        If insertAt <= 1 Then Exit Sub
    Else
        notesRow = FormRowFor(sheetName, "notes")
        If notesRow = 0 Then notesRow = FormRowFor(sheetName, "Note")
        If notesRow > 0 Then
            insertAt = notesRow
        Else
            insertAt = ENTRY_FIRST_FIELD_ROW
        End If
    End If

    ws.Rows(insertAt).Insert Shift:=xlDown
    ws.Cells(insertAt, 1).Value = fieldName
    ws.Cells(insertAt, 1).Font.Bold = True
    If IsArray(opts) Then
        On Error Resume Next
        For i = LBound(opts) To UBound(opts)
            If Err.Number <> 0 Then Exit For
            ws.Cells(insertAt, ENTRY_OPTION_FIRST_COL + (i - LBound(opts))).Value = opts(i)
        Next i
        On Error GoTo 0
    End If
    ' Sheet protection: inserted rows often inherit Locked=True — keep value/options editable.
    UnlockEntryTagRow ws, insertAt
End Sub

' Under sheet protection, keep tag value (B) and options (C+) user-editable.
Private Sub UnlockEntryTagRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim lastOptCol As Long
    On Error Resume Next
    ws.Cells(rowNum, 2).Locked = False
    ' C through CZ — room to type additional options later
    lastOptCol = 104
    ws.Range(ws.Cells(rowNum, ENTRY_OPTION_FIRST_COL), ws.Cells(rowNum, lastOptCol)).Locked = False
    On Error GoTo 0
End Sub

' Add Keyword preview under "Metadata written to file on Append".
Private Sub InsertMetadataPreviewRow(ByVal sheetName As String, ByVal fieldName As String)
    Dim ws As Worksheet
    Dim fieldRow As Long
    Dim keyName As String
    Dim label As String
    Dim r As Long
    Dim sectionStart As Long
    Dim insertAt As Long

    Set ws = Worksheets(sheetName)
    fieldRow = FormRowFor(sheetName, fieldName)
    If fieldRow = 0 Then Exit Sub

    keyName = KeywordKeyForLabel(fieldName)
    label = "Keyword " & keyName & ":"

    ' Already present?
    For r = 1 To 200
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), label, vbTextCompare) = 0 Then Exit Sub
    Next r

    sectionStart = 0
    For r = 1 To 200
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Metadata written to file on Append", vbTextCompare) = 0 Then
            sectionStart = r
            Exit For
        End If
    Next r
    If sectionStart = 0 Then Exit Sub

    ' Insert before Location / Description, else after last Keyword row in the section.
    insertAt = 0
    For r = sectionStart + 1 To sectionStart + 40
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Location → Locale", vbTextCompare) = 0 Or _
           StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Description → notes", vbTextCompare) = 0 Or _
           StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Description → Note", vbTextCompare) = 0 Or _
           StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Description → NOTE:", vbTextCompare) = 0 Or _
           StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), "Description → CAPTION:", vbTextCompare) = 0 Then
            insertAt = r
            Exit For
        End If
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then Exit For
    Next r
    If insertAt = 0 Then
        insertAt = sectionStart + 1
        For r = sectionStart + 1 To sectionStart + 40
            If Left$(Trim$(CStr(ws.Cells(r, 1).Value)), 8) = "Keyword " Then
                insertAt = r + 1
            ElseIf Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then
                Exit For
            End If
        Next r
    End If

    ws.Rows(insertAt).Insert Shift:=xlDown
    ws.Cells(insertAt, 1).Value = label
    ws.Cells(insertAt, 1).Font.Bold = True
    ' Re-resolve field cell (supports A/B, C/D, E/F layouts).
    Dim valueCol As Long
    If Not FindFormField(sheetName, fieldName, fieldRow, valueCol) Then Exit Sub
    ws.Cells(insertAt, 2).Formula = _
        "=IF($" & ColLetter(valueCol) & "$" & fieldRow & "="""",""(none)"",""" & _
        keyName & ":""&$" & ColLetter(valueCol) & "$" & fieldRow & ")"
End Sub

Private Sub InsertCatalogColumn(ByVal sheetName As String, ByVal fieldName As String, ByVal afterName As String)
    Dim ws As Worksheet
    Dim hdr As Long
    Dim col As Long
    Dim afterCol As Long
    Dim notesCol As Long

    Set ws = Worksheets(sheetName)
    hdr = SheetHeaderRow(ws)
    If HeaderCol(ws, fieldName) > 0 Then Exit Sub

    If Len(afterName) > 0 Then
        afterCol = HeaderCol(ws, afterName)
        If afterCol > 0 Then
            col = afterCol + 1
        Else
            col = HeaderLastCol(ws) + 1
        End If
    Else
        notesCol = HeaderCol(ws, "notes")
        If notesCol > 0 Then
            col = notesCol
        Else
            col = HeaderLastCol(ws) + 1
        End If
    End If

    ws.Columns(col).Insert Shift:=xlToRight
    ws.Cells(hdr, col).Value = fieldName
    ws.Cells(hdr, col).Font.Bold = True
    ' Keep catalog data cells editable under sheet protection.
    On Error Resume Next
    ws.Cells(hdr, col).Locked = True
    ws.Range(ws.Cells(hdr + 1, col), ws.Cells(500, col)).Locked = False
    On Error GoTo 0
End Sub

' Click an option in columns C+, then run this macro (fills column B).
Public Sub PickSelectedOption()
    HandleEntryOptionClick ActiveSheet, ActiveCell
End Sub

Public Sub HandleEntryOptionClick(ByVal Sh As Object, ByVal Target As Range)
    Dim label As String
    Dim picked As String

    On Error GoTo safeExit
    If Sh Is Nothing Then GoTo safeExit
    If Not IsEntryFormSheet(Sh.Name) Then GoTo safeExit
    If Target Is Nothing Then GoTo safeExit
    If Target.Rows.Count <> 1 Or Target.Columns.Count <> 1 Then GoTo safeExit
    If Target.Column < ENTRY_OPTION_FIRST_COL Then GoTo safeExit

    picked = Trim$(CStr(Target.Value))
    If Len(picked) = 0 Then GoTo safeExit

    label = Trim$(CStr(Sh.Cells(Target.Row, 1).Value))
    If Not IsClickPickField(label) Then GoTo safeExit

    Application.EnableEvents = False
    Sh.Cells(Target.Row, 2).Value = picked
safeExit:
    Application.EnableEvents = True
End Sub

Private Function IsEntryFormSheet(ByVal sheetName As String) As Boolean
    If StrComp(sheetName, ENTRY_SHEET, vbTextCompare) = 0 Then IsEntryFormSheet = True
    If StrComp(sheetName, VIDEO_ENTRY_SHEET, vbTextCompare) = 0 Then IsEntryFormSheet = True
End Function

' Click-pick if Lists has a matching header (row 2) — new fields need no Case edit.
Private Function IsClickPickField(ByVal label As String) As Boolean
    Dim ws As Worksheet
    Dim c As Long
    Dim lastCol As Long
    Dim hdr As String

    IsClickPickField = False
    If Len(Trim$(label)) = 0 Then Exit Function
    On Error Resume Next
    Set ws = Worksheets(LISTS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastCol = 1
    For c = 1 To 64
        hdr = Trim$(CStr(ws.Cells(LISTS_HEADER_ROW, c).Value))
        If Len(hdr) > 0 Then lastCol = c
    Next c
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(LISTS_HEADER_ROW, c).Value)), label, vbTextCompare) = 0 Then
            IsClickPickField = True
            Exit Function
        End If
    Next c
End Function

Public Sub LoadFromFinderSelection()
    LoadPhotoMetadata "loadFromFinder", "LoadFromFinderSelection"
End Sub

Public Sub LoadFromPreviewSelection()
    LoadPhotoMetadata "loadFromPreview", "LoadFromPreviewSelection"
End Sub

' Trash front Preview image via Preview's own Edit → Move Selected Image to Trash.
Public Sub TrashFrontPreviewDocument()
    Dim raw As String
    Dim parts() As String
    Dim path As String
    Dim entryPath As String

#If Mac Then
    On Error GoTo appleScriptMissing
    ' Remember path first so we can clear Entry if it matches
    path = NormalizePath(Trim$(AppleScriptTask("StriderFinderMeta.scpt", "frontPreviewPOSIXPath", "")))
    raw = AppleScriptTask("StriderFinderMeta.scpt", "trashFrontPreview", "")
    On Error GoTo 0
#Else
    MsgBox "TrashFrontPreviewDocument is implemented for Mac Excel.", vbInformation
    Exit Sub
#End If

    parts = Split(raw, vbTab)
    If UBound(parts) >= 0 Then
        If parts(0) = "ERROR" Then
            If UBound(parts) >= 1 Then MsgBox parts(1), vbExclamation Else MsgBox "Trash failed.", vbExclamation
            Exit Sub
        End If
        If UBound(parts) >= 1 Then path = NormalizePath(parts(1))
    End If

    If Len(path) = 0 Then Exit Sub

    entryPath = NormalizePath(Trim$(CStr(FormValue(ENTRY_SHEET, "Path"))))
    If StrComp(entryPath, path, vbTextCompare) = 0 Then
        ClearFormFields ENTRY_SHEET, PhotoFieldList()
    End If
    entryPath = NormalizePath(Trim$(CStr(FormValue(VIDEO_ENTRY_SHEET, "Path"))))
    If StrComp(entryPath, path, vbTextCompare) = 0 Then
        ClearFormFields VIDEO_ENTRY_SHEET, VideoFieldList()
    End If
    Exit Sub

appleScriptMissing:
    MsgBox "Could not run StriderFinderMeta.scpt (trashFrontPreview)." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel." & vbCrLf & _
           "Also allow Automation control of Preview if macOS prompts.", vbCritical
End Sub

Private Sub LoadPhotoMetadata(ByVal appleHandler As String, ByVal macroName As String)
    Dim raw As String
    Dim parts() As String
    Dim meta() As String
    Dim path As String
    Dim fileName As String
    Dim ext As String
    Dim mime As String
    Dim dto As String
    Dim createDate As String
    Dim fileSize As String
    Dim shutter As String
    Dim aperture As String
    Dim isoVal As String
    Dim camera As String
    Dim mode As String
    Dim lens As String
    Dim stamp As String

#If Mac Then
    On Error GoTo appleScriptMissing
    raw = AppleScriptTask("StriderFinderMeta.scpt", appleHandler, "")
    On Error GoTo 0
#Else
    MsgBox macroName & " is implemented for Mac Excel.", vbInformation
    Exit Sub
#End If

    parts = Split(raw, vbTab)
    If UBound(parts) < 1 Then
        MsgBox "Unexpected metadata response.", vbExclamation
        Exit Sub
    End If
    If parts(0) = "ERROR" Then
        MsgBox parts(1), vbExclamation
        Exit Sub
    End If
    If UBound(parts) < 2 Then
        MsgBox "Metadata response was incomplete. Is exiftool installed?", vbExclamation
        Exit Sub
    End If

    path = NormalizePath(parts(1))
    If UBound(parts) >= 2 Then
        meta = Split(parts(2), "|")
    Else
        ReDim meta(0 To 0)
        meta(0) = ""
    End If

    fileName = MetaAt(meta, 0)
    ext = MetaAt(meta, 1)
    mime = LCase$(MetaAt(meta, 2))
    dto = MetaAt(meta, 3)
    createDate = MetaAt(meta, 4)
    fileSize = MetaAt(meta, 5)
    shutter = MetaAt(meta, 6)
    aperture = MetaAt(meta, 7)
    isoVal = MetaAt(meta, 8)
    camera = MetaAt(meta, 9)
    mode = MetaAt(meta, 10)
    lens = MetaAt(meta, 11)

    If Len(fileName) = 0 Then fileName = BaseName(path)
    stamp = FirstNonEmpty(dto, createDate)

    ApplyLoadedPhotoMeta path, fileName, ext, fileSize, stamp, shutter, aperture, isoVal, _
                         camera, mode, lens, meta
    ' Description NOTE: → Entry Note; CAPTION: → Append preview only (preserved on file).
    If UBound(parts) >= 3 Then
        SetEntryValue "Note", parts(3)
    End If
    If UBound(parts) >= 4 Then
        SetDescriptionCaptionPreview ENTRY_SHEET, parts(4)
    Else
        SetDescriptionCaptionPreview ENTRY_SHEET, ""
    End If

    Worksheets(ENTRY_SHEET).Activate
    MsgBox "Loaded metadata for:" & vbCrLf & fileName & vbCrLf & vbCrLf & _
           "Set tags, then run AppendPhotoRow.", vbInformation
    Exit Sub

appleScriptMissing:
    MsgBox "Could not run StriderFinderMeta.scpt (" & appleHandler & ")." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
End Sub

' meta pipe: …|serial(14)|vertPix(15)|horPix(16)|haveRaw(17)|seq(18)|seqID(19)|seqPos(20)
Private Sub ApplyLoadedPhotoMeta(ByVal path As String, ByVal fileName As String, _
                                 ByVal ext As String, ByVal fileSize As String, _
                                 ByVal stamp As String, ByVal shutter As String, _
                                 ByVal aperture As String, ByVal isoVal As String, _
                                 ByVal camera As String, ByVal mode As String, _
                                 ByVal lens As String, meta() As String)
    SetEntryValue "Path", path
    SetEntryValue "Name", fileName
    SetEntryValue "Type", UCase$(ext)
    SetEntryValue "FileSize", fileSize
    SetEntryValue "HaveRaw?", MetaAt(meta, 17)
    SetEntryValue "Date", ExtractDate(stamp)
    SetEntryValue "Time", ExtractTime(stamp)
    FormatEntryField "Date", "m/d/yy"
    FormatEntryField "Time", "[$-F400]h:mm:ss AM/PM"
    SetEntryValue "Speed", shutter
    SetEntryValue "f-stop", aperture
    SetEntryValue "ISO", isoVal
    SetEntryValue "Camera", camera
    SetEntryValue "Mode", mode
    SetEntryValue "Lens", lens
    SetEntryValue "VertPix", MetaAt(meta, 15)
    SetEntryValue "HorPix", MetaAt(meta, 16)
    ' Focus Distance = Sony FocusDistance2 (e.g. "1.654 m"); blank on phones.
    SetEntryValue "Focus Distance", MetaAt(meta, 12)
    ' GPS = GPSPosition when geotagged (phones); does not touch Locale.
    SetEntryValue "GPS", MetaAt(meta, 13)
    SetEntryValue "Serial Number", MetaAt(meta, 14)
    ' SequencePos is like 6/54 — must be text or Excel treats it as a fraction.
    SetEntryTextValue "Sequence", MetaAt(meta, 18)
    SetEntryTextValue "SequenceID", MetaAt(meta, 19)
    SetEntryTextValue "SequencePos", MetaAt(meta, 20)
End Sub

Public Sub AppendCatalogRow()
    AppendPhotoRow
End Sub

Public Sub AppendPhotoRow()
    Dim ws As Worksheet
    Dim path As String
    Dim fileName As String
    Dim destRow As Long

    EnsureColumns Worksheets(PHOTOS_SHEET), PhotoFieldList()
    Set ws = Worksheets(PHOTOS_SHEET)

    path = Trim$(CStr(FormValue(ENTRY_SHEET, "Path")))
    If Len(path) = 0 Then
        MsgBox "Load a Finder selection first (LoadFromFinderSelection).", vbExclamation
        Exit Sub
    End If

    path = NormalizePath(path)
    fileName = Trim$(CStr(FormValue(ENTRY_SHEET, "Name")))
    If Len(fileName) = 0 Then fileName = BaseName(path)

    destRow = FindRowByFileName(ws, fileName)
    If destRow = 0 Then destRow = NextDataRow(ws)
    CopyFormFieldsToRow ENTRY_SHEET, ws, destRow, PhotoFieldList()
    FormatCatalogField ws, destRow, "Date", "m/d/yy"
    FormatCatalogField ws, destRow, "Time", "[$-F400]h:mm:ss AM/PM"
    AddNameHyperlink ws, destRow, path, fileName
    WriteCatalogTagsToFile ENTRY_SHEET, path
    ClearFormFields ENTRY_SHEET, PhotoFieldList()
    ws.Activate
End Sub

' Bulk-append every frame sharing Entry's SequenceID.
' Seed user tags from Entry (intended: load/tag the 1/n frame first).
' Each frame keeps its own Path/Name/Date/Time/camera EXIF + SequencePos.
Public Sub AppendPhotoSequence()
    Dim ws As Worksheet
    Dim path As String
    Dim seqID As String
    Dim seqPos As String
    Dim seqFlag As String
    Dim raw As String
    Dim parts() As String
    Dim i As Long
    Dim memberPath As String
    Dim memberPos As String
    Dim seedLabels As Variant
    Dim seedVals() As String
    Dim nSeed As Long
    Dim label As String
    Dim countOk As Long
    Dim destRow As Long
    Dim fileName As String
    Dim ans As VbMsgBoxResult

    path = NormalizePath(Trim$(CStr(FormValue(ENTRY_SHEET, "Path"))))
    seqID = Trim$(CStr(FormValue(ENTRY_SHEET, "SequenceID")))
    seqPos = Trim$(CStr(FormValue(ENTRY_SHEET, "SequencePos")))
    seqFlag = LCase$(Trim$(CStr(FormValue(ENTRY_SHEET, "Sequence"))))

    If Len(path) = 0 Then
        MsgBox "Load the sequence's first frame (1/n) into Entry first.", vbExclamation
        Exit Sub
    End If
    If Len(seqID) = 0 Then
        MsgBox "Entry has no SequenceID. Load a sequence-tagged photo first.", vbExclamation
        Exit Sub
    End If
    If seqFlag <> "yes" Then
        ans = MsgBox("Entry Sequence is not ""yes"". Continue anyway?", vbYesNo + vbQuestion)
        If ans <> vbYes Then Exit Sub
    End If
    If Left$(seqPos, 2) <> "1/" And Len(seqPos) > 0 Then
        ans = MsgBox( _
            "Entry SequencePos is """ & seqPos & """ (expected 1/n)." & vbCrLf & _
            "Tags on Entry will still be applied to every frame." & vbCrLf & vbCrLf & _
            "Continue?", vbYesNo + vbQuestion)
        If ans <> vbYes Then Exit Sub
    End If

    ' Snapshot seed tags (user + notes); exclude per-frame identity/camera/sequence.
    seedLabels = PhotoFieldList()
    nSeed = 0
    ReDim seedVals(LBound(seedLabels) To UBound(seedLabels))
    For i = LBound(seedLabels) To UBound(seedLabels)
        label = CStr(seedLabels(i))
        If IsSequenceSeedField(label) Then
            seedVals(i) = CStr(FormValue(ENTRY_SHEET, label) & "")
            nSeed = nSeed + 1
        Else
            seedVals(i) = Chr$(0) ' marker: not a seed field
        End If
    Next i

#If Mac Then
    On Error GoTo appleScriptMissing
    raw = AppleScriptTask("StriderFinderMeta.scpt", "listSequenceMembers", path & vbTab & seqID)
    On Error GoTo 0
#Else
    MsgBox "AppendPhotoSequence is implemented for Mac Excel.", vbInformation
    Exit Sub
#End If

    parts = Split(raw, vbTab)
    If UBound(parts) < 0 Then
        MsgBox "Empty response listing sequence members.", vbExclamation
        Exit Sub
    End If
    If parts(0) = "ERROR" Then
        If UBound(parts) >= 1 Then
            MsgBox parts(1), vbExclamation
        Else
            MsgBox "Could not list sequence members.", vbExclamation
        End If
        Exit Sub
    End If
    If UBound(parts) < 2 Then
        MsgBox "No files found for SequenceID " & seqID & " in that folder." & vbCrLf & _
               "Run tag_sequences.py --write on the ingest folder first.", vbExclamation
        Exit Sub
    End If

    EnsureColumns Worksheets(PHOTOS_SHEET), PhotoFieldList()
    Set ws = Worksheets(PHOTOS_SHEET)
    Application.ScreenUpdating = False
    countOk = 0

    ' parts: OK, path1, pos1, path2, pos2, ...
    For i = 1 To UBound(parts) Step 2
        If i + 1 > UBound(parts) Then Exit For
        memberPath = NormalizePath(Trim$(parts(i)))
        memberPos = Trim$(parts(i + 1))
        If Len(memberPath) = 0 Then GoTo nextMember

        If Not FillEntryAutoFromPath(memberPath) Then GoTo nextMember

        ' Restore seed tags; Sequence* come from the file / member list below.
        Dim j As Long
        For j = LBound(seedLabels) To UBound(seedLabels)
            label = CStr(seedLabels(j))
            If seedVals(j) <> Chr$(0) Then
                If StrComp(label, "Note", vbTextCompare) = 0 Or StrComp(label, "notes", vbTextCompare) = 0 Then
                    SetEntryTextValue label, seedVals(j)
                Else
                    SetFormValue ENTRY_SHEET, label, seedVals(j)
                End If
            End If
        Next j
        ' Cluster id + this frame's position
        SetEntryTextValue "Sequence", "yes"
        SetEntryTextValue "SequenceID", seqID
        If Len(memberPos) > 0 Then SetEntryTextValue "SequencePos", memberPos

        fileName = Trim$(CStr(FormValue(ENTRY_SHEET, "Name")))
        If Len(fileName) = 0 Then fileName = BaseName(memberPath)
        destRow = FindRowByFileName(ws, fileName)
        If destRow = 0 Then destRow = NextDataRow(ws)
        CopyFormFieldsToRow ENTRY_SHEET, ws, destRow, PhotoFieldList()
        FormatCatalogField ws, destRow, "Date", "m/d/yy"
        FormatCatalogField ws, destRow, "Time", "[$-F400]h:mm:ss AM/PM"
        AddNameHyperlink ws, destRow, memberPath, fileName
        WriteCatalogTagsToFile ENTRY_SHEET, memberPath
        countOk = countOk + 1
nextMember:
    Next i

    ClearFormFields ENTRY_SHEET, PhotoFieldList()
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Appended " & countOk & " frame(s) for SequenceID " & seqID & "." & vbCrLf & _
           "Shared tags came from Entry (1/n seed); each row kept its own SequencePos.", vbInformation
    Exit Sub

appleScriptMissing:
    Application.ScreenUpdating = True
    MsgBox "Could not run StriderFinderMeta.scpt (listSequenceMembers)." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
End Sub

' User tags applied to every frame; not camera/identity/sequence fields.
Private Function IsSequenceSeedField(ByVal label As String) As Boolean
    Select Case label
        Case "Name", "Path", "HaveRaw?", "Date", "Time", "Type", "FileSize", _
             "VertPix", "HorPix", _
             "Camera", "Mode", "Lens", "Speed", "f-stop", "ISO", _
             "Focus Distance", "GPS", "Serial Number", _
             "Duration", "FrameRate", "Resolution", _
             "Sequence", "SequenceID", "SequencePos"
            IsSequenceSeedField = False
        Case Else
            IsSequenceSeedField = IsFormFieldLabel(label)
    End Select
End Function

' Fill Entry auto/camera/sequence fields from a path (no MsgBox). False on failure.
Private Function FillEntryAutoFromPath(ByVal posixPath As String) As Boolean
    Dim raw As String
    Dim parts() As String
    Dim meta() As String
    Dim path As String
    Dim fileName As String
    Dim ext As String
    Dim dto As String
    Dim createDate As String
    Dim fileSize As String
    Dim shutter As String
    Dim aperture As String
    Dim isoVal As String
    Dim camera As String
    Dim mode As String
    Dim lens As String
    Dim stamp As String

    FillEntryAutoFromPath = False
#If Mac Then
    On Error GoTo failFill
    raw = AppleScriptTask("StriderFinderMeta.scpt", "readPath", posixPath)
    On Error GoTo 0
#Else
    Exit Function
#End If

    parts = Split(raw, vbTab)
    If UBound(parts) < 2 Then Exit Function
    If parts(0) = "ERROR" Then Exit Function

    path = NormalizePath(parts(1))
    If UBound(parts) >= 2 Then
        meta = Split(parts(2), "|")
    Else
        ReDim meta(0 To 0)
        meta(0) = ""
    End If

    fileName = MetaAt(meta, 0)
    ext = MetaAt(meta, 1)
    dto = MetaAt(meta, 3)
    createDate = MetaAt(meta, 4)
    fileSize = MetaAt(meta, 5)
    shutter = MetaAt(meta, 6)
    aperture = MetaAt(meta, 7)
    isoVal = MetaAt(meta, 8)
    camera = MetaAt(meta, 9)
    mode = MetaAt(meta, 10)
    lens = MetaAt(meta, 11)
    If Len(fileName) = 0 Then fileName = BaseName(path)
    stamp = FirstNonEmpty(dto, createDate)

    ApplyLoadedPhotoMeta path, fileName, ext, fileSize, stamp, shutter, aperture, isoVal, _
                         camera, mode, lens, meta
    If UBound(parts) >= 3 Then
        SetEntryValue "Note", parts(3)
    End If
    If UBound(parts) >= 4 Then
        SetDescriptionCaptionPreview ENTRY_SHEET, parts(4)
    Else
        SetDescriptionCaptionPreview ENTRY_SHEET, ""
    End If
    FillEntryAutoFromPath = True
    Exit Function
failFill:
    FillEntryAutoFromPath = False
End Function

Public Sub LoadVideoFromFinderSelection()
    LoadVideoMetadata "loadVideoFromFinder", "LoadVideoFromFinderSelection"
End Sub

Public Sub LoadVideoFromPreviewSelection()
    LoadVideoMetadata "loadVideoFromPreview", "LoadVideoFromPreviewSelection"
End Sub

Private Sub LoadVideoMetadata(ByVal appleHandler As String, ByVal macroName As String)
    Dim raw As String
    Dim parts() As String
    Dim meta() As String
    Dim path As String
    Dim fileName As String
    Dim ext As String
    Dim mime As String
    Dim createDate As String
    Dim mediaCreate As String
    Dim fileSize As String
    Dim duration As String
    Dim fps As String
    Dim width As String
    Dim height As String
    Dim deviceModel As String
    Dim model As String
    Dim shutter As String
    Dim aperture As String
    Dim isoVal As String
    Dim stamp As String
    Dim camera As String
    Dim resolution As String

#If Mac Then
    On Error GoTo appleScriptMissing
    raw = AppleScriptTask("StriderFinderMeta.scpt", appleHandler, "")
    On Error GoTo 0
#Else
    MsgBox macroName & " is implemented for Mac Excel.", vbInformation
    Exit Sub
#End If

    parts = Split(raw, vbTab)
    If UBound(parts) < 1 Then
        MsgBox "Unexpected metadata response.", vbExclamation
        Exit Sub
    End If
    If parts(0) = "ERROR" Then
        MsgBox parts(1), vbExclamation
        Exit Sub
    End If
    If UBound(parts) < 2 Then
        MsgBox "Metadata response was incomplete. Is exiftool installed?", vbExclamation
        Exit Sub
    End If

    path = NormalizePath(parts(1))
    If UBound(parts) >= 2 Then
        meta = Split(parts(2), "|")
    Else
        ReDim meta(0 To 0)
        meta(0) = ""
    End If

    fileName = MetaAt(meta, 0)
    ext = MetaAt(meta, 1)
    mime = LCase$(MetaAt(meta, 2))
    createDate = MetaAt(meta, 3)
    mediaCreate = MetaAt(meta, 4)
    fileSize = MetaAt(meta, 5)
    duration = MetaAt(meta, 6)
    fps = MetaAt(meta, 7)
    width = MetaAt(meta, 8)
    height = MetaAt(meta, 9)
    deviceModel = MetaAt(meta, 10)
    model = MetaAt(meta, 11)
    shutter = MetaAt(meta, 12)
    aperture = MetaAt(meta, 13)
    isoVal = MetaAt(meta, 14)

    If Len(fileName) = 0 Then fileName = BaseName(path)
    stamp = FirstNonEmpty(createDate, mediaCreate)
    camera = FirstNonEmpty(deviceModel, model)
    If Len(width) > 0 And Len(height) > 0 Then
        resolution = width & "x" & height
    Else
        resolution = ""
    End If

    SetFormValue VIDEO_ENTRY_SHEET, "Path", path
    SetFormValue VIDEO_ENTRY_SHEET, "Name", fileName
    SetFormValue VIDEO_ENTRY_SHEET, "Type", UCase$(ext)
    SetFormValue VIDEO_ENTRY_SHEET, "FileSize", fileSize
    SetFormValue VIDEO_ENTRY_SHEET, "Date", ExtractDate(stamp)
    SetFormValue VIDEO_ENTRY_SHEET, "Time", ExtractTime(stamp)
    FormatFormField VIDEO_ENTRY_SHEET, "Date", "m/d/yy"
    FormatFormField VIDEO_ENTRY_SHEET, "Time", "[$-F400]h:mm:ss AM/PM"
    SetFormValue VIDEO_ENTRY_SHEET, "Duration", duration
    SetFormValue VIDEO_ENTRY_SHEET, "FrameRate", fps
    SetFormValue VIDEO_ENTRY_SHEET, "Resolution", resolution
    SetFormValue VIDEO_ENTRY_SHEET, "Camera", camera
    SetFormValue VIDEO_ENTRY_SHEET, "Speed", shutter
    SetFormValue VIDEO_ENTRY_SHEET, "f-stop", aperture
    SetFormValue VIDEO_ENTRY_SHEET, "ISO", isoVal

    Worksheets(VIDEO_ENTRY_SHEET).Activate
    MsgBox "Loaded video metadata for:" & vbCrLf & fileName & vbCrLf & vbCrLf & _
           "Set tags, then run AppendVideoRow.", vbInformation
    Exit Sub

appleScriptMissing:
    MsgBox "Could not run StriderFinderMeta.scpt (" & appleHandler & ")." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
End Sub

Public Sub AppendVideoRow()
    Dim ws As Worksheet
    Dim path As String
    Dim fileName As String
    Dim destRow As Long

    EnsureColumns Worksheets(VIDEOS_SHEET), VideoFieldList()
    Set ws = Worksheets(VIDEOS_SHEET)

    path = Trim$(CStr(FormValue(VIDEO_ENTRY_SHEET, "Path")))
    If Len(path) = 0 Then
        MsgBox "Load a Finder selection first (LoadVideoFromFinderSelection).", vbExclamation
        Exit Sub
    End If

    path = NormalizePath(path)
    fileName = Trim$(CStr(FormValue(VIDEO_ENTRY_SHEET, "Name")))
    If Len(fileName) = 0 Then fileName = BaseName(path)

    destRow = FindRowByFileName(ws, fileName)
    If destRow = 0 Then destRow = NextDataRow(ws)
    CopyFormFieldsToRow VIDEO_ENTRY_SHEET, ws, destRow, VideoFieldList()
    FormatCatalogField ws, destRow, "Date", "m/d/yy"
    FormatCatalogField ws, destRow, "Time", "[$-F400]h:mm:ss AM/PM"
    AddNameHyperlink ws, destRow, path, fileName
    WriteCatalogTagsToFile VIDEO_ENTRY_SHEET, path
    ClearFormFields VIDEO_ENTRY_SHEET, VideoFieldList()
    ws.Activate
End Sub

Public Sub OpenSelectedInPreview()
    Dim paths As Collection
    Dim joined As String
    Dim result As String

    Set paths = CollectSelectedPaths()
    If paths.Count = 0 Then
        MsgBox "Select Name/Path cells or rows on Photos or Videos.", vbExclamation
        Exit Sub
    End If

    joined = JoinPaths(paths)
#If Mac Then
    On Error GoTo appleScriptMissing
    result = AppleScriptTask("StriderOpenPreview.scpt", "openInPreview", joined)
    If Len(result) > 0 Then MsgBox result, vbExclamation
    Exit Sub
appleScriptMissing:
    MsgBox "Could not run StriderOpenPreview.scpt." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
#Else
    Dim p As Variant
    For Each p In paths
        ThisWorkbook.FollowHyperlink CStr(p)
    Next p
#End If
End Sub

Public Sub RevealSelectedInFinder()
    Dim paths As Collection
    Dim joined As String
    Dim result As String

    Set paths = CollectSelectedPaths()
    If paths.Count = 0 Then
        MsgBox "Select Name/Path cells or rows on Photos or Videos.", vbExclamation
        Exit Sub
    End If

    joined = JoinPaths(paths)
#If Mac Then
    On Error GoTo appleScriptMissing
    result = AppleScriptTask("StriderRevealFinder.scpt", "revealInFinder", joined)
    If Len(result) > 0 Then MsgBox result, vbExclamation
    Exit Sub
appleScriptMissing:
    MsgBox "Could not run StriderRevealFinder.scpt." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
#Else
    Dim p As Variant
    For Each p In paths
        Shell "explorer /select,""" & CStr(p) & """", vbNormalFocus
    Next p
#End If
End Sub

' Caption any time from Photos: select Name/Path cell(s) or whole row(s) → Ctrl+Shift+Y
' (or Macros → CaptionSelectedPhotos). Writes Photos Caption column + embeds CAPTION: in the file.
' Click a Caption cell to edit in a Mac dialog popup. Default provider: Gemini Flash.
Public Sub CaptionSelectedPhotos()
    Dim paths As Collection
    Dim p As Variant
    Dim path As String
    Dim raw As String
    Dim parts() As String
    Dim caption As String
    Dim fileName As String
    Dim photosWs As Worksheet
    Dim destRow As Long
    Dim okCount As Long
    Dim errCount As Long
    Dim missCount As Long
    Dim embedFail As Long
    Dim sheetName As String
    Dim i As Long

    If Selection Is Nothing Then
        MsgBox "On the Photos sheet, select a Name (filename) cell, Path cell, or row(s), then run again.", vbExclamation
        Exit Sub
    End If
    sheetName = Selection.Worksheet.Name
    If StrComp(sheetName, PHOTOS_SHEET, vbTextCompare) <> 0 Then
        MsgBox "Switch to the Photos sheet, select Name/Path cell(s) or row(s)," & vbCrLf & _
               "then Ctrl+Shift+Y (or Macros -> CaptionSelectedPhotos).", vbExclamation
        Exit Sub
    End If

    Set photosWs = Worksheets(PHOTOS_SHEET)
    Set paths = CollectSelectedPaths()
    If paths.Count = 0 Then
        MsgBox "Could not resolve a file path from the selection." & vbCrLf & _
               "Select the Name or Path cell (or the whole row) for a Photos entry that has a Path.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ensureFail
    EnsureCaptionColumn photosWs
    On Error GoTo 0

    okCount = 0
    errCount = 0
    missCount = 0
    embedFail = 0
    Application.StatusBar = "Captioning " & paths.Count & " photo(s)... (Ctrl+Shift+Y)"

    For Each p In paths
        path = NormalizePath(CStr(p))
        fileName = BaseName(path)
        destRow = FindPhotosRowForPath(photosWs, path)
        If destRow = 0 Then
            missCount = missCount + 1
            GoTo nextPath
        End If
#If Mac Then
        On Error GoTo appleScriptMissing
        raw = AppleScriptTask("StriderCaption.scpt", "captionPhoto", path & vbTab & DEFAULT_CAPTION_PROVIDER)
        On Error GoTo 0
#Else
        MsgBox "CaptionSelectedPhotos is implemented for Mac Excel.", vbInformation
        Application.StatusBar = False
        Exit Sub
#End If
        parts = Split(raw, vbTab)
        If UBound(parts) < 0 Then
            errCount = errCount + 1
            GoTo nextPath
        End If
        If parts(0) = "ERROR" Then
            errCount = errCount + 1
            If UBound(parts) >= 1 Then
                MsgBox "Caption failed for " & fileName & ":" & vbCrLf & parts(1), vbExclamation
            End If
            GoTo nextPath
        End If
        If UBound(parts) < 4 Then
            errCount = errCount + 1
            GoTo nextPath
        End If

        caption = parts(4)
        If UBound(parts) > 4 Then
            For i = 5 To UBound(parts)
                caption = caption & vbTab & parts(i)
            Next i
        End If

        If Not WriteCaptionToPhotosRow(photosWs, destRow, path, caption) Then
            embedFail = embedFail + 1
        End If
        okCount = okCount + 1
nextPath:
    Next p

    Application.StatusBar = False
    photosWs.Activate
    MsgBox "Captions written: " & okCount & "  /  failed: " & errCount & _
           "  /  no Photos row: " & missCount & _
           "  /  file embed failed: " & embedFail & vbCrLf & _
           "Stored in Photos Caption column; embedded as MWG:Description CAPTION: (NOTE: preserved)." & vbCrLf & _
           "Click a Caption cell to edit. Provider: " & DEFAULT_CAPTION_PROVIDER & "." & vbCrLf & vbCrLf & _
           "Any time: Photos -> select Name/row(s) -> Ctrl+Shift+Y.", _
           vbInformation
    Exit Sub

appleScriptMissing:
    Application.StatusBar = False
    MsgBox "Could not run StriderCaption.scpt." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
    Exit Sub

ensureFail:
    Application.StatusBar = False
    MsgBox "Could not prepare the Photos Caption column:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "If the workbook is protected, run macro ReapplyUiOnlyProtection, then try again.", vbCritical
End Sub

' Photos / Videos SelectionChange: Mac dialog to edit Caption (UserForms unavailable on Mac Excel).
Public Sub HandleCatalogCaptionClick(ByVal Sh As Object, ByVal Target As Range)
    Dim capCol As Long
    Dim caption As String
    Dim edited As String
    Dim cancelled As Boolean
    Dim eventsWas As Boolean

    On Error GoTo safeExit
    If gCaptionEditorOpen Then GoTo safeExit
    If Sh Is Nothing Then GoTo safeExit
    If Target Is Nothing Then GoTo safeExit
    If Target.Rows.Count <> 1 Or Target.Columns.Count <> 1 Then GoTo safeExit
    If StrComp(Sh.Name, PHOTOS_SHEET, vbTextCompare) <> 0 And _
       StrComp(Sh.Name, VIDEOS_SHEET, vbTextCompare) <> 0 Then GoTo safeExit
    If Target.Row <= SheetHeaderRow(Sh) Then GoTo safeExit

    ' Resolve Caption column first — do not mutate the sheet until we know this is a Caption click.
    capCol = HeaderCol(Sh, "Caption")
    If capCol = 0 Then
        eventsWas = Application.EnableEvents
        Application.EnableEvents = False
        EnsureCaptionColumn Sh
        Application.EnableEvents = eventsWas
        capCol = HeaderCol(Sh, "Caption")
    End If
    If capCol = 0 Or Target.Column <> capCol Then GoTo safeExit

    caption = ""
    On Error Resume Next
    caption = CStr(Sh.Cells(Target.Row, capCol).Value2)
    If Len(caption) = 0 Then caption = CStr(Sh.Cells(Target.Row, capCol).Text)
    On Error GoTo safeExit

    gCaptionEditorOpen = True
    eventsWas = Application.EnableEvents
    Application.EnableEvents = False
    edited = PromptEditCaption(caption, cancelled)
    Application.EnableEvents = eventsWas
    If Not cancelled Then
        CommitCaptionEdit Sh.Name, Target.Row, edited
    End If
safeExit:
    gCaptionEditorOpen = False
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' Mac AppleScript dialog (Save/Cancel). Returns edited text; cancelled=True if Cancel.
Private Function PromptEditCaption(ByVal currentText As String, ByRef cancelled As Boolean) As String
    Dim raw As String
    Dim parts() As String
    Dim i As Long
    Dim out As String

    cancelled = True
    PromptEditCaption = currentText
#If Mac Then
    On Error GoTo promptFail
    raw = AppleScriptTask("StriderCaption.scpt", "editCaption", currentText)
    On Error GoTo 0
    If Len(raw) = 0 Then
        MsgBox "Caption editor returned no response." & vbCrLf & _
               "Run scripts/install_mac_helpers.sh, restart Excel, then try again." & vbCrLf & _
               "Or run macro FixEnableEvents.", vbExclamation
        Exit Function
    End If
    parts = Split(raw, vbTab)
    If UBound(parts) < 0 Then Exit Function
    If StrComp(parts(0), "CANCEL", vbTextCompare) = 0 Then Exit Function
    If StrComp(parts(0), "ERROR", vbTextCompare) = 0 Then
        If UBound(parts) >= 1 Then
            MsgBox "Could not open caption editor:" & vbCrLf & parts(1) & vbCrLf & vbCrLf & _
                   "Run scripts/install_mac_helpers.sh, then restart Excel.", vbExclamation
        End If
        Exit Function
    End If
    If StrComp(parts(0), "OK", vbTextCompare) <> 0 Then Exit Function
    out = ""
    If UBound(parts) >= 1 Then out = parts(1)
    If UBound(parts) > 1 Then
        For i = 2 To UBound(parts)
            out = out & vbTab & parts(i)
        Next i
    End If
    cancelled = False
    PromptEditCaption = out
    Exit Function
promptFail:
    MsgBox "Could not run StriderCaption.scpt (editCaption)." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
#Else
    MsgBox "Caption editing is implemented for Mac Excel.", vbInformation
#End If
End Function

' Save Caption cell + re-embed CAPTION: in the file.
Public Sub CommitCaptionEdit(ByVal sheetName As String, ByVal rowNum As Long, ByVal captionText As String)
    Dim ws As Worksheet
    Dim path As String
    Dim caption As String

    On Error GoTo failCommit
    Set ws = Worksheets(sheetName)
    EnsureCaptionColumn ws
    caption = Trim$(captionText)
    path = NormalizePath(PathFromCatalogRow(ws, rowNum))
    SetCellByHeader ws, rowNum, "Caption", caption
    FormatCaptionCell ws, rowNum
    If Len(path) > 0 Then
        If Not EmbedCaptionInFile(path, caption) Then
            MsgBox "Caption saved in the catalog, but file embed failed for:" & vbCrLf & path, vbExclamation
        End If
    End If
    Exit Sub
failCommit:
    MsgBox "Could not save caption:" & vbCrLf & Err.Description, vbCritical
End Sub

' One-shot: copy old Captions-sheet rows into Photos Caption (by Name/Path).
Public Sub MigrateCaptionsSheetToPhotos()
    Dim capWs As Worksheet
    Dim photosWs As Worksheet
    Dim hdr As Long
    Dim lastRow As Long
    Dim r As Long
    Dim destRow As Long
    Dim path As String
    Dim fileName As String
    Dim caption As String
    Dim okCount As Long
    Dim missCount As Long
    Dim nameCol As Long
    Dim pathCol As Long
    Dim captionCol As Long

    On Error Resume Next
    Set capWs = Worksheets(CAPTIONS_SHEET)
    On Error GoTo 0
    If capWs Is Nothing Then
        MsgBox "No Captions sheet found. Nothing to migrate.", vbInformation
        Exit Sub
    End If

    Set photosWs = Worksheets(PHOTOS_SHEET)
    EnsureCaptionColumn photosWs
    hdr = CAPTIONS_HEADER_ROW
    nameCol = HeaderCol(capWs, "Name")
    pathCol = HeaderCol(capWs, "Path")
    captionCol = HeaderCol(capWs, "Caption")
    If captionCol = 0 Then
        MsgBox "Captions sheet has no Caption column.", vbExclamation
        Exit Sub
    End If

    lastRow = CatalogLastDataRow(capWs)
    okCount = 0
    missCount = 0
    For r = hdr + 1 To lastRow
        caption = Trim$(CStr(capWs.Cells(r, captionCol).Value))
        If Len(caption) = 0 Then GoTo nextCapRow
        path = ""
        fileName = ""
        If pathCol > 0 Then path = NormalizePath(Trim$(CStr(capWs.Cells(r, pathCol).Value)))
        If nameCol > 0 Then fileName = Trim$(CStr(capWs.Cells(r, nameCol).Value))
        destRow = 0
        If Len(path) > 0 Then destRow = FindPhotosRowForPath(photosWs, path)
        If destRow = 0 And Len(fileName) > 0 Then destRow = FindRowByFileName(photosWs, fileName)
        If destRow = 0 Then
            missCount = missCount + 1
            GoTo nextCapRow
        End If
        SetCellByHeader photosWs, destRow, "Caption", caption
        FormatCaptionCell photosWs, destRow
        okCount = okCount + 1
nextCapRow:
    Next r

    MsgBox "Migrated " & okCount & " caption(s) into Photos." & vbCrLf & _
           "No matching Photos row: " & missCount & vbCrLf & vbCrLf & _
           "You can delete the Captions sheet when satisfied.", vbInformation
End Sub

Private Sub EnsureCaptionColumn(ByVal ws As Worksheet)
    Dim capCol As Long
    Dim hdr As Long
    EnsureColumns ws, Array("Caption")
    capCol = HeaderCol(ws, "Caption")
    If capCol > 0 Then
        hdr = SheetHeaderRow(ws)
        On Error Resume Next
        ws.Columns(capCol).WrapText = True
        ws.Columns(capCol).ColumnWidth = 48
        ' Keep Caption selectable under EnableSelection=xlUnlockedCells (do not re-lock whole sheet here).
        ws.Cells(hdr, capCol).Locked = True
        ws.Range(ws.Cells(hdr + 1, capCol), ws.Cells(500, capCol)).Locked = False
        On Error GoTo 0
    End If
End Sub

' Photos/Videos: header locked, data unlocked (Name/Caption/etc. must be selectable).
Private Sub UnlockCatalogListDataArea(ByVal ws As Worksheet)
    Dim hdr As Long
    Dim lastCol As Long
    If ws Is Nothing Then Exit Sub
    If StrComp(ws.Name, PHOTOS_SHEET, vbTextCompare) <> 0 And _
       StrComp(ws.Name, VIDEOS_SHEET, vbTextCompare) <> 0 Then Exit Sub
    hdr = SheetHeaderRow(ws)
    lastCol = HeaderLastCol(ws)
    If lastCol < 1 Then lastCol = 1
    If lastCol < 40 Then lastCol = 40
    On Error Resume Next
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, lastCol)).Locked = True
    ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(500, lastCol)).Locked = False
    On Error GoTo 0
End Sub

Private Function FindPhotosRowForPath(ByVal ws As Worksheet, ByVal path As String) As Long
    Dim pathCol As Long
    Dim r As Long
    Dim lastRow As Long
    Dim cellPath As String
    Dim want As String

    FindPhotosRowForPath = 0
    want = NormalizePath(path)
    If Len(want) = 0 Then Exit Function
    pathCol = HeaderCol(ws, "Path")
    If pathCol = 0 Then
        FindPhotosRowForPath = FindRowByFileName(ws, BaseName(want))
        Exit Function
    End If
    lastRow = CatalogLastDataRow(ws)
    For r = SheetHeaderRow(ws) + 1 To lastRow
        cellPath = NormalizePath(Trim$(CStr(ws.Cells(r, pathCol).Value)))
        If StrComp(cellPath, want, vbTextCompare) = 0 Then
            FindPhotosRowForPath = r
            Exit Function
        End If
    Next r
    FindPhotosRowForPath = FindRowByFileName(ws, BaseName(want))
End Function

' Writes Caption cell + embeds; returns True if embed OK (or no Mac embed attempted).
Private Function WriteCaptionToPhotosRow(ByVal ws As Worksheet, ByVal destRow As Long, _
                                         ByVal path As String, ByVal caption As String) As Boolean
    SetCellByHeader ws, destRow, "Caption", caption
    FormatCaptionCell ws, destRow
    WriteCaptionToPhotosRow = EmbedCaptionInFile(path, caption)
End Function

Private Sub FormatCaptionCell(ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim capCol As Long
    capCol = HeaderCol(ws, "Caption")
    If capCol = 0 Then Exit Sub
    On Error Resume Next
    ws.Cells(rowNum, capCol).WrapText = True
    ws.Cells(rowNum, capCol).Locked = False
    On Error GoTo 0
End Sub

Private Function EmbedCaptionInFile(ByVal path As String, ByVal caption As String) As Boolean
    Dim embedRaw As String
    Dim embedParts() As String

    EmbedCaptionInFile = True
    If Len(NormalizePath(path)) = 0 Then Exit Function
#If Mac Then
    On Error GoTo embedFail
    embedRaw = AppleScriptTask("StriderFinderMeta.scpt", "writeCaption", path & vbTab & caption)
    On Error GoTo 0
    embedParts = Split(embedRaw, vbTab)
    If UBound(embedParts) < 0 Or embedParts(0) = "ERROR" Then
        EmbedCaptionInFile = False
    End If
    Exit Function
embedFail:
    EmbedCaptionInFile = False
#End If
End Function

Private Function CaptionFieldList() As Variant
    CaptionFieldList = Array( _
        "Name", "Path", "Provider", "Model", "Words", "Generated", "Caption")
End Function

Private Sub EnsureCaptionsSheet()
    Dim ws As Worksheet
    Dim fields As Variant
    Dim i As Long
    Dim hdr As Long
    Dim structureOn As Boolean
    Dim created As Boolean

    On Error Resume Next
    Set ws = Worksheets(CAPTIONS_SHEET)
    On Error GoTo 0

    created = False
    If ws Is Nothing Then
        ' Workbook structure lock blocks adding sheets — lift briefly if needed.
        structureOn = ThisWorkbook.ProtectStructure
        If structureOn Then UnprotectWorkbookStructure
        On Error Resume Next
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        If Not ws Is Nothing Then
            ws.Name = CAPTIONS_SHEET
            created = True
        End If
        On Error GoTo 0
        If ws Is Nothing Then
            If structureOn Then ProtectWorkbookStructure
            Err.Raise 1004, "EnsureCaptionsSheet", _
                "Could not create Captions sheet (workbook structure may still be protected)."
        End If
    End If

    ' Sheet protection without UserInterfaceOnly blocks VBA writes after reopen.
    UnprotectSheet ws

    hdr = CAPTIONS_HEADER_ROW
    ws.Cells(1, 1).Value = "Captioning prompt"
    ws.Cells(1, 1).Font.Name = "Times New Roman"
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Bold = True

    ' Keep in sync with catalog/macros/caption_prompt.txt (and scripts/caption_photo.py).
    ws.Cells(2, 1).Value = CaptionPromptText()
    ws.Cells(2, 1).Font.Name = "Times New Roman"
    ws.Cells(2, 1).Font.Size = 12
    ws.Cells(2, 1).WrapText = True
    On Error Resume Next
    ws.Range("A2:G2").Merge
    ws.Rows(2).RowHeight = 150
    On Error GoTo 0

    ws.Cells(3, 1).Value = "Any time on Photos: select Name/Path cell(s) or row(s) -> Ctrl+Shift+Y (CaptionSelectedPhotos). Re-run to refresh a caption. Default provider: Gemini Flash (google)."
    ws.Cells(3, 1).Font.Name = "Times New Roman"
    ws.Cells(3, 1).Font.Size = 12
    ws.Cells(3, 1).Font.Italic = True

    fields = CaptionFieldList()
    For i = LBound(fields) To UBound(fields)
        If Len(Trim$(CStr(ws.Cells(hdr, i - LBound(fields) + 1).Value))) = 0 Then
            ws.Cells(hdr, i - LBound(fields) + 1).Value = CStr(fields(i))
        End If
        ws.Cells(hdr, i - LBound(fields) + 1).Font.Name = "Times New Roman"
        ws.Cells(hdr, i - LBound(fields) + 1).Font.Size = 14
        ws.Cells(hdr, i - LBound(fields) + 1).Font.Bold = True
    Next i
    EnsureColumns ws, CaptionFieldList()
    UnlockCaptionsDataArea ws
    ProtectCaptionsSheet ws
    If structureOn Or created Then ProtectWorkbookStructure
End Sub

Private Sub UnlockCaptionsDataArea(ByVal ws As Worksheet)
    Dim hdr As Long
    On Error Resume Next
    hdr = CAPTIONS_HEADER_ROW
    UnprotectSheet ws
    ws.Rows(hdr).Locked = True
    ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(500, 20)).Locked = False
    On Error GoTo 0
End Sub

Private Sub UnprotectWorkbookStructure()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=""
    On Error GoTo 0
End Sub

Private Sub ProtectWorkbookStructure()
    On Error Resume Next
    ThisWorkbook.Protect Password:="", Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

Private Sub UnprotectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo 0
End Sub

Private Sub ProtectCaptionsSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        UserInterfaceOnly:=True, _
        AllowFormattingCells:=True, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
        AllowInsertingColumns:=True, AllowInsertingRows:=True, _
        AllowDeletingColumns:=False, AllowDeletingRows:=True, _
        AllowSorting:=True, AllowFiltering:=True
    ws.EnableSelection = xlUnlockedCells
    On Error GoTo 0
End Sub

' Re-apply protection so VBA can write (UserInterfaceOnly is session-only).
Public Sub ReapplyUiOnlyProtection()
    Dim ws As Worksheet
    Dim nm As String

    On Error Resume Next
    UnprotectWorkbookStructure
    For Each ws In ThisWorkbook.Worksheets
        nm = ws.Name
        UnprotectSheet ws
        Select Case nm
            Case ENTRY_SHEET, VIDEO_ENTRY_SHEET
                ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    UserInterfaceOnly:=True, _
                    AllowFormattingCells:=True, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
                    AllowInsertingColumns:=True, AllowInsertingRows:=True, _
                    AllowDeletingColumns:=False, AllowDeletingRows:=False, _
                    AllowSorting:=True, AllowFiltering:=True
                ws.EnableSelection = xlUnlockedCells
            Case PHOTOS_SHEET, VIDEOS_SHEET, LISTS_SHEET
                If StrComp(nm, PHOTOS_SHEET, vbTextCompare) = 0 Or _
                   StrComp(nm, VIDEOS_SHEET, vbTextCompare) = 0 Then
                    StripNameHyperlinksOnSheet ws
                    UnlockCatalogListDataArea ws
                End If
                ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    UserInterfaceOnly:=True, _
                    AllowFormattingCells:=True, AllowFormattingColumns:=False, AllowFormattingRows:=False, _
                    AllowInsertingColumns:=True, AllowInsertingRows:=True, _
                    AllowDeletingColumns:=False, AllowDeletingRows:=True, _
                    AllowSorting:=True, AllowFiltering:=True
                ws.EnableSelection = xlUnlockedCells
            Case CAPTIONS_SHEET
                ProtectCaptionsSheet ws
            Case Else
                ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    UserInterfaceOnly:=True, _
                    AllowFormattingCells:=False, AllowFormattingColumns:=False, AllowFormattingRows:=False
                ws.EnableSelection = xlNoRestrictions
        End Select
    Next ws
    ProtectWorkbookStructure
    On Error GoTo 0
End Sub

Private Function CaptionPromptText() As String
    CaptionPromptText = _
        "You are writing a single Scientific American–style photo caption." & vbLf & vbLf & _
        "Requirements:" & vbLf & _
        "- About 60 words (50–70 is fine). Prefer one compact paragraph." & vbLf & _
        "- Use plain, straightforward descriptive prose — clear and factual, as in a Scientific American article. Avoid lyrical, dramatic, or flowery wording." & vbLf & _
        "- Ground the caption in what is actually visible in the image (organisms, leaf, ripples, reflections, light on water). Metadata may name species/place/behavior; use it only when it fits what you see." & vbLf & _
        "- Photographer notes are trusted field observations — incorporate them when they fit the image (do not quote them verbatim unless natural)." & vbLf & _
        "- Note any patterns of color, shape, or ripples." & vbLf & _
        "- Be on the lookout for small water striders and what they might be doing. Before writing, scan the entire frame and count every visible strider (they can be tiny and easy to miss). If more than one is present, say how many and briefly what they appear to be doing." & vbLf & _
        "- When water-surface optics are visible, you may briefly nod to ripples or caustics without lecturing." & vbLf & _
        "- Do NOT list camera settings, filenames, ratings, sequence IDs, or keyword dumps." & vbLf & _
        "- No title, hashtags, or preamble. Output ONLY the caption text."
End Function

Private Function SheetHeaderRow(ws As Worksheet) As Long
    If Not ws Is Nothing Then
        If StrComp(ws.Name, CAPTIONS_SHEET, vbTextCompare) = 0 Then
            SheetHeaderRow = CAPTIONS_HEADER_ROW
            Exit Function
        End If
    End If
    SheetHeaderRow = HEADER_ROW
End Function

' --- Form sheets (Entry / VideoEntry) ---
' Layout supports label/value pairs in columns A/B, C/D, and E/F
' (e.g. Name in A with Camera in C on the same row; Sequence* on one row).

' Write Entry/VideoEntry tags into the media file via exiftool (modifies the original).
' Walks form labels dynamically; Keyword keys use legacy aliases where needed.
Private Sub WriteCatalogTagsToFile(ByVal sheetName As String, ByVal path As String)
    Dim payload As String
    Dim raw As String
    Dim parts() As String
    Dim labels As Variant
    Dim i As Long
    Dim label As String
    Dim keyName As String
    Dim val As String
    Dim blob As String
    Dim anyTag As Boolean

    labels = FieldListFromForm(sheetName)
    blob = ""
    anyTag = False
    For i = LBound(labels) To UBound(labels)
        label = CStr(labels(i))
        If IsWritebackField(label) Then
            keyName = KeywordKeyForLabel(label)
            val = Trim$(CStr(FormValue(sheetName, label)))
            If Len(val) > 0 Then anyTag = True
            If Len(blob) > 0 Then blob = blob & "|"
            blob = blob & keyName & "=" & val
        End If
    Next i

    If Not anyTag Then Exit Sub

    payload = path & vbTab & blob

#If Mac Then
    On Error GoTo appleScriptMissing
    raw = AppleScriptTask("StriderFinderMeta.scpt", "writeCatalogTags", payload)
    On Error GoTo 0
#Else
    MsgBox "Writing tags to media files is implemented for Mac Excel.", vbInformation
    Exit Sub
#End If

    parts = Split(raw, vbTab)
    If UBound(parts) < 0 Then Exit Sub
    If parts(0) = "ERROR" Then
        If UBound(parts) >= 1 Then
            MsgBox "Catalog row saved, but writing tags to the file failed:" & vbCrLf & parts(1), vbExclamation
        Else
            MsgBox "Catalog row saved, but writing tags to the file failed.", vbExclamation
        End If
    End If
    Exit Sub

appleScriptMissing:
    MsgBox "Catalog row saved, but could not run StriderFinderMeta.scpt (writeCatalogTags)." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbExclamation
End Sub

' One-shot: rename Description preview rows and ensure CAPTION line exists (working catalogs).
Public Sub FixDescriptionPreviewRows()
    Dim sheetName As Variant
    For Each sheetName In Array(ENTRY_SHEET, VIDEO_ENTRY_SHEET)
        FixDescriptionPreviewRowsOnSheet CStr(sheetName)
    Next sheetName
    MsgBox "Description preview rows updated (NOTE: / CAPTION:).", vbInformation
End Sub

Private Sub FixDescriptionPreviewRowsOnSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim label As String
    Dim noteRow As Long
    Dim captionRow As Long
    Dim formula As String
    Dim notesRow As Long

    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    noteRow = 0
    captionRow = 0
    For r = 1 To 200
        label = Trim$(CStr(ws.Cells(r, 1).Value))
        If StrComp(label, "Description → notes", vbTextCompare) = 0 Or _
           StrComp(label, "Description → Note", vbTextCompare) = 0 Or _
           StrComp(label, "Description → NOTE:", vbTextCompare) = 0 Then
            noteRow = r
        End If
        If StrComp(label, "Description → CAPTION:", vbTextCompare) = 0 Then
            captionRow = r
        End If
    Next r
    If noteRow = 0 Then Exit Sub

    notesRow = FormRowFor(sheetName, "notes")
    If notesRow = 0 Then notesRow = FormRowFor(sheetName, "Note")
    If notesRow > 0 Then
        formula = "=IF($B$" & notesRow & "="""",""(none)"",""NOTE: ""&$B$" & notesRow & ")"
    Else
        formula = CStr(ws.Cells(noteRow, 2).Formula)
    End If

    ws.Cells(noteRow, 1).Value = "Description → NOTE:"
    If Len(formula) > 0 Then ws.Cells(noteRow, 2).Formula = formula

    If captionRow = 0 Then
        ws.Rows(noteRow + 1).Insert Shift:=xlDown
        captionRow = noteRow + 1
        ws.Cells(captionRow, 1).Value = "Description → CAPTION:"
        ws.Cells(captionRow, 2).Value = "(none)"
    End If
End Sub

' Fill Append preview "Description → CAPTION:" from file (not an Entry edit field).
Private Sub SetDescriptionCaptionPreview(ByVal sheetName As String, ByVal captionBody As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim label As String
    Dim body As String

    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    body = Trim$(captionBody)
    For r = 1 To 200
        label = Trim$(CStr(ws.Cells(r, 1).Value))
        If StrComp(label, "Description → CAPTION:", vbTextCompare) = 0 Then
            If Len(body) = 0 Then
                ws.Cells(r, 2).Value = "(none)"
            Else
                ws.Cells(r, 2).Value = "CAPTION: " & body
            End If
            Exit Sub
        End If
    Next r
End Sub

' Camera / identity fields — not written as Keywords.
' HaveRaw? / VertPix / HorPix ARE written (rawYes|rawNo and VertPix:/HorPix:).
Private Function IsWritebackField(ByVal label As String) As Boolean
    Select Case label
        Case "Name", "Path", "Date", "Time", "Type", "FileSize", _
             "Camera", "Mode", "Lens", "Speed", "f-stop", "ISO", _
             "Focus Distance", "GPS", "Serial Number", _
             "Duration", "FrameRate", "Resolution", _
             "Caption"
            IsWritebackField = False
        Case Else
            IsWritebackField = True
    End Select
End Function

' Legacy keyword spellings so existing files stay consistent; else PascalCase words.
Private Function KeywordKeyForLabel(ByVal label As String) As String
    Select Case label
        Case "Send to UK"
            KeywordKeyForLabel = "SendToUK"
        Case "Caustic camera aim"
            KeywordKeyForLabel = "CausticAim"
        Case "Caustic appearance"
            KeywordKeyForLabel = "CausticAppearance"
        Case "notes", "Note"
            KeywordKeyForLabel = "notes"
        Case "HaveRaw?"
            KeywordKeyForLabel = "HaveRaw"
        Case Else
            KeywordKeyForLabel = PascalCaseKey(label)
    End Select
End Function

' Entry "Note" maps to Photos/Videos header "notes".
Private Function CatalogHeaderName(ByVal fieldLabel As String) As String
    If StrComp(fieldLabel, "Note", vbTextCompare) = 0 Then
        CatalogHeaderName = "notes"
    Else
        CatalogHeaderName = fieldLabel
    End If
End Function

Private Function LabelsMatch(ByVal a As String, ByVal b As String) As Boolean
    If StrComp(a, b, vbTextCompare) = 0 Then
        LabelsMatch = True
        Exit Function
    End If
    If (StrComp(a, "Note", vbTextCompare) = 0 And StrComp(b, "notes", vbTextCompare) = 0) Or _
       (StrComp(a, "notes", vbTextCompare) = 0 And StrComp(b, "Note", vbTextCompare) = 0) Then
        LabelsMatch = True
    End If
End Function

' Real form fields only — skips instructions and Metadata preview rows.
Private Function IsFormFieldLabel(ByVal label As String) As Boolean
    If Len(label) = 0 Then Exit Function
    Select Case label
        Case "Name", "Path", "HaveRaw?", "Date", "Time", "Type", "FileSize", _
             "VertPix", "HorPix", _
             "Camera", "Mode", "Lens", "Speed", "f-stop", "ISO", _
             "Focus Distance", "GPS", "Serial Number", _
             "Duration", "FrameRate", "Resolution", _
             "Sequence", "SequenceID", "SequencePos", _
             "notes", "Note"
            IsFormFieldLabel = True
        Case Else
            IsFormFieldLabel = IsClickPickField(label)
    End Select
End Function

Private Function ColLetter(ByVal col As Long) As String
    Dim n As Long
    Dim rem_ As Long
    Dim s As String
    n = col
    Do While n > 0
        rem_ = (n - 1) Mod 26
        s = Chr$(65 + rem_) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function PascalCaseKey(ByVal label As String) As String
    Dim parts() As String
    Dim i As Long
    Dim w As String
    Dim out As String
    Dim ch As String
    Dim cleaned As String
    Dim j As Long

    parts = Split(Replace(Replace(Trim$(label), "-", " "), "_", " "), " ")
    out = ""
    For i = LBound(parts) To UBound(parts)
        w = Trim$(parts(i))
        If Len(w) = 0 Then GoTo nextWord
        cleaned = ""
        For j = 1 To Len(w)
            ch = Mid$(w, j, 1)
            If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
                cleaned = cleaned & ch
            End If
        Next j
        If Len(cleaned) = 0 Then GoTo nextWord
        out = out & UCase$(Left$(cleaned, 1)) & Mid$(cleaned, 2)
nextWord:
    Next i
    PascalCaseKey = out
End Function

' Field list from Entry / VideoEntry label columns A/C/E (skips instructions).
Private Function FieldListFromForm(ByVal sheetName As String) As Variant
    Dim labels() As String
    Dim n As Long
    Dim r As Long
    Dim c As Long
    Dim label As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim labelCols As Variant
    Dim i As Long
    Dim dup As Boolean
    Dim j As Long

    Set ws = Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    labelCols = Array(1, 3, 5)
    n = 0
    ReDim labels(0 To 0)
    For r = ENTRY_FIRST_FIELD_ROW To lastRow + 5
        For i = LBound(labelCols) To UBound(labelCols)
            c = CLng(labelCols(i))
            label = Trim$(CStr(ws.Cells(r, c).Value))
            If IsFormFieldLabel(label) Then
                dup = False
                For j = 0 To n - 1
                    If LabelsMatch(labels(j), label) Then dup = True: Exit For
                Next j
                If Not dup Then
                    ReDim Preserve labels(0 To n)
                    labels(n) = label
                    n = n + 1
                End If
            End If
        Next i
    Next r
    If n = 0 Then
        FieldListFromForm = Array()
    Else
        FieldListFromForm = labels
    End If
End Function

Private Function PhotoFieldList() As Variant
    PhotoFieldList = FieldListFromForm(ENTRY_SHEET)
End Function

Private Function VideoFieldList() As Variant
    VideoFieldList = FieldListFromForm(VIDEO_ENTRY_SHEET)
End Function

' Find label in cols A/C/E; outCol is the adjacent value column (B/D/F).
Private Function FindFormField(ByVal sheetName As String, ByVal fieldName As String, _
                               ByRef outRow As Long, ByRef outCol As Long) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long
    Dim lastRow As Long
    Dim label As String
    Dim labelCols As Variant
    Dim i As Long

    outRow = 0
    outCol = 0
    FindFormField = False
    Set ws = Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row > lastRow Then _
        lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    labelCols = Array(1, 3, 5)
    For r = ENTRY_FIRST_FIELD_ROW To lastRow + 5
        For i = LBound(labelCols) To UBound(labelCols)
            c = CLng(labelCols(i))
            label = Trim$(CStr(ws.Cells(r, c).Value))
            If LabelsMatch(label, fieldName) Then
                outRow = r
                outCol = c + 1
                FindFormField = True
                Exit Function
            End If
        Next i
    Next r
End Function

Private Function FormValue(ByVal sheetName As String, fieldName As String) As Variant
    Dim r As Long
    Dim c As Long
    If Not FindFormField(sheetName, fieldName, r, c) Then Exit Function
    FormValue = Worksheets(sheetName).Cells(r, c).Value
End Function

Private Sub SetFormValue(ByVal sheetName As String, fieldName As String, value As Variant)
    Dim r As Long
    Dim c As Long
    If Not FindFormField(sheetName, fieldName, r, c) Then Exit Sub
    Worksheets(sheetName).Cells(r, c).Value = value
End Sub

' Photo Entry wrappers (keep existing LoadFromFinderSelection calls working)
Private Function EntryValue(fieldName As String) As Variant
    EntryValue = FormValue(ENTRY_SHEET, fieldName)
End Function

Private Sub SetEntryValue(fieldName As String, value As Variant)
    SetFormValue ENTRY_SHEET, fieldName, value
End Sub

Private Sub SetEntryTextValue(fieldName As String, value As Variant)
    SetFormTextValue ENTRY_SHEET, fieldName, value
End Sub

Private Sub SetFormTextValue(ByVal sheetName As String, fieldName As String, value As Variant)
    Dim r As Long
    Dim c As Long
    If Not FindFormField(sheetName, fieldName, r, c) Then Exit Sub
    With Worksheets(sheetName).Cells(r, c)
        .NumberFormat = "@"
        .Value = CStr(value)
    End With
End Sub

Private Sub FormatEntryField(fieldName As String, ByVal numFmt As String)
    FormatFormField ENTRY_SHEET, fieldName, numFmt
End Sub

Private Sub FormatFormField(ByVal sheetName As String, fieldName As String, ByVal numFmt As String)
    Dim r As Long
    Dim c As Long
    If Not FindFormField(sheetName, fieldName, r, c) Then Exit Sub
    Worksheets(sheetName).Cells(r, c).NumberFormat = numFmt
End Sub

Private Sub FormatCatalogField(ws As Worksheet, ByVal rowNum As Long, headerName As String, ByVal numFmt As String)
    Dim col As Long
    col = HeaderCol(ws, headerName)
    If col = 0 Then Exit Sub
    ws.Cells(rowNum, col).NumberFormat = numFmt
End Sub

Private Function FormRowFor(ByVal sheetName As String, fieldName As String) As Long
    Dim r As Long
    Dim c As Long
    If FindFormField(sheetName, fieldName, r, c) Then
        FormRowFor = r
    Else
        FormRowFor = 0
    End If
End Function

Private Sub ClearFormFields(ByVal sheetName As String, fields As Variant)
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        SetFormValue sheetName, CStr(fields(i)), ""
    Next i
End Sub

Private Sub CopyFormFieldsToRow(ByVal sheetName As String, ws As Worksheet, ByVal destRow As Long, fields As Variant)
    Dim i As Long
    Dim nm As String
    For i = LBound(fields) To UBound(fields)
        nm = CStr(fields(i))
        SetCellByHeader ws, destRow, CatalogHeaderName(nm), FormValue(sheetName, nm)
    Next i
End Sub

Private Sub EnsureColumns(ws As Worksheet, fields As Variant)
    Dim i As Long
    Dim nm As String
    Dim hdrName As String
    Dim hdr As Long
    hdr = SheetHeaderRow(ws)
    For i = LBound(fields) To UBound(fields)
        nm = CStr(fields(i))
        hdrName = CatalogHeaderName(nm)
        If HeaderCol(ws, hdrName) = 0 Then
            ws.Cells(hdr, HeaderLastCol(ws) + 1).Value = hdrName
        End If
    Next i
End Sub

Private Function HeaderCol(ws As Worksheet, headerName As String) As Long
    Dim c As Long
    Dim lastCol As Long
    Dim hdr As Long
    hdr = SheetHeaderRow(ws)
    lastCol = HeaderLastCol(ws)
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(hdr, c).Value)), headerName, vbTextCompare) = 0 Then
            HeaderCol = c
            Exit Function
        End If
    Next c
    HeaderCol = 0
End Function

Private Function HeaderLastCol(ws As Worksheet) As Long
    Dim c As Long
    Dim hdr As Long
    hdr = SheetHeaderRow(ws)
    HeaderLastCol = 1
    For c = 1 To 64
        If Len(Trim$(CStr(ws.Cells(hdr, c).Value))) > 0 Then HeaderLastCol = c
    Next c
End Function

Private Function CatalogLastDataRow(ws As Worksheet) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim bottom As Long
    Dim cellBottom As Long
    Dim hdr As Long

    hdr = SheetHeaderRow(ws)
    lastCol = HeaderLastCol(ws)
    bottom = hdr
    For c = 1 To lastCol
        cellBottom = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If cellBottom > bottom Then bottom = cellBottom
    Next c
    If bottom < hdr Then bottom = hdr
    CatalogLastDataRow = bottom
End Function

Private Function NextDataRow(ws As Worksheet) As Long
    Dim pathCol As Long
    Dim nameCol As Long
    Dim lastRow As Long
    Dim pathVal As String
    Dim nameVal As String
    Dim hdr As Long

    hdr = SheetHeaderRow(ws)
    pathCol = HeaderCol(ws, "Path")
    nameCol = HeaderCol(ws, "Name")
    lastRow = CatalogLastDataRow(ws)

    If lastRow = hdr Then
        NextDataRow = hdr + 1
        Exit Function
    End If

    If lastRow = hdr + 1 Then
        If pathCol > 0 Then pathVal = Trim$(CStr(ws.Cells(lastRow, pathCol).Value))
        If nameCol > 0 Then nameVal = Trim$(CStr(ws.Cells(lastRow, nameCol).Value))
        If Len(pathVal) = 0 And Len(nameVal) = 0 Then
            NextDataRow = lastRow
            Exit Function
        End If
    End If

    NextDataRow = lastRow + 1
End Function

' Match on Name column (filename). Returns 0 if not found.
Private Function FindRowByFileName(ws As Worksheet, ByVal fileName As String) As Long
    Dim nameCol As Long
    Dim pathCol As Long
    Dim r As Long
    Dim lastRow As Long
    Dim cellName As String
    Dim cellPath As String
    Dim want As String

    want = Trim$(fileName)
    If Len(want) = 0 Then
        FindRowByFileName = 0
        Exit Function
    End If

    Dim hdr As Long
    hdr = SheetHeaderRow(ws)
    nameCol = HeaderCol(ws, "Name")
    pathCol = HeaderCol(ws, "Path")
    lastRow = CatalogLastDataRow(ws)
    If lastRow <= hdr Then
        FindRowByFileName = 0
        Exit Function
    End If

    For r = hdr + 1 To lastRow
        cellName = ""
        If nameCol > 0 Then cellName = Trim$(CStr(ws.Cells(r, nameCol).Value))
        If Len(cellName) = 0 And nameCol > 0 Then cellName = BaseName(PathFromCell(ws.Cells(r, nameCol)))
        If StrComp(cellName, want, vbTextCompare) = 0 Then
            FindRowByFileName = r
            Exit Function
        End If
        If pathCol > 0 Then
            cellPath = Trim$(CStr(ws.Cells(r, pathCol).Value))
            If StrComp(BaseName(cellPath), want, vbTextCompare) = 0 Then
                FindRowByFileName = r
                Exit Function
            End If
        End If
    Next r
    FindRowByFileName = 0
End Function

Private Sub SetCellByHeader(ws As Worksheet, ByVal rowNum As Long, headerName As String, value As Variant)
    Dim col As Long
    col = HeaderCol(ws, headerName)
    If col = 0 Then Exit Sub
    ' Keep Sequence* as text (SequencePos like 6/54 must not become a fraction).
    If StrComp(headerName, "Sequence", vbTextCompare) = 0 Or _
       StrComp(headerName, "SequenceID", vbTextCompare) = 0 Or _
       StrComp(headerName, "SequencePos", vbTextCompare) = 0 Then
        ws.Cells(rowNum, col).NumberFormat = "@"
        ws.Cells(rowNum, col).Value = CStr(value)
    Else
        ws.Cells(rowNum, col).Value = value
    End If
End Sub

' Writes Name as plain text (no hyperlink). Open media via double-click or Ctrl+Shift+O.
Private Sub AddNameHyperlink(ws As Worksheet, ByVal rowNum As Long, path As String, fileName As String)
    Dim col As Long
    Dim target As Range

    col = HeaderCol(ws, "Name")
    If col = 0 Then Exit Sub
    Set target = ws.Cells(rowNum, col)
    On Error Resume Next
    If target.Hyperlinks.Count > 0 Then target.Hyperlinks.Delete
    target.Value = fileName
    target.Font.ColorIndex = xlColorIndexAutomatic
    target.Font.Underline = xlUnderlineStyleNone
    target.Locked = False
    On Error GoTo 0
End Sub

' Photos / Videos BeforeDoubleClick: Name → Preview; Caption → edit dialog (Cancel in-cell edit).
Public Sub HandleCatalogNameDoubleClick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim nameCol As Long
    Dim capCol As Long
    Dim path As String
    Dim result As String

    On Error GoTo safeExit
    If Sh Is Nothing Then GoTo safeExit
    If Target Is Nothing Then GoTo safeExit
    If Target.Rows.Count <> 1 Or Target.Columns.Count <> 1 Then GoTo safeExit
    If StrComp(Sh.Name, PHOTOS_SHEET, vbTextCompare) <> 0 And _
       StrComp(Sh.Name, VIDEOS_SHEET, vbTextCompare) <> 0 Then GoTo safeExit
    If Target.Row <= SheetHeaderRow(Sh) Then GoTo safeExit

    capCol = HeaderCol(Sh, "Caption")
    If capCol > 0 And Target.Column = capCol Then
        Cancel = True
        HandleCatalogCaptionClick Sh, Target
        Exit Sub
    End If

    nameCol = HeaderCol(Sh, "Name")
    If nameCol = 0 Or Target.Column <> nameCol Then GoTo safeExit

    path = NormalizePath(PathFromCatalogRow(Sh, Target.Row))
    If Len(path) = 0 Then GoTo safeExit

    Cancel = True
#If Mac Then
    On Error GoTo appleScriptMissing
    result = AppleScriptTask("StriderOpenPreview.scpt", "openInPreview", path)
    On Error GoTo 0
    If Len(result) > 0 Then MsgBox result, vbExclamation
#Else
    ThisWorkbook.FollowHyperlink path
#End If
    Exit Sub

appleScriptMissing:
    MsgBox "Could not run StriderOpenPreview.scpt." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
safeExit:
End Sub

' One-shot / bootstrap: remove Name-column hyperlinks so single-click no longer opens files.
Public Sub StripCatalogNameHyperlinks()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(PHOTOS_SHEET)
    UnprotectSheet ws
    StripNameHyperlinksOnSheet ws
    Set ws = Worksheets(VIDEOS_SHEET)
    UnprotectSheet ws
    StripNameHyperlinksOnSheet ws
    On Error GoTo 0
End Sub

Private Sub StripNameHyperlinksOnSheet(ByVal ws As Worksheet)
    Dim nameCol As Long
    Dim hdr As Long
    Dim lastRow As Long
    Dim r As Long
    Dim cell As Range
    Dim display As String

    If ws Is Nothing Then Exit Sub
    nameCol = HeaderCol(ws, "Name")
    If nameCol = 0 Then Exit Sub
    hdr = SheetHeaderRow(ws)
    lastRow = CatalogLastDataRow(ws)
    On Error Resume Next
    For r = hdr + 1 To lastRow
        Set cell = ws.Cells(r, nameCol)
        If cell.Hyperlinks.Count > 0 Then
            display = cell.Hyperlinks(1).TextToDisplay
            If Len(Trim$(display)) = 0 Then display = CStr(cell.Value)
            cell.Hyperlinks.Delete
            cell.Value = display
            cell.Font.ColorIndex = xlColorIndexAutomatic
            cell.Font.Underline = xlUnderlineStyleNone
        End If
        cell.Locked = False
    Next r
    On Error GoTo 0
End Sub

Private Function CollectSelectedPaths() As Collection
    Dim paths As Collection
    Dim ws As Worksheet
    Dim cell As Range
    Dim path As String
    Dim sheetName As String

    Set paths = New Collection

    If Selection Is Nothing Then
        Set CollectSelectedPaths = paths
        Exit Function
    End If

    sheetName = Selection.Worksheet.Name
    If StrComp(sheetName, PHOTOS_SHEET, vbTextCompare) <> 0 And _
       StrComp(sheetName, VIDEOS_SHEET, vbTextCompare) <> 0 Then
        Set CollectSelectedPaths = paths
        Exit Function
    End If

    Set ws = Worksheets(sheetName)
    For Each cell In Selection.Cells
        If cell.Row = SheetHeaderRow(ws) Then GoTo continueCell
        If cell.Worksheet.Name <> sheetName Then GoTo continueCell
        path = PathFromCatalogRow(ws, cell.Row)
        If Len(path) = 0 Then path = PathFromCell(cell)
        path = NormalizePath(path)
        If Len(path) > 0 Then AddPathUnique paths, path
continueCell:
    Next cell

    Set CollectSelectedPaths = paths
End Function

Private Function PathFromCatalogRow(ws As Worksheet, ByVal rowNum As Long) As String
    Dim pathCol As Long
    Dim nameCol As Long

    pathCol = HeaderCol(ws, "Path")
    nameCol = HeaderCol(ws, "Name")

    If pathCol > 0 Then
        If Len(Trim$(CStr(ws.Cells(rowNum, pathCol).Value))) > 0 Then
            PathFromCatalogRow = CStr(ws.Cells(rowNum, pathCol).Value)
            Exit Function
        End If
    End If

    If nameCol > 0 Then PathFromCatalogRow = PathFromCell(ws.Cells(rowNum, nameCol))
End Function

Private Sub AddPathUnique(paths As Collection, ByVal path As String)
    Dim i As Long
    For i = 1 To paths.Count
        If StrComp(CStr(paths(i)), path, vbTextCompare) = 0 Then Exit Sub
    Next i
    paths.Add path
End Sub

Private Function PathFromCell(cell As Range) As String
    On Error Resume Next
    If cell.Hyperlinks.Count > 0 Then
        PathFromCell = cell.Hyperlinks(1).Address
    Else
        PathFromCell = CStr(cell.Value)
    End If
    On Error GoTo 0
End Function

Private Function NormalizePath(ByVal path As String) As String
    path = Trim$(path)
    If Len(path) = 0 Then Exit Function
    If LCase$(Left$(path, 7)) = "file://" Then
        path = Mid$(path, 8)
        If Left$(path, 1) <> "/" And InStr(path, ":") = 0 Then path = "/" & path
        path = Replace(path, "%20", " ")
    End If
    NormalizePath = path
End Function

Private Function BaseName(ByVal path As String) As String
    Dim p As Long
    path = Replace(path, "\", "/")
    p = InStrRev(path, "/")
    If p > 0 Then BaseName = Mid$(path, p + 1) Else BaseName = path
End Function

Private Function JoinPaths(paths As Collection) As String
    Dim i As Long
    Dim out As String
    For i = 1 To paths.Count
        If i > 1 Then out = out & Chr(0)
        out = out & CStr(paths(i))
    Next i
    JoinPaths = out
End Function

Private Function MetaAt(meta() As String, ByVal idx As Long) As String
    If idx >= LBound(meta) And idx <= UBound(meta) Then MetaAt = Trim$(meta(idx))
End Function

Private Function FirstNonEmpty(ParamArray vals() As Variant) As String
    Dim i As Long
    Dim s As String
    For i = LBound(vals) To UBound(vals)
        s = Trim$(CStr(vals(i)))
        If Len(s) > 0 And s <> "-" Then
            FirstNonEmpty = s
            Exit Function
        End If
    Next i
End Function

Private Function ExtractDate(ByVal stamp As String) As String
    Dim s As String
    s = Trim$(stamp)
    If Len(s) = 0 Then Exit Function
    s = Replace(s, "-", ":")
    If Len(s) >= 10 Then ExtractDate = Left$(s, 4) & "-" & Mid$(s, 6, 2) & "-" & Mid$(s, 9, 2)
End Function

Private Function ExtractTime(ByVal stamp As String) As String
    Dim s As String
    Dim p As Long
    s = Trim$(stamp)
    If Len(s) = 0 Then Exit Function
    p = InStr(s, " ")
    If p > 0 Then ExtractTime = Mid$(s, p + 1)
End Function
