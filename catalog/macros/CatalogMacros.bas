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
' Re-import this module, then run AssignMacroShortcuts (or reopen the workbook).
' If click-to-pick stops working, run FixEnableEvents, then re-paste Entry/VideoEntry SelectionChange.

Private Const ENTRY_SHEET As String = "Entry"
Private Const VIDEO_ENTRY_SHEET As String = "VideoEntry"
Private Const PHOTOS_SHEET As String = "Photos"
Private Const VIDEOS_SHEET As String = "Videos"
Private Const HEADER_ROW As Long = 1
Private Const ENTRY_FIRST_FIELD_ROW As Long = 9
Private Const ENTRY_OPTION_FIRST_COL As Long = 3

Public Sub Auto_Open()
    Application.EnableEvents = True
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

    Application.OnKey "^+q", "LoadFromFinderSelection"
    Application.OnKey "^+i", "LoadFromPreviewSelection"
    Application.OnKey "^+w", "AppendPhotoRow"
    Application.OnKey "^+e", "LoadVideoFromFinderSelection"
    Application.OnKey "^+u", "LoadVideoFromPreviewSelection"
    Application.OnKey "^+r", "AppendVideoRow"
    Application.OnKey "^+o", "OpenSelectedInPreview"
    Application.OnKey "^+j", "RevealSelectedInFinder"
    Application.OnKey "^+d", "TrashFrontPreviewDocument"
    On Error GoTo 0
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

Private Function IsClickPickField(ByVal label As String) As Boolean
    Select Case label
        Case "Rating", "Species", "Behavior", "Locale", "Send to UK", "View", "Lighting"
            IsClickPickField = True
        Case Else
            IsClickPickField = False
    End Select
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

    SetEntryValue "Path", path
    SetEntryValue "Name", fileName
    SetEntryValue "Type", UCase$(ext)
    SetEntryValue "FileSize", fileSize
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

    Worksheets(ENTRY_SHEET).Activate
    MsgBox "Loaded metadata for:" & vbCrLf & fileName & vbCrLf & vbCrLf & _
           "Set tags, then run AppendPhotoRow.", vbInformation
    Exit Sub

appleScriptMissing:
    MsgBox "Could not run StriderFinderMeta.scpt (" & appleHandler & ")." & vbCrLf & _
           "Run scripts/install_mac_helpers.sh, then restart Excel.", vbCritical
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

' --- Form sheets (Entry / VideoEntry): A=label, B=value ---

' Write Entry/VideoEntry tags into the media file via exiftool (modifies the original).
Private Sub WriteCatalogTagsToFile(ByVal sheetName As String, ByVal path As String)
    Dim payload As String
    Dim raw As String
    Dim parts() As String
    Dim rating As String
    Dim species As String
    Dim behavior As String
    Dim locale As String
    Dim sendUK As String
    Dim viewTag As String
    Dim lighting As String
    Dim notes As String

    rating = Trim$(CStr(FormValue(sheetName, "Rating")))
    species = Trim$(CStr(FormValue(sheetName, "Species")))
    behavior = Trim$(CStr(FormValue(sheetName, "Behavior")))
    locale = Trim$(CStr(FormValue(sheetName, "Locale")))
    sendUK = Trim$(CStr(FormValue(sheetName, "Send to UK")))
    viewTag = Trim$(CStr(FormValue(sheetName, "View")))
    lighting = Trim$(CStr(FormValue(sheetName, "Lighting")))
    notes = Trim$(CStr(FormValue(sheetName, "notes")))

    If Len(rating) = 0 And Len(species) = 0 And Len(behavior) = 0 And Len(locale) = 0 And _
       Len(sendUK) = 0 And Len(viewTag) = 0 And Len(lighting) = 0 And Len(notes) = 0 Then
        Exit Sub
    End If

    payload = path & vbTab & rating & "|" & species & "|" & behavior & "|" & locale & "|" & _
              sendUK & "|" & viewTag & "|" & lighting & "|" & notes

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

Private Function PhotoFieldList() As Variant
    PhotoFieldList = Array( _
        "Name", "Path", "Date", "Time", "Type", "FileSize", "Camera", "Mode", "Lens", _
        "Speed", "f-stop", "ISO", "Rating", "Species", "Behavior", "Locale", _
        "Send to UK", "View", "Lighting", "notes")
End Function

Private Function VideoFieldList() As Variant
    VideoFieldList = Array( _
        "Name", "Path", "Date", "Time", "Type", "FileSize", "Duration", "FrameRate", _
        "Resolution", "Camera", "Speed", "f-stop", "ISO", _
        "Rating", "Species", "Behavior", "Locale", "Send to UK", "View", "Lighting", "notes")
End Function

Private Function FormValue(ByVal sheetName As String, fieldName As String) As Variant
    Dim r As Long
    r = FormRowFor(sheetName, fieldName)
    If r = 0 Then Exit Function
    FormValue = Worksheets(sheetName).Cells(r, 2).Value
End Function

Private Sub SetFormValue(ByVal sheetName As String, fieldName As String, value As Variant)
    Dim r As Long
    r = FormRowFor(sheetName, fieldName)
    If r = 0 Then Exit Sub
    Worksheets(sheetName).Cells(r, 2).Value = value
End Sub

' Photo Entry wrappers (keep existing LoadFromFinderSelection calls working)
Private Function EntryValue(fieldName As String) As Variant
    EntryValue = FormValue(ENTRY_SHEET, fieldName)
End Function

Private Sub SetEntryValue(fieldName As String, value As Variant)
    SetFormValue ENTRY_SHEET, fieldName, value
End Sub

Private Sub FormatEntryField(fieldName As String, ByVal numFmt As String)
    FormatFormField ENTRY_SHEET, fieldName, numFmt
End Sub

Private Sub FormatFormField(ByVal sheetName As String, fieldName As String, ByVal numFmt As String)
    Dim r As Long
    r = FormRowFor(sheetName, fieldName)
    If r = 0 Then Exit Sub
    Worksheets(sheetName).Cells(r, 2).NumberFormat = numFmt
End Sub

Private Sub FormatCatalogField(ws As Worksheet, ByVal rowNum As Long, headerName As String, ByVal numFmt As String)
    Dim col As Long
    col = HeaderCol(ws, headerName)
    If col = 0 Then Exit Sub
    ws.Cells(rowNum, col).NumberFormat = numFmt
End Sub

Private Function FormRowFor(ByVal sheetName As String, fieldName As String) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim label As String

    Set ws = Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = ENTRY_FIRST_FIELD_ROW To lastRow
        label = Trim$(CStr(ws.Cells(r, 1).Value))
        If StrComp(label, fieldName, vbTextCompare) = 0 Then
            FormRowFor = r
            Exit Function
        End If
    Next r
    FormRowFor = 0
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
        SetCellByHeader ws, destRow, nm, FormValue(sheetName, nm)
    Next i
End Sub

Private Sub EnsureColumns(ws As Worksheet, fields As Variant)
    Dim i As Long
    Dim nm As String
    For i = LBound(fields) To UBound(fields)
        nm = CStr(fields(i))
        If HeaderCol(ws, nm) = 0 Then
            ws.Cells(HEADER_ROW, HeaderLastCol(ws) + 1).Value = nm
        End If
    Next i
End Sub

Private Function HeaderCol(ws As Worksheet, headerName As String) As Long
    Dim c As Long
    Dim lastCol As Long
    lastCol = HeaderLastCol(ws)
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(HEADER_ROW, c).Value)), headerName, vbTextCompare) = 0 Then
            HeaderCol = c
            Exit Function
        End If
    Next c
    HeaderCol = 0
End Function

Private Function HeaderLastCol(ws As Worksheet) As Long
    Dim c As Long
    HeaderLastCol = 1
    For c = 1 To 64
        If Len(Trim$(CStr(ws.Cells(HEADER_ROW, c).Value))) > 0 Then HeaderLastCol = c
    Next c
End Function

Private Function CatalogLastDataRow(ws As Worksheet) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim bottom As Long
    Dim cellBottom As Long

    lastCol = HeaderLastCol(ws)
    bottom = HEADER_ROW
    For c = 1 To lastCol
        cellBottom = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If cellBottom > bottom Then bottom = cellBottom
    Next c
    If bottom < HEADER_ROW Then bottom = HEADER_ROW
    CatalogLastDataRow = bottom
End Function

Private Function NextDataRow(ws As Worksheet) As Long
    Dim pathCol As Long
    Dim nameCol As Long
    Dim lastRow As Long
    Dim pathVal As String
    Dim nameVal As String

    pathCol = HeaderCol(ws, "Path")
    nameCol = HeaderCol(ws, "Name")
    lastRow = CatalogLastDataRow(ws)

    If lastRow = HEADER_ROW Then
        NextDataRow = HEADER_ROW + 1
        Exit Function
    End If

    If lastRow = HEADER_ROW + 1 Then
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

    nameCol = HeaderCol(ws, "Name")
    pathCol = HeaderCol(ws, "Path")
    lastRow = CatalogLastDataRow(ws)
    If lastRow <= HEADER_ROW Then
        FindRowByFileName = 0
        Exit Function
    End If

    For r = HEADER_ROW + 1 To lastRow
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
    ws.Cells(rowNum, col).Value = value
End Sub

Private Sub AddNameHyperlink(ws As Worksheet, ByVal rowNum As Long, path As String, fileName As String)
    Dim col As Long
    Dim target As Range

    col = HeaderCol(ws, "Name")
    If col = 0 Then Exit Sub
    Set target = ws.Cells(rowNum, col)
    If target.Hyperlinks.Count > 0 Then target.Hyperlinks.Delete
    target.Hyperlinks.Add Anchor:=target, Address:=path, TextToDisplay:=fileName
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
        If cell.Row = HEADER_ROW Then GoTo continueCell
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
