-- Excel AppleScriptTask helpers: Finder selection + exiftool metadata
-- Photo return: OK <tab> path <tab> Name|ext|mime|dto|create|size|shutter|aperture|iso|camera|mode|lens
-- Video return: OK <tab> path <tab> Name|ext|mime|create|mediaCreate|size|duration|fps|width|height|deviceModel|model|shutter|aperture|iso

on loadFromFinder(dummy)
    set p to my selectedPOSIXPath()
    if p is "" then return "ERROR" & tab & "Select a photo or video in Finder first."
    return my metadataLine(p)
end loadFromFinder

on loadFromPreview(dummy)
    set p to my frontPreviewPOSIXPath("")
    if p is "" then return "ERROR" & tab & "Open a saved photo in Preview first (Untitled / unsaved windows have no path)."
    return my metadataLine(p)
end loadFromPreview

on loadVideoFromFinder(dummy)
    set p to my selectedPOSIXPath()
    if p is "" then return "ERROR" & tab & "Select a video in Finder first."
    return my videoMetadataLine(p)
end loadVideoFromFinder

on loadVideoFromPreview(dummy)
    set p to my frontPreviewPOSIXPath("")
    if p is "" then return "ERROR" & tab & "Open a saved video in Preview first (Untitled / unsaved windows have no path)."
    return my videoMetadataLine(p)
end loadVideoFromPreview

on readPath(posixPath)
    if posixPath is "" then return "ERROR" & tab & "No path provided."
    return my metadataLine(posixPath)
end readPath

on frontPreviewPOSIXPath(dummy)
    -- Callers (AppleScriptTask and internal) must pass one arg; use "".
    -- Prefer "tell … to set …" so Preview tell-scoping cannot swallow locals.
    try
        tell application "Preview" to set docCount to count of documents
        if docCount is 0 then return ""
    on error
        return ""
    end try

    set pathTry to my pathOfFrontPreviewDocument()
    if pathTry is not "" then return pathTry

    try
        tell application "Preview" to set pathTry to (get path of document of front window)
        if pathTry is missing value then return ""
        set pathText to pathTry as text
        if pathText starts with "/" then return pathText
    end try
    return ""
end frontPreviewPOSIXPath

on pathOfFrontPreviewDocument()
    try
        tell application "Preview" to set pathTry to (get path of front document)
        if pathTry is missing value then return ""
        set pathText to pathTry as text
        if pathText starts with "/" then return pathText
        return ""
    on error
        return ""
    end try
end pathOfFrontPreviewDocument

-- Trash the front/selected image using Preview's own Edit menu
-- (Edit → Move Selected Image to Trash). Fast; no Finder hop.
-- payload: ignored (kept for AppleScriptTask). Returns OK <tab> path.
on trashFrontPreview(payload)
    set p to my frontPreviewPOSIXPath("")
    if p is "" then return "ERROR" & tab & "No front document in Preview."

    try
        tell application "Preview" to activate
        delay 0.05
        tell application "System Events"
            tell process "Preview"
                set frontmost to true
                click menu item "Move Selected Image to Trash" of menu "Edit" of menu bar 1
            end tell
        end tell
        -- Dismiss Preview's confirmation if it appears (Move to Trash / Delete)
        delay 0.15
        try
            tell application "System Events"
                tell process "Preview"
                    if exists (sheet 1 of window 1) then
                        try
                            click button "Move to Trash" of sheet 1 of window 1
                        on error
                            try
                                click button "Delete" of sheet 1 of window 1
                            on error
                                keystroke return
                            end try
                        end try
                    end if
                end tell
            end tell
        end try
        return "OK" & tab & p
    on error errMsg
        return "ERROR" & tab & errMsg
    end try
end trashFrontPreview

-- payload: path <tab> Rating|Species|Behavior|Locale|SendToUK|View|Lighting|notes|Sequence|SequenceID|SequencePos
on writeCatalogTags(payload)
    if payload is "" then return "ERROR" & tab & "No tag payload."
    set AppleScript's text item delimiters to tab
    set parts to text items of payload
    set AppleScript's text item delimiters to ""
    if (count of parts) < 2 then return "ERROR" & tab & "Tag payload incomplete."

    set posixPath to item 1 of parts
    set tagBlob to item 2 of parts
    return my writeTagsToFile(posixPath, tagBlob)
end writeCatalogTags

on writeTagsToFile(posixPath, tagBlob)
    set exif to my exiftoolPath()
    if exif is "" then return "ERROR" & tab & "exiftool not found. Install with Homebrew: brew install exiftool"

    try
        do shell script "test -e " & quoted form of posixPath
    on error
        return "ERROR" & tab & "File not found: " & posixPath
    end try

    set AppleScript's text item delimiters to "|"
    set tags to text items of tagBlob
    set AppleScript's text item delimiters to ""

    set rating to my tagAt(tags, 1)
    set species to my tagAt(tags, 2)
    set behavior to my tagAt(tags, 3)
    set locale to my tagAt(tags, 4)
    set sendUK to my tagAt(tags, 5)
    set viewTag to my tagAt(tags, 6)
    set lighting to my tagAt(tags, 7)
    set notes to my tagAt(tags, 8)
    set sequence to my tagAt(tags, 9)
    set sequenceID to my tagAt(tags, 10)
    set sequencePos to my tagAt(tags, 11)

    -- If Entry left Sequence* blank, keep any existing Sequence* keywords on the file
    if sequence is "" and sequenceID is "" and sequencePos is "" then
        set existingSeq to my readSequenceFields(posixPath)
        set sequence to item 1 of existingSeq
        set sequenceID to item 2 of existingSeq
        set sequencePos to item 3 of existingSeq
    end if

    set kwList to {}
    if species is not "" then set end of kwList to "Species:" & species
    if behavior is not "" then set end of kwList to "Behavior:" & behavior
    if locale is not "" then set end of kwList to "Locale:" & locale
    if sendUK is not "" then set end of kwList to "SendToUK:" & sendUK
    if viewTag is not "" then set end of kwList to "View:" & viewTag
    if lighting is not "" then set end of kwList to "Lighting:" & lighting
    if sequence is not "" then set end of kwList to "Sequence:" & sequence
    if sequenceID is not "" then set end of kwList to "SequenceID:" & sequenceID
    if sequencePos is not "" then set end of kwList to "SequencePos:" & sequencePos

    set AppleScript's text item delimiters to ", "
    set kwJoined to kwList as text
    set AppleScript's text item delimiters to ""

    -- Build exiftool args; overwrite in place (no _original backup clutter)
    set cmd to quoted form of exif & " -m -overwrite_original -api LargeFileSupport=1"

    if rating is not "" then
        set cmd to cmd & " -Rating=" & quoted form of rating & " -XMP:Rating=" & quoted form of rating
    end if
    if notes is not "" then
        set cmd to cmd & " -XMP-dc:Description=" & quoted form of notes
        set cmd to cmd & " -EXIF:ImageDescription=" & quoted form of notes
        set cmd to cmd & " -IPTC:Caption-Abstract=" & quoted form of notes
    end if
    if locale is not "" then
        set cmd to cmd & " -XMP-iptcCore:Location=" & quoted form of locale
        set cmd to cmd & " -IPTC:Sub-location=" & quoted form of locale
    end if
    if kwJoined is not "" then
        set cmd to cmd & " -sep \", \" -Keywords=" & quoted form of kwJoined
        set cmd to cmd & " -XMP-dc:Subject=" & quoted form of kwJoined
    end if

    set cmd to cmd & " " & quoted form of posixPath

    try
        do shell script cmd
        return "OK" & tab & posixPath
    on error errMsg
        return "ERROR" & tab & errMsg
    end try
end writeTagsToFile

on tagAt(tags, idx)
    if (count of tags) < idx then return ""
    set t to item idx of tags
    try
        return trimText(t)
    on error
        return t as text
    end try
end tagAt

on trimText(s)
    set s to s as text
    repeat while s starts with " " or s starts with tab
        if (length of s) is 0 then exit repeat
        set s to text 2 thru -1 of s
    end repeat
    repeat while s ends with " " or s ends with tab
        if (length of s) is 0 then exit repeat
        set s to text 1 thru -2 of s
    end repeat
    return s
end trimText

on selectedPOSIXPath()
    tell application "Finder"
        set sel to selection as list
        if (count of sel) is 0 then return ""
        set f to item 1 of sel
        try
            return POSIX path of (f as alias)
        on error
            return ""
        end try
    end tell
end selectedPOSIXPath

on exiftoolPath()
    set exif to "/usr/local/bin/exiftool"
    try
        do shell script "test -x " & quoted form of exif
        return exif
    on error
        set exif to "/opt/homebrew/bin/exiftool"
    end try
    try
        do shell script "test -x " & quoted form of exif
        return exif
    on error
        return ""
    end try
end exiftoolPath

on metadataLine(posixPath)
    set exif to my exiftoolPath()
    if exif is "" then return "ERROR" & tab & "exiftool not found. Install with Homebrew: brew install exiftool"

    try
        do shell script "test -e " & quoted form of posixPath
    on error
        return "ERROR" & tab & "File not found: " & posixPath
    end try

    set fmt to "$FileName|$FileTypeExtension|$MIMEType|$DateTimeOriginal|$CreateDate|$FileSize|$ExposureTime|$FNumber|$ISO|$Model|$ExposureProgram|$LensID"
    set cmd to quoted form of exif & " -m -s3 -p " & quoted form of fmt & " " & quoted form of posixPath
    try
        set meta to do shell script cmd
    on error errMsg
        set fmt2 to "$FileName|$FileTypeExtension|$MIMEType|$DateTimeOriginal|$CreateDate|$FileSize|$ExposureTime|$FNumber|$ISO|$Model|$ExposureProgram|$LensModel"
        set cmd2 to quoted form of exif & " -m -s3 -p " & quoted form of fmt2 & " " & quoted form of posixPath
        try
            set meta to do shell script cmd2
        on error errMsg2
            return "ERROR" & tab & errMsg2
        end try
    end try

    set seqFields to my readSequenceFields(posixPath)
    set meta to meta & "|" & (item 1 of seqFields) & "|" & (item 2 of seqFields) & "|" & (item 3 of seqFields)
    return "OK" & tab & posixPath & tab & meta
end metadataLine

-- Returns {Sequence, SequenceID, SequencePos} from Keywords/Subject
on readSequenceFields(posixPath)
    set exif to my exiftoolPath()
    set seq to ""
    set seqID to ""
    set seqPos to ""
    try
        set kwRaw to do shell script quoted form of exif & " -m -s3 -Keywords -Subject " & quoted form of posixPath
    on error
        return {seq, seqID, seqPos}
    end try
    set AppleScript's text item delimiters to {return, linefeed, ","}
    set bits to text items of kwRaw
    set AppleScript's text item delimiters to ""
    repeat with bit in bits
        set t to my trimText(bit as text)
        -- Check longer keys first (SequenceID/Pos also start with "Sequence")
        if t starts with "SequenceID:" then
            set seqID to text 12 thru -1 of t
        else if t starts with "SequencePos:" then
            set seqPos to text 13 thru -1 of t
        else if t starts with "Sequence:" then
            set seq to text 10 thru -1 of t
        end if
    end repeat
    return {seq, seqID, seqPos}
end readSequenceFields

on videoMetadataLine(posixPath)
    set exif to my exiftoolPath()
    if exif is "" then return "ERROR" & tab & "exiftool not found. Install with Homebrew: brew install exiftool"

    try
        do shell script "test -e " & quoted form of posixPath
    on error
        return "ERROR" & tab & "File not found: " & posixPath
    end try

    -- Container-level tags
    set fmt to "$FileName|$FileTypeExtension|$MIMEType|$CreateDate|$MediaCreateDate|$FileSize|$Duration|$VideoFrameRate|$ImageWidth|$ImageHeight|$DeviceModelName|$Model"
    set cmd to quoted form of exif & " -m -s3 -p " & quoted form of fmt & " " & quoted form of posixPath
    try
        set meta to do shell script cmd
    on error errMsg
        return "ERROR" & tab & errMsg
    end try

    -- First embedded exposure values (Sony timed metadata)
    set shutter to ""
    set aperture to ""
    set isoVal to ""
    try
        set shutter to do shell script quoted form of exif & " -m -ee -s3 -ExposureTime " & quoted form of posixPath & " 2>/dev/null | head -1"
    end try
    try
        set aperture to do shell script quoted form of exif & " -m -ee -s3 -FNumber " & quoted form of posixPath & " 2>/dev/null | head -1"
    end try
    try
        set isoVal to do shell script quoted form of exif & " -m -ee -s3 -ISO " & quoted form of posixPath & " 2>/dev/null | head -1"
    end try

    return "OK" & tab & posixPath & tab & meta & "|" & shutter & "|" & aperture & "|" & isoVal
end videoMetadataLine
