-- Excel AppleScriptTask helpers: Finder selection + exiftool metadata
-- Photo return: OK <tab> path <tab> Name|ext|mime|dto|create|size|shutter|aperture|iso|camera|mode|lens|focus|gps|serial|vertPix|horPix|haveRaw|seq|seqID|seqPos
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

-- payload: fileOrFolderPath <tab> SequenceID
-- Returns OK <tab> path1 <tab> pos1 <tab> path2 <tab> pos2 …
-- (via StriderSequence.sh → list_sequence_members.py)
on listSequenceMembers(payload)
    if payload is "" then return "ERROR" & tab & "No sequence payload."
    set AppleScript's text item delimiters to tab
    set parts to text items of payload
    set AppleScript's text item delimiters to ""
    if (count of parts) < 2 then return "ERROR" & tab & "Need path and SequenceID."

    set anchor to item 1 of parts
    set seqID to item 2 of parts
    if anchor is "" or seqID is "" then return "ERROR" & tab & "Need path and SequenceID."

    set sh to (POSIX path of (path to home folder)) & "Library/Application Scripts/com.microsoft.Excel/StriderSequence.sh"
    try
        do shell script "test -x " & quoted form of sh
    on error
        return "ERROR" & tab & "StriderSequence.sh missing. Run scripts/install_mac_helpers.sh, then restart Excel."
    end try

    try
        set raw to do shell script quoted form of sh & " " & quoted form of anchor & " " & quoted form of seqID
    on error errMsg
        return "ERROR" & tab & errMsg
    end try

    set AppleScript's text item delimiters to {return, linefeed}
    set lines to text items of raw
    set AppleScript's text item delimiters to ""
    if (count of lines) < 1 then return "ERROR" & tab & "Empty sequence list response."
    set head to item 1 of lines
    if head starts with "ERROR" then return head
    if head is not "OK" then return "ERROR" & tab & "Unexpected sequence list header: " & head

    set out to "OK"
    repeat with i from 2 to count of lines
        set ln to item i of lines
        if ln is "" then
            -- skip
        else
            set AppleScript's text item delimiters to tab
            set bits to text items of ln
            set AppleScript's text item delimiters to ""
            if (count of bits) >= 1 then
                set out to out & tab & (item 1 of bits)
                if (count of bits) >= 2 then
                    set out to out & tab & (item 2 of bits)
                else
                    set out to out & tab & ""
                end if
            end if
        end if
    end repeat
    return out
end listSequenceMembers

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

-- payload: path <tab> key=value|key=value|...
-- Labeled fields are order-independent. New catalog tags need no AppleScript edits:
-- every non-empty key except Rating/notes becomes Keywords "Key:value".
-- Specials: Rating → MWG:Rating; notes → Description NOTE: section (+ notes: Keyword);
-- Locale → Keywords + MWG:Location; blank Sequence* preserves file Sequence* keywords.
-- MWG:Description holds optional sections: "NOTE: …" and/or "CAPTION: …" (legacy plain = NOTE).
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

-- payload: posixPath <tab> captionText — merges CAPTION: section; preserves NOTE:.
on writeCaption(payload)
    if payload is "" then return "ERROR" & tab & "No caption payload."
    set AppleScript's text item delimiters to tab
    set parts to text items of payload
    set AppleScript's text item delimiters to ""
    if (count of parts) < 2 then return "ERROR" & tab & "Caption payload incomplete."
    set posixPath to item 1 of parts
    set captionText to item 2 of parts
    if (count of parts) > 2 then
        set AppleScript's text item delimiters to tab
        set captionText to (items 2 thru -1 of parts) as text
        set AppleScript's text item delimiters to ""
    end if
    return my mergeDescriptionSection(posixPath, "caption", captionText)
end writeCaption

on writeTagsToFile(posixPath, tagBlob)
    set exif to my exiftoolPath()
    if exif is "" then return "ERROR" & tab & "exiftool not found. Install with Homebrew: brew install exiftool"

    try
        do shell script "test -e " & quoted form of posixPath
    on error
        return "ERROR" & tab & "File not found: " & posixPath
    end try

    set rating to my fieldValue(tagBlob, "Rating")
    set notes to my fieldValue(tagBlob, "notes")
    set locale to my fieldValue(tagBlob, "Locale")
    set sequence to my fieldValue(tagBlob, "Sequence")
    set sequenceID to my fieldValue(tagBlob, "SequenceID")
    set sequencePos to my fieldValue(tagBlob, "SequencePos")

    -- If Entry left Sequence* blank, keep any existing Sequence* keywords on the file
    set preservedSeq to false
    if sequence is "" and sequenceID is "" and sequencePos is "" then
        set existingSeq to my readSequenceFields(posixPath)
        set sequence to item 1 of existingSeq
        set sequenceID to item 2 of existingSeq
        set sequencePos to item 3 of existingSeq
        set preservedSeq to true
    end if

    set kwList to {}
    set AppleScript's text item delimiters to "|"
    set bits to text items of tagBlob
    set AppleScript's text item delimiters to ""
    repeat with bit in bits
        set t to my trimText(bit as text)
        if t is "" then
            -- skip
        else
            set eq to offset of "=" in t
            if eq > 1 then
                set k to text 1 thru (eq - 1) of t
                if eq = (length of t) then
                    set v to ""
                else
                    set v to text (eq + 1) thru -1 of t
                end if
                -- Rating is not a Keyword; notes also go into Description NOTE: section
                -- HaveRaw Y/N → bare keywords rawYes / rawNo (not HaveRaw:Y)
                if k is "Rating" then
                    -- skip
                else if k is "HaveRaw" or k is "HaveRaw?" then
                    set hv to my upperText(v)
                    if hv is "Y" or hv is "YES" then
                        set end of kwList to "rawYes"
                    else if hv is "N" or hv is "NO" then
                        set end of kwList to "rawNo"
                    end if
                else if k is "notes" then
                    if v is not "" then set end of kwList to "notes:" & v
                else if k is "Sequence" or k is "SequenceID" or k is "SequencePos" then
                    if v is not "" then set end of kwList to k & ":" & v
                else if v is not "" then
                    set end of kwList to k & ":" & v
                end if
            end if
        end if
    end repeat

    if preservedSeq then
        if sequence is not "" then set end of kwList to "Sequence:" & sequence
        if sequenceID is not "" then set end of kwList to "SequenceID:" & sequenceID
        if sequencePos is not "" then set end of kwList to "SequencePos:" & sequencePos
    end if

    -- Use || as list sep so commas inside tag values stay intact
    set AppleScript's text item delimiters to "||"
    set kwJoined to kwList as text
    set AppleScript's text item delimiters to ""

    -- Merge NOTE: into Description; preserve CAPTION:
    set curDesc to my readDescriptionRaw(posixPath)
    set sections to my parseDescriptionSections(curDesc)
    set keepCaption to item 2 of sections
    set newDesc to my formatDescription(notes, keepCaption)

    -- MWG composites sync overlapping EXIF/IPTC/XMP in one write.
    -- CodedCharacterSet creates an IPTC block when missing so Keywords land in IPTC too.
    -- -P keeps Finder Date Modified / FileCreateDate from jumping to "now"
    set cmd to quoted form of exif & " -m -P -overwrite_original -api LargeFileSupport=1"
    set cmd to cmd & " -use MWG -charset iptc=UTF8 -IPTC:CodedCharacterSet=UTF8"

    if rating is not "" then
        set cmd to cmd & " -MWG:Rating=" & quoted form of rating
    end if
    set cmd to cmd & " -MWG:Description=" & quoted form of newDesc
    if locale is not "" then
        set cmd to cmd & " -MWG:Location=" & quoted form of locale
    end if
    -- Replace Keywords/Subject with the catalog set (clears prior swapped junk)
    set cmd to cmd & " -sep \"||\" -MWG:Keywords=" & quoted form of kwJoined

    set cmd to cmd & " " & quoted form of posixPath

    try
        do shell script cmd
        return "OK" & tab & posixPath
    on error errMsg
        return "ERROR" & tab & errMsg
    end try
end writeTagsToFile

-- sectionKind: "note" or "caption"; newText is the body (no prefix).
on mergeDescriptionSection(posixPath, sectionKind, newText)
    set exif to my exiftoolPath()
    if exif is "" then return "ERROR" & tab & "exiftool not found. Install with Homebrew: brew install exiftool"
    try
        do shell script "test -e " & quoted form of posixPath
    on error
        return "ERROR" & tab & "File not found: " & posixPath
    end try

    set curDesc to my readDescriptionRaw(posixPath)
    set sections to my parseDescriptionSections(curDesc)
    set noteBody to item 1 of sections
    set captionBody to item 2 of sections
    if sectionKind is "caption" then
        set captionBody to my trimText(newText)
    else
        set noteBody to my trimText(newText)
    end if
    set newDesc to my formatDescription(noteBody, captionBody)

    set cmd to quoted form of exif & " -m -P -overwrite_original -api LargeFileSupport=1"
    set cmd to cmd & " -use MWG -charset iptc=UTF8 -IPTC:CodedCharacterSet=UTF8"
    set cmd to cmd & " -MWG:Description=" & quoted form of newDesc
    set cmd to cmd & " " & quoted form of posixPath
    try
        do shell script cmd
        return "OK" & tab & posixPath
    on error errMsg
        return "ERROR" & tab & errMsg
    end try
end mergeDescriptionSection

on readDescriptionRaw(posixPath)
    set exif to my exiftoolPath()
    try
        -- Use -b (not -s3): -s3 replaces newlines with "." and would corrupt NOTE/CAPTION splits.
        set raw to do shell script quoted form of exif & " -m -b -MWG:Description " & quoted form of posixPath
        return my trimText(raw)
    on error
        return ""
    end try
end readDescriptionRaw

-- Returns {noteBody, captionBody}. Legacy unprefixed text → noteBody.
on parseDescriptionSections(raw)
    set t to my trimText(raw)
    set noteBody to ""
    set captionBody to ""
    if t is "" then return {noteBody, captionBody}

    set noteMark to offset of "NOTE:" in t
    set capMark to offset of "CAPTION:" in t
    if noteMark is 0 and capMark is 0 then
        return {t, ""}
    end if

    if noteMark > 0 and (capMark is 0 or noteMark < capMark) then
        -- NOTE first (or only)
        if capMark > noteMark then
            set noteBody to my trimText(text (noteMark + 5) thru (capMark - 1) of t)
            set captionBody to my trimText(text (capMark + 8) thru -1 of t)
        else
            set noteBody to my trimText(text (noteMark + 5) thru -1 of t)
        end if
    else if capMark > 0 and (noteMark is 0 or capMark < noteMark) then
        -- CAPTION first (or only)
        if noteMark > capMark then
            set captionBody to my trimText(text (capMark + 8) thru (noteMark - 1) of t)
            set noteBody to my trimText(text (noteMark + 5) thru -1 of t)
        else
            set captionBody to my trimText(text (capMark + 8) thru -1 of t)
        end if
    end if
    return {noteBody, captionBody}
end parseDescriptionSections

on formatDescription(noteBody, captionBody)
    set n to my trimText(noteBody)
    set c to my trimText(captionBody)
    set parts to {}
    if n is not "" then set end of parts to "NOTE: " & n
    if c is not "" then set end of parts to "CAPTION: " & c
    if (count of parts) is 0 then return ""
    -- LF joins cleanly in IPTC/XMP; exiftool -s3 shows CR as "."
    set AppleScript's text item delimiters to linefeed
    set out to parts as text
    set AppleScript's text item delimiters to ""
    return out
end formatDescription

-- Parse "key=value|key=value|..." (keys are case-sensitive as sent by VBA).
on fieldValue(tagBlob, keyName)
    set AppleScript's text item delimiters to "|"
    set bits to text items of tagBlob
    set AppleScript's text item delimiters to ""
    set prefix to keyName & "="
    repeat with bit in bits
        set t to my trimText(bit as text)
        if t starts with prefix then
            if (length of t) = (length of prefix) then return ""
            return text ((length of prefix) + 1) thru -1 of t
        end if
    end repeat
    return ""
end fieldValue

-- Y if this file is a camera RAW, or a sibling RAW shares the same basename (Sony/Canon/Nikon/Fuji/DNG).
on siblingRawFlag(posixPath)
    set sh to "f=" & quoted form of posixPath & "; d=$(dirname \"$f\"); b=$(basename \"$f\"); stem=${b%.*}; ext=$(printf %s \"${b##*.}\" | tr '[:upper:]' '[:lower:]'); case \"$ext\" in arw|cr2|cr3|nef|nrw|raf|dng) echo Y; exit 0;; esac; for e in arw ARW cr2 CR2 cr3 CR3 nef NEF nrw NRW raf RAF dng DNG; do if [ -f \"$d/$stem.$e\" ]; then echo Y; exit 0; fi; done; echo N"
    try
        set flag to do shell script sh
        if flag is "Y" or flag is "N" then return flag
    end try
    return "N"
end siblingRawFlag

on upperText(s)
    try
        return do shell script "printf %s " & quoted form of (s as text) & " | tr '[:lower:]' '[:upper:]'"
    on error
        return s as text
    end try
end upperText

on trimText(s)
    set s to s as text
    repeat while s starts with " " or s starts with tab or s starts with return or s starts with linefeed
        if (length of s) is 0 then exit repeat
        set s to text 2 thru -1 of s
    end repeat
    repeat while s ends with " " or s ends with tab or s ends with return or s ends with linefeed
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

    -- 12 FocusDistance2; 13 GPSPosition; 14 SerialNumber; 15 ImageHeight (VertPix); 16 ImageWidth (HorPix);
    -- then HaveRaw Y/N; then Sequence*.
    set fmt to "$FileName|$FileTypeExtension|$MIMEType|$DateTimeOriginal|$CreateDate|$FileSize|$ExposureTime|$FNumber|$ISO|$Model|$ExposureProgram|$LensID|$FocusDistance2|$GPSPosition|$SerialNumber|$ImageHeight|$ImageWidth"
    set cmd to quoted form of exif & " -m -s3 -p " & quoted form of fmt & " " & quoted form of posixPath
    try
        set meta to do shell script cmd
    on error errMsg
        set fmt2 to "$FileName|$FileTypeExtension|$MIMEType|$DateTimeOriginal|$CreateDate|$FileSize|$ExposureTime|$FNumber|$ISO|$Model|$ExposureProgram|$LensModel|$FocusDistance2|$GPSPosition|$SerialNumber|$ImageHeight|$ImageWidth"
        set cmd2 to quoted form of exif & " -m -s3 -p " & quoted form of fmt2 & " " & quoted form of posixPath
        try
            set meta to do shell script cmd2
        on error errMsg2
            return "ERROR" & tab & errMsg2
        end try
    end try

    set meta to meta & "|" & (my siblingRawFlag(posixPath))
    set seqFields to my readSequenceFields(posixPath)
    set meta to meta & "|" & (item 1 of seqFields) & "|" & (item 2 of seqFields) & "|" & (item 3 of seqFields)
    -- Fourth/fifth fields: NOTE and CAPTION bodies from Description (legacy plain = note)
    set descParts to my parseDescriptionSections(my readDescriptionRaw(posixPath))
    set noteOnly to item 1 of descParts
    set captionOnly to item 2 of descParts
    return "OK" & tab & posixPath & tab & meta & tab & noteOnly & tab & captionOnly
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
