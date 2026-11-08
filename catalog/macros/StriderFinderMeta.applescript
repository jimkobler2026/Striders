-- Excel AppleScriptTask helpers: Finder selection + exiftool metadata
-- Photo return: OK <tab> path <tab> Name|ext|mime|dto|create|size|shutter|aperture|iso|camera|mode|lens
-- Video return: OK <tab> path <tab> Name|ext|mime|create|mediaCreate|size|duration|fps|width|height|deviceModel|model|shutter|aperture|iso

on loadFromFinder(dummy)
    set p to my selectedPOSIXPath()
    if p is "" then return "ERROR" & tab & "Select a photo or video in Finder first."
    return my metadataLine(p)
end loadFromFinder

on loadFromPreview(dummy)
    set p to my frontPreviewPOSIXPath()
    if p is "" then return "ERROR" & tab & "Open a photo in Preview first (front document is used)."
    return my metadataLine(p)
end loadFromPreview

on loadVideoFromFinder(dummy)
    set p to my selectedPOSIXPath()
    if p is "" then return "ERROR" & tab & "Select a video in Finder first."
    return my videoMetadataLine(p)
end loadVideoFromFinder

on loadVideoFromPreview(dummy)
    set p to my frontPreviewPOSIXPath()
    if p is "" then return "ERROR" & tab & "Open a video in Preview first (front document is used)."
    return my videoMetadataLine(p)
end loadVideoFromPreview

on readPath(posixPath)
    if posixPath is "" then return "ERROR" & tab & "No path provided."
    return my metadataLine(posixPath)
end readPath

on frontPreviewPOSIXPath(dummy)
    try
        tell application "Preview"
            if (count of documents) is 0 then return ""
            try
                return path of front document
            on error
                try
                    return path of document of front window
                on error
                    return ""
                end try
            end try
        end tell
    on error
        return ""
    end try
end frontPreviewPOSIXPath

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

-- payload: path <tab> Rating|Species|Behavior|Locale|SendToUK|View|Lighting|notes
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

    set kwList to {}
    if species is not "" then set end of kwList to "Species:" & species
    if behavior is not "" then set end of kwList to "Behavior:" & behavior
    if locale is not "" then set end of kwList to "Locale:" & locale
    if sendUK is not "" then set end of kwList to "SendToUK:" & sendUK
    if viewTag is not "" then set end of kwList to "View:" & viewTag
    if lighting is not "" then set end of kwList to "Lighting:" & lighting

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

    return "OK" & tab & posixPath & tab & meta
end metadataLine

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
