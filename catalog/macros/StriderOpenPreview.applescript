-- Installed as StriderOpenPreview.scpt for Excel AppleScriptTask
on openInPreview(pathList)
    set AppleScript's text item delimiters to {character id 0}
    set thePaths to text items of pathList
    set AppleScript's text item delimiters to {""}

    set posixList to {}
    repeat with p in thePaths
        set pText to p as text
        if length of pText > 0 then
            set end of posixList to pText
        end if
    end repeat

    if (count of posixList) is 0 then return "No paths provided."

    set fileList to {}
    set missing to {}
    repeat with pText in posixList
        try
            set end of fileList to (POSIX file (pText as text) as alias)
        on error
            set end of missing to (pText as text)
        end try
    end repeat

    if (count of fileList) is 0 then
        return "None of the selected files were found on disk."
    end if

    tell application "Preview"
        activate
        open fileList
    end tell

    if (count of missing) > 0 then
        return "Opened " & (count of fileList) & " file(s). Missing: " & (count of missing)
    end if
    return ""
end openInPreview
