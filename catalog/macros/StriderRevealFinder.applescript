-- Installed as StriderRevealFinder.scpt for Excel AppleScriptTask
on revealInFinder(pathList)
    set AppleScript's text item delimiters to {character id 0}
    set thePaths to text items of pathList
    set AppleScript's text item delimiters to {""}

    set fileList to {}
    set missing to {}
    repeat with p in thePaths
        set pText to p as text
        if length of pText > 0 then
            try
                set end of fileList to (POSIX file pText as alias)
            on error
                set end of missing to pText
            end try
        end if
    end repeat

    if (count of fileList) is 0 then
        return "None of the selected files were found on disk."
    end if

    tell application "Finder"
        activate
        reveal fileList
        select fileList
    end tell

    if (count of missing) > 0 then
        return "Revealed " & (count of fileList) & " file(s). Missing: " & (count of missing)
    end if
    return ""
end revealInFinder
