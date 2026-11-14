-- Excel AppleScriptTask: generate a Nat Geo–style caption via scripts/caption_photo.py
-- Default provider is Gemini Flash (cheap + good enough). Override with payload provider.
--
-- payload: photoPOSIXPath  OR  photoPOSIXPath <tab> provider
--   provider: google | anthropic | openai | openrouter | groq
-- Returns: OK <tab> provider <tab> model <tab> words <tab> caption
--      or: ERROR <tab> message

-- Editable caption popup for Mac Excel (UserForms are not available in Mac VBE).
-- payload: current caption text
-- Returns: OK <tab> edited text   or   CANCEL
on editCaption(payload)
    set currentText to payload
    try
        set dlg to display dialog "Edit caption (Save writes catalog + file):" default answer currentText buttons {"Cancel", "Save"} default button "Save" with title "Caption"
        return "OK" & tab & (text returned of dlg)
    on error errMsg number errNum
        if errNum is -128 then
            return "CANCEL"
        else
            return "ERROR" & tab & errMsg
        end if
    end try
end editCaption

on captionPhoto(payload)
    set photoPath to ""
    set provider to "google"
    try
        set AppleScript's text item delimiters to tab
        set parts to text items of payload
        set AppleScript's text item delimiters to ""
        if (count of parts) >= 1 then set photoPath to item 1 of parts
        if (count of parts) >= 2 then
            set p to item 2 of parts
            if p is not "" then set provider to p
        end if
    on error
        set photoPath to payload
    end try

    if photoPath is "" then return "ERROR" & tab & "No photo path."

    set sh to (POSIX path of (path to home folder)) & "Library/Application Scripts/com.microsoft.Excel/StriderCaption.sh"
    try
        do shell script "test -x " & quoted form of sh
    on error
        return "ERROR" & tab & "Missing StriderCaption.sh — run scripts/install_mac_helpers.sh"
    end try

    try
        set out to do shell script quoted form of sh & " " & quoted form of photoPath & " " & quoted form of provider
        return out
    on error errMsg
        return "ERROR" & tab & errMsg
    end try
end captionPhoto
