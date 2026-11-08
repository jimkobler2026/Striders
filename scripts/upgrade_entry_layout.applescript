-- Rebuild the Entry sheet layout for metadata fields.
-- Run while catalog is available; preserves VBA in the .xlsm.

on run
    set catalogPath to "/Users/jim/Development/projects/Striders/catalog/Strider_Media_Catalog.xlsm"
    tell application "Microsoft Excel"
        activate
        try
            set wb to active workbook
            if (full name of wb) does not contain "Strider_Media_Catalog" then error "wrong book"
        on error
            open POSIX file catalogPath
            delay 1
            set wb to active workbook
        end try

        tell worksheet "Entry" of wb
            activate object
            try
                clear contents range "A1:G20"
            end try

            set value of range "A1" to "Add media"
            set value of range "B1" to "LoadFromFinderSelection"
            set value of range "D1" to "AppendCatalogRow"
            set value of range "F1" to "Finder select → Load → tag → Append"

            set value of range "A2" to "FilePath"
            set value of range "A3" to "FileName"
            set value of range "C3" to "MediaType"
            set value of range "E3" to "FileFormat"
            set value of range "A4" to "DateCaptured"
            set value of range "C4" to "TimeCaptured"
            set value of range "E4" to "FileSize"
            set value of range "A5" to "ShutterSpeed"
            set value of range "C5" to "Aperture"
            set value of range "E5" to "ISO"
            set value of range "A6" to "Behavior"
            set value of range "C6" to "Keep"
            set value of range "E6" to "Quality"
            set value of range "A7" to "Species"
            set value of range "C7" to "LifeStage"
            set value of range "E7" to "ShotType"
            set value of range "A8" to "CausticType"
            set value of range "C8" to "CausticClarity"
            set value of range "E8" to "Location"
            set value of range "A9" to "Notes"

            set value of range "A11" to "Tips"
            set value of range "A12" to "1) Select file in Finder (Preview can stay open).  2) Run LoadFromFinderSelection.  3) Set Behavior/caustics/Keep.  4) Run AppendCatalogRow."
        end tell

        save wb
        return "Entry layout updated: " & (full name of wb)
    end tell
end run
