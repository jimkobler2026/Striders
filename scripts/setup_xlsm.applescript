-- Convert the built .xlsx into a macro-enabled workbook and import VBA.
-- Run: osascript scripts/setup_xlsm.applescript

on run
    set projectRoot to do shell script "cd \"$(dirname " & quoted form of POSIX path of (path to me) & ")/..\" && pwd"
    -- path to me is unreliable for osascript file args; use argv-free fixed relative from cwd instead
end run
