-- This script will import to-dos and Projects into Things from OmniFocus.
-- To run the script, click the green "Run" button in the toolbar above.

on run

  display dialog "Make sure you have all tasks in Omni Focus selected that you want to import into Things." buttons {"OK", "Cancel"} default button 1

  tell application "OmniFocus.app"
    tell «class FCDo»
      if number of «class FCdw» is 0 then
        make new «class FCdw» with properties {bounds:{0, 0, 1000, 500}}
      end if
    end tell

    tell «class FCdw» 1 of front document
      set lstTrees to every «class OTst» of «class FCcn»
      if (count of lstTrees) = 0 then
        try
          display dialog "Nothing selected in the right-hand panel." & return & return & "Select material to export, and try again." & return
        end try
      else
        set blnContext to («class FCvm» is not equal to "project")
        set lngIndent to 0
        my ExportTrees(lstTrees, lngIndent, blnContext)

      end if
    end tell
  end tell
end run
-- -------------------------------------
-- Walks the omni focus tree
-- -------------------------------------
on ExportTrees(lstTrees, lngIndent, blnContextView)

  using terms from application "OmniFocus.app"
    repeat with oTree in lstTrees
      -- intialize task string
      set strNotes to ""
      set strTags to ""
      set strProject to ""

      set oValue to «class valL» of oTree

      try
        set strName to name of oValue
      on error
        set strName to "Inbox"
      end try

      if strName ≠ "Inbox" then
        set strNote to «class FCno» of oValue

        set oContext to «class FCct» of oValue
        if oContext is not equal to missing value then
          set strTags to " @" & name of oContext & ","
        end if

        set DueDate to «class FCDd» of oValue
        if «class FCfl» of oValue then
          set strTags to strTags & " Flagged" & ","
        end if

        set aProject to «class FCPr» of oValue
        if aProject is not equal to missing value then
          set ProjectName to name of aProject
        else
          set ProjectName to "NoProjInOmniFocus"
        end if

      end if

      set clValue to class of oValue

      if (clValue is equal to «class FCac») then
        try
          if DueDate is equal to missing value then
            tell application "Things"
              set newToDo to make new to do with properties {name:strName, tag names:strTags, notes:strNote} at beginning of project ProjectName
            end tell
          else
            tell application "Things"
              set newToDo to make new to do with properties {name:strName, tag names:strTags, due date:DueDate, notes:strNote} at beginning of project ProjectName
            end tell
          end if
        on error
          tell application "Things"
            set newProject to make new project with properties {name:ProjectName}
          end tell
          if DueDate is equal to missing value then
            tell application "Things"
              set newToDo to make new to do with properties {name:strName, tag names:strTags, notes:strNote} at beginning of project ProjectName
            end tell
          else
            tell application "Things"
              set newToDo to make new to do with properties {name:strName, tag names:strTags, due date:DueDate, notes:strNote} at beginning of project ProjectName
            end tell
          end if
        end try

      end if

      if (clValue is equal to «class FCit») then
        if DueDate is equal to missing value then
          tell application "Things"
            set newToDo to make new to do with properties {name:strName, tag names:strTags, notes:strNote}
          end tell
        else
          tell application "Things"
            set newToDo to make new to do with properties {name:strName, tag names:strTags, due date:DueDate, notes:strNote}
          end tell
        end if
      end if

    end repeat
  end using terms from

end ExportTrees
