(*
OmniFocus - Weekly Project Report Generator
Author Ben Allen
Derived from work authored by Chris Brogan and Rob Trew

// GITHUB
https://github.com/bensallen/omnifocus_report

Original Script Posted At: http://veritrope.com/code/omnifocus-weekly-project-report-generator

*)

(*
======================================
// MAIN PROGRAM
======================================
*)

tell application "OmniFocus"
	
	--SET THE REPORT TITLE
	set ExportList to "Current Active Projects" & return & "---" & return as Unicode text
	
	--PROCESS THE PROJECTS
	tell default document
		set refFolders to a reference to (flattened folders where hidden is false)
		repeat with idFolder in (id of refFolders) as list
			set oFolder to folder id idFolder
			set ExportList to ExportList & my IndentAndProjects(oFolder) & return
		end repeat
		
		--ASSEMBLE THE COMPLETED TASK LIST
		set week_ago_start to my DateOfThisInstanceOfThisWeekdayBeforeOrAfterThisDate(current date, Monday, -1)
		set week_ago_end to week_ago_start + 7 * days
		set ExportList to ExportList & "Completed Tasks From Last Week - " & my dateISOformat(week_ago_start) & return & "---" & return & return & ""
		
		set refDoneInLastWorkWeek to a reference to (flattened tasks where (completion date ³ week_ago_start and completion date < week_ago_end))
		set {lstName, lstContext, lstProject, lstDate} to {name, name of its context, containing project, completion date} of refDoneInLastWorkWeek
		set strText to ""
		set oldFullPath to ""
		repeat with iTask from 1 to length of lstName
			set {strName, varContext, varProject, varDate} to {item iTask of lstName, item iTask of lstContext, item iTask of lstProject, item iTask of lstDate}
			set pParent to container of varProject
			set fullPath to name of varProject
			repeat while class of pParent is folder
				set fullPath to name of pParent & " / " & fullPath
				set pParent to container of pParent
			end repeat
			(*if oldFullPath is not fullPath then set strText to strText & "### " & fullPath & return*)
			set oldFullPath to fullPath
			
			set strText to strText & "- "
			
			if varDate is not missing value then set strText to strText & short date string of varDate & " - "
			
			(*if varProject is not missing value then set strText to strText & " [" & fullPath & "] - "
			*)
			set strText to strText & strName
			if varContext is not missing value then set strText to strText & " *@" & varContext & "*"
			set strText to strText & " " & return
		end repeat
	end tell
	set ExportList to ExportList & strText & return & "Generated: " & (current date) as Unicode text
	
	--CHOOSE FILE NAME FOR EXPORT AND SAVE AS MARKDOWN
	set taskTrackingFolder to path to home folder
	
	set fn to choose file name with prompt "Name this file" default name my dateISOformat(week_ago_start) & Â
		".md" default location taskTrackingFolder
	tell application "System Events"
		set fid to open for access fn with write permission
		write ExportList to fid
		close access fid
	end tell
end tell

(*
======================================
// MAIN HANDLER SUBROUTINES
======================================
*)

on IndentAndProjects(oFolder)
	tell application id "OFOC"
		
		set {dlm, my text item delimiters} to {my text item delimiters, return & return}
		set dteNow to (current date)
		set strActive to ""
		repeat with proj in (projects of oFolder where its status is active and it is not singleton action holder and its defer date is missing value or defer date < dteNow)
			set strActive to (strActive & "- " & name of proj as string) & return
		end repeat
		
		if strActive is not "" then
			set oParent to container of oFolder
			
			set fullPath to name of oFolder
			repeat while class of oParent is folder
				set fullPath to name of oParent & " / " & fullPath
				set oParent to container of oParent
			end repeat
			
			set my text item delimiters to dlm
			
			return "## " & fullPath & " ##" & return & return & strActive
		else
			return ""
		end if
	end tell
end IndentAndProjects

-- Source: http://macscripter.net/viewtopic.php?id=39553
on DateOfThisInstanceOfThisWeekdayBeforeOrAfterThisDate(d, w, i) -- returns a date
	-- Keep an note of whether the instance value *starts* as zero
	set instanceIsZero to (i is 0)
	-- Increment negative instances to compensate for the following subtraction loop
	if i < 0 and d's weekday is not w then set i to i + 1
	-- Subtract a day at a time until the required weekday is reached
	repeat until d's weekday is w
		set d to d - days
		-- Increment an original zero instance to 1 if subtracting from Sunday into Saturday 
		if instanceIsZero and d's weekday is Saturday then set i to 1
	end repeat
	-- Add (adjusted instance) * weeks to the date just obtained and zero the time
	d + i * weeks - (d's time)
end DateOfThisInstanceOfThisWeekdayBeforeOrAfterThisDate

-- Source: http://henrysmac.org/blog/2014/1/4/formatting-short-dates-in-applescript.html
on dateISOformat(theDate)
	set y to text -4 thru -1 of ("0000" & (year of theDate))
	set m to text -2 thru -1 of ("00" & ((month of theDate) as integer))
	set d to text -2 thru -1 of ("00" & (day of theDate))
	return y & "-" & m & "-" & d
end dateISOformat