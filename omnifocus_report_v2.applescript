global showTaskDate
global showTaskContext
global showProjectStatus
global flattenTasks

set showTaskDate to false
set showTaskContext to false
set showProjectStatus to true
set flattenTasks to true

set outputToFile to true
set outputDirectory to (path to home folder as Unicode text) & "ANL:Task Tracking:"


tell application "OmniFocus"
	
	tell default document
		--SET THE REPORT TITLE
		set ExportList to "Current Active Projects" & return & "---" & return & return as Unicode text
		
		set strActive to ""
		set dteNow to (current date)
		
		--GENERATE LIST OF ACTIVE PROJECTS
		set strProjActive to ""
		repeat with proj in (projects where its status is active and it is not singleton action holder and (its defer date is missing value or defer date < dteNow))
			set breadcrumb to "### " & name of proj
			set strProjActive to (breadcrumb & return & strProjActive) & return
		end repeat
		
		repeat with oFolder in (folders where hidden is false)
			if class of oFolder is folder then
				set breadcrumb to "### " & name of oFolder
				set strActive to strActive & my RecurseActiveProjects(oFolder, breadcrumb, dteNow)
			end if
		end repeat
		
		set ExportList to ExportList & strProjActive & strActive
		
		--PROCESS THE PROJECTS			
		set strText to ""
		set start to (my DateOfThisInstanceOfThisWeekdayBeforeOrAfterThisDate(current date, Monday, 0)) - 7 * days
		set ExportList to ExportList & "Completed Tasks From Last Week - " & my dateISOformat(start, ".") & return & "---" & return & return & ""
		
		repeat with oFolder in (folders where hidden is false)
			if class of oFolder is folder then
				set breadcrumb to "### " & name of oFolder
				set strText to strText & my RecurseFolders(oFolder, start, breadcrumb)
			end if
		end repeat
		set strText to strText & my ParseProjects(projects, start, "")
		
	end tell
	set ExportList to ExportList & strText & return & "Generated: " & (current date) as Unicode text
	
end tell

if outputToFile then
	--CHOOSE FILE NAME FOR EXPORT AND SAVE AS MARKDOWN
	set outputFile to outputDirectory & year of start & ":" & my dateISOformat(start, "-") & ".md"
	tell application "System Events"
		set fid to (open for access file outputFile with write permission)
		set eof fid to 0
		write ExportList to fid
		close access fid
	end tell
end if


on RecurseFolders(iFolder, start, breadcrumb)
	tell application "OmniFocus"
		tell default document
			set strText to ""
			repeat with oFolder in (folders of iFolder where hidden is false)
				if class of oFolder is folder then
					set newbreadcrumb to breadcrumb & " / " & name of oFolder
					set strText to strText & my RecurseFolders(oFolder, start, newbreadcrumb)
				end if
			end repeat
			set strText to strText & my ParseProjects(projects of iFolder, start, breadcrumb)
			
			return strText
			
		end tell
	end tell
end RecurseFolders


on RecurseActiveProjects(iFolder, breadcrumb, dteNow)
	tell application "OmniFocus"
		tell default document
			set strText to ""
			set strActive to ""
			repeat with proj in (projects of iFolder where its status is active and it is not singleton action holder and (its defer date is missing value or defer date < dteNow))
				set strActive to (strActive & "- " & name of proj as string) & return
			end repeat
			if strActive is not "" then
				set strActive to breadcrumb & return & strActive & return
			end if
			
			repeat with oFolder in (folders of iFolder where hidden is false)
				if class of oFolder is folder then
					set newbreadcrumb to breadcrumb & " / " & name of oFolder
					set strText to strText & my RecurseActiveProjects(oFolder, newbreadcrumb, dteNow)
				end if
			end repeat
			
			return strActive & strText
			
		end tell
	end tell
end RecurseActiveProjects


on ParseProjects(iProjects, start, breadcrumb)
	tell application "OmniFocus"
		tell default document
			set week_ago_end to start + 7 * days
			set oldFullPath to ""
			set strText to ""
			
			repeat with oProject in iProjects
				if breadcrumb is "" then
					set newbreadcrumb to "### " & name of oProject
				else
					set newbreadcrumb to breadcrumb & " / " & name of oProject
				end if
				if showProjectStatus then
					set newbreadcrumb to newbreadcrumb & " [ " & status of oProject & " ]"
				end if
				
				if flattenTasks then
					set refDoneInLastWorkWeek to (a reference to (flattened tasks of oProject where (completion date ³ start and completion date < week_ago_end)))
				else
					set refDoneInLastWorkWeek to (a reference to (tasks of oProject where (completion date ³ start and completion date < week_ago_end)))
				end if
				set {lstName, lstContext, lstProject, lstDate} to {name, name of its context, containing project, completion date} of refDoneInLastWorkWeek
				set cTasks to my CompletedTasks(lstName, lstContext, lstProject, lstDate)
				if cTasks is not "" then
					set strText to strText & newbreadcrumb & return
					set strText to strText & cTasks & return
				end if
				
			end repeat
			
			return strText
			
		end tell
	end tell
end ParseProjects


on CompletedTasks(lstName, lstContext, lstProject, lstDate)
	
	try
		get showTaskDate
	on error
		set showTaskDate to true
	end try
	try
		get showTaskContext
	on error
		set showTaskContext to false
	end try
	
	set strText to ""
	repeat with iTask from 1 to length of lstName
		set {strName, varContext, varProject, varDate} to {item iTask of lstName, item iTask of lstContext, item iTask of lstProject, item iTask of lstDate}
		
		set strText to strText & "- "
		
		if varDate is not missing value and showTaskDate then set strText to strText & my dateISOformat(varDate, ".") & " - "
		
		set strText to strText & strName
		
		if varContext is not missing value and showTaskContext then set strText to strText & " *@" & varContext & "*"
		set strText to strText & return
	end repeat
	
	return strText
	
end CompletedTasks


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
on dateISOformat(theDate, delm)
	set y to text -4 thru -1 of ("0000" & (year of theDate))
	set m to text -2 thru -1 of ("00" & ((month of theDate) as integer))
	set d to text -2 thru -1 of ("00" & (day of theDate))
	return y & delm & m & delm & d
end dateISOformat