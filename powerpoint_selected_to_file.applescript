tell application "Microsoft PowerPoint"
	set allText to "" -- Initialize a variable to store all the text
	
	-- Get the reference to the active presentation
	set thePresentation to active presentation
	set slideCount to count slides of thePresentation
	
	-- Iterate through each slide
	repeat with i from 1 to slideCount
		set tSlide to slide i of thePresentation
		set tShapes to get shapes of tSlide
		set slideText to ""
		
		-- Iterate through shapes in the slide
		repeat with t_shape in tShapes
			if has text frame of t_shape then
				if has text of text frame of t_shape then
					set shapeText to content of text range of text frame of t_shape
					set slideText to slideText & shapeText & return
				end if
			end if
		end repeat
		
		-- Check the slide text and concatenate
		
		set allText to allText & slideText & return
		
	end repeat
end tell

-- Writing to a file
tell application "Finder"
	set desktopPath to (path to desktop folder as text)
	set textFilePath to desktopPath & "PowerPointNotes3.txt"
	set fileRef to (open for access file textFilePath with write permission)
	try
		write allText to fileRef starting at eof
	on error
		close access fileRef
	end try
	close access fileRef
end tell
