outlook-2011-gtd-shortcuts
==========================

A couple of brief AppleScripts for filing messages in Outlook 2011

Copy into `~/Library/Application Support/Microsoft/Office/Outlook Script Menu Items`

Add application-specific keyboard shortcuts from System Preferences

![Keyboard Shortcuts](https://raw.githubusercontent.com/jakub/outlook-2011-gtd-shortcuts/master/keyboard.png)

Which should then appear in Outlook:

![Outlook's script menu](https://raw.githubusercontent.com/jakub/outlook-2011-gtd-shortcuts/master/menubar.png)

Scripts are of the form:
```applescript
on run {}
	
	tell application "Microsoft Outlook"
		
		set listMessages to current messages
		if ((count of listMessages) < 1) then return
		
		set gtdFolder to folder "@action"
		
		repeat with objInSelection in listMessages
			if (class of account of objInSelection is not imap account) then
				move objInSelection to gtdFolder
				set category of objInSelection to {category "Not in OmniFocus"}
			end if
		end repeat
		
	end tell
	
end run
```

In this case, selected messages are moved to the folder `@action`. This should be a unique name.

You can optionally assign categories to messages as well, for example `set category of objInSelection to {category "Not in OmniFocus"}`

*Note: I only have a single Exchange account configured. Multiple accounts or IMAP accounts may not work as expected.*
