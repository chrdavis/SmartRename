# SmartRename

Have you ever needed to modify the file names of a large number of files but didn't want to rename all of the files the same name? Wanted to do a simple search/replace on a sub-string of various file names? Wanted to perform a regular expression rename on multiple items?  

SmartRename is a Windows Shell Extension for advanced bulk renaming using search and replace or regular expressions.  SmartRename allows simple search and replace or more advanced regular expression matching.  While you type in the search and replace input fields, the preview area will show what the items will be renamed to.  SmartRename then calls into the Windows Explorer file operations engine to perform the rename.  This has the benefit of allowing the rename operation to be undone after SmartRename exits.

If you find this useful, feel free to buy me a coffee!

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=chrisdavis%40outlook%2ecom&lc=US&item_name=Chris%20Davis&item_number=SmartRename&no_note=0&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHostedGuest)

### Download
[Latest 32 and 64 bit versions](https://github.com/chrdavis/SmartRename/releases/latest) 
Windows Vista,7,8,10

You will likely need to restart windows for the extension to be picked up by Windows Explorer.

### Demo

In the below example, I am replacing all of the instances of "Pampalona" with "Pamplona" from all the image filenames in the folder.  Since all the files are uniquely named, this would have taken a long time to complete manually.  With SmartRename this tasks seconds.  Notice that I can undo the rename if I want to from the Windows Explorer context menu.

![Image description](/Images/SmartRenameDemo.gif)

