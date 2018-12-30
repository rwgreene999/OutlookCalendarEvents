# OutlookCalendarEvent

Used to make sure calendar events are noticeable on my screen

## Justification

Have you ever missed a meeting because the outlook calendar reminder was behind another window or on a different monitor than the one you were staring at? 

_It's happened to me!_

I wrote this tool to make sure I had no reason to miss a meeting, unless of course, I want to miss that meeting. :) 

## How does it work

Every minute or so, scan windows for an Outlook Calendar Reminder window.  If one is found bring it to the foreground, so it is noticed.

There is an icon in the tray with a few options, such as pause looking for reminders for 15 minutes.  

## Roadmap

-   Create a better icon 
-   Add option to start up on windows boot
-   Bundle this into a single EXE file
-   Maybe create an installer for those people that feel comfortable with an installation process
-   Add option to flash all the windows on a calendar event 
-   Consider changing to .Net Core 
-   Microsoft changed Skype to use the "Action Center Notifications" which means there is no windows to search for.  If they do the same to Outlook Calendar Reminders, then this will have to be modified for that environment. 

## Technical

-   This is a windows app, using Framework 4.x 
-   It was tested on Windows 10 with office 365.  A previous version of this code ran on Windows 7 with Office 2013 so I think it will work in Office 2013 environments. 
-   I still use Visual Studio instead of VS Code because for this kind of project, it does more for me that VS Code. 
