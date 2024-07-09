# Description
Import Google Calendar Entries and Events into Google Sheets for timekeeping via Apps Scripts

## Use Case
I am working in a Company and I'm required to submit my timesheets in a software like HARVEST (https://www.getharvest.com/). In my timesheets usually I input how much time I wanted for which category. Now, my Manager wants me to add descriptions to my timesheets as well.

My Manager and colleagues also want me to block my calendar (Google Calendar), if I don't want to be disturbed to focus on some task. My colleagues, managers and client schedule calls on a daily basis. My Manager also checks my calender to see if I had attended any calls. Sometimes, my colleagues deny having spent time with me on a task.

I also add comments on my tickets on platforms like JIRA, Trello, etc. So that I don't forget what task I worked on, I used to log small notes in notes.

At the end of the day/week, when I sit down to fill in my timesheets; tt's hard to keep track of everything and make things consistent. This takes additonal times if I want the timesheets to be detailed as well.


## Solution
I log everything in my Google Calendar, as I am working on them.

- I attended a call? I create an event and add the people on the call.
- I worked on ticket? I create an event and put the ticket and link(optional)
- Worked with/helped a colleage? I create an event and sent an invite to the Collegue. That way the colleague also knows how much time we both have to put our respective timesheets and we both have accountability.
- Took a call with client on the phone/Zoom? Log it as Calendar Event
- ...

After logging everything, at the end of the day/week, I just run this script. I get the estimated time and text. Using that I fill up my timesheet.


## Setup
1) Create a Google Sheet with the template
2) Configure your GMAIL id in "Calendar ID"
3) COnfigure start and end date for the time range you want to be updated.
4) From Menu, Extensions > Apps Script
5) In Apps Script, Copy paste source code and save
6) Run
7) Go back to the Google Sheet. See the updated values from row 7. Modify values in columns J,K,L.
8) Log the final text in timesheet manually. (This can be automated by you based on your timesheet software)