# COMMethodSearcher
This is a PowerShell script that searches through all COM objects and look for any methods containing a keyword of your choosing. There's three "depths" that are possible. The first "depth" is just getting the members of the COM object. Second "depth" is getting the member's members of the COM object. The third "depth", is you guessed it, the member's member's members of the object. I figured three was enough. The first two depths complete within a few minutes, the third can take 5-8 minutes (in my experience). There's also two options of searching: By CLSID or ProgID. Not all COM Objects will have ProgIDs, but searching by CLSIDs takes almost twice as long and has caused some unexpected Powershell crashes.

This has caused some weird shit to happen because it's literally instantiating every registered COM object, so use at your own risk. Because of this, this isn't supported, so if one day you log into Windows and see Microsoft Word pop up 12 times shortly after you ran this script, then don't open an issue because I warned you. Godspeed.

Usage Example: .\CMS.ps1 -CLSIDs -Depth 3 -Term ExecuteShell

Usage Example: .\CMS.ps1 -ProgIDs -Depth 2 -Term ExecuteShell


![Example](https://i.imgur.com/R7Bx5fb.png)
