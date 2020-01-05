# NetCheck
Network connectivity monitor
NetCheck

Version: 1.0
Developer: Tyson Ackland
Licence: MIT

Introduction

NetCheck is a utility to monitor the connectivity to nominated destinations.  Destinations could be on your local network (eg router, printers) or further afield (eg your ISP).  I actually wanted to use SmokePing on my Windows PC but it was developed for Linux and after spending some days trying to install the various components I gave up in frustration 😊.  So in NetCheck I have tried to replicate the set of graphs that SmokePing produces.

NetCheck is multi-threaded – quite the learning curve for me!  Each destination you add spawns a dedicated thread to gather the network data over the whole time NetCheck is running.  Consequently, NetCheck needs to be left running continuously in order to gather data over long periods.  It is therefore recommended to allow it to run automatically each time you log on to Windows.  On Windows 10, you need to go into Settings-Apps-Start-up and allow NetCheck.

NetCheck saves its data to the file system so that you don’t lose it all each time it restarts eg after a reboot of your system.

Adding Sites

When NetCheck first runs, it will display its main window so that you can start adding sites.  Sites can be added by hostname (eg www.google.com) or IP address (eg 192.168.0.1).  If you enter an IP address, NetCheck will automatically convert it to a hostname if one exists.

Alternatively, you can use the Discover LAN button to have NetCheck automatically add all devices it can find attached to your local network.  As this takes some time, it runs in a separate thread so that you can continue to use NetCheck while it adds the local devices.

NetCheck saves its data in the standard Application Data path in the file system.

On subsequent runs when NetCheck has destinations to monitor, it starts minimised with its icon in the System Tray.  You can open it by double-clicking on this icon.

Removing Sites

When sites are removed NetCheck stops the data gathering thread and removes the associated data files.

Viewing Sites

Gathered data is displayed when you select a site.  Each graph displays an ever-increasing period to provide a long-term view of the health of your connection to that site.  Health is depicted by showing the average response time to pings.  The minimum and maximum response times are recorded, and the range is also displayed on the graphs.

Graph/Ping Frequency/Average, Minimum, Maximum Recording Frequency
 
last 60 mins/5 seconds/n/a

Last 24 hrs/5 seconds/5 mins

last 7 days/5 seconds/30 mins

last 30 days/5 seconds/2 hrs

Last 12 months/5 seconds/24 hrs


 

The Y axis scale automatically adjusts as data arrives.  If data has wide variances, the scale on the Y axis displays breaks so that the extreme data points appear.


Interactive Graphs

You can zoom into a graph by dragging your mouse over the region of interest.  This process can be repeated so that you can keep zooming in.  You can also zoom in using + or PageUp and zoom out using – or PageDown.  Press Esc to revert to the previous level of zoom.

The initial view in each graph is zoomed to try and present the data clearly.

While zoomed, you can use the scrollbars on either axis to move your view along that axis.
