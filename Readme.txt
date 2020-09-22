This is a demonstration of how you can minimize your app to the system tray and reactivates it by double clicking in a registered file type for your app from Explorer, without leaving the second instance running.

When you minimize it, the following occurs:

1 - The file type "*.tst" is registered into Windows registry. If you want to view it, open Regedit and look at HKEY_CLASSES_ROOT\.tst, then look at HKEY_CLASSES_ROOT\Hook test;
2 - The app starts hooking (trapping the messages that come from Windows);
3 - The app is minimized to the systray. You can restore it by left or right clicking on its icon in the systray.

Now all that you need to do is open Explorer and double click on any file ".tst", and the following will occur:
1 - Explorer opens a second instance of the program, passing to it the command line, ie the name of the file;
2 - The second instance will find that there is a previous one running (App.PrevInstance);
3 - The FindWindow function locates the handle for the first instance;
4 - The second instance sends the command line to the first one (SendMessage), wich is "listening" (hooking).
5 - The first one receives the message and starts processing;
6 - The second one ends.

Please note that the FindWindow function will only work with the compiled version, that creates a class named "ThunderRT6FormDC" for the program - VB6. VB5 creates another class name, but I don't remember it for now.
In IDE the class created is "ThunderFormDC".
Also it is important to set the form caption after FindWindow, else it will find two classes with the same window text (caption).
 
You can also define another icon for the file type the program is registering.
Just create a resource file using resource editor (if its not installed, click on Add-Ins / Add-In Manager, VB6 Resource Editor and Load on Startup), "Add Icon" and insert the icon that will be visible in Explorer for the file type you are registering.

These functionality was difficult to find and debug, but I know that are much more improvements that can be done. 

Thanks should also go to Paul Mather (PSC) for the Registry functions. His module is great!


Alexandre Moro
alb_moro@ig.com.br

Cheers from Brasil!
