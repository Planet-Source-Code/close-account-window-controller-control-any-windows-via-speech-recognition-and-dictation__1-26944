Q) What is this application?

A) It allows you to control windows and dictate your speech. That means you can talk into your microphone, and it types for you. You can also say the names of command buttons, option buttons, check boxs, etc etc and they will be clicked for you.

Q) Cool! What do I need to use it?

A) - Microphone
   - Microsoft Speech Object Library
   - And you should probally do a training session or two so recognization is better. To do that: Control Panel -> Speech -> Click 'Train Profile' --- It may differ... your smart... figure it out

Q) Ok, got it. How do I use it?

A) Simple example of dictation:

   1) Look in the included directory for 'Releasev1000' for a file called: grpWindowController.vbg - Open it
   2) Once VB has loaded hit F5 (shortcut for run)
   3) Load up notepad
   4) Say something like 'Hello World' to your mic and it will appear in notepad. (Notepad must be the active form.)
   
   Notepad is just an example. It will type into anything. Textbox etc etc

   Simple example of controlling windows:

   1) Load up API Viewer
   2) Type in 'FindWindow'
   3) On the main form of the window controller application click 'Refresh List'
   4) Select the API Viewer application from the combo box
   5) Then press the Control button on the form.
   6) Now say 'Add' and whala the 'Add' button on API Viewer will be clicked and the FindWindow Declaration will be added!!
   7) Now say 'Clear'   poof... its gone
   8) Now say 'Private' the private option box is selected!
   9) Now say 'Add' and it will make private FindWindow declaration.

   You get the idea :) Have fun


Q) WOW! What are you going to add in the next version.

A) I'm going to make it also control menus. (Its partially started if you look in my code)
   I'll make it automatically control the active window so you don't have to manually select the window each time.

   (If anyone knows how to find the top active window please email me. GetActiveWindow doesn't seem to work. And
   GetTopWindow doesn't either :(  Thanks)



David Fiala - djf1010@aol.com - Sept. 03 2001

plz vote for me on planet source code!!!