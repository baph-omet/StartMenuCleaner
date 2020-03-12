# StartMenuCleaner
A quick script to clean up your Start menu by removing empty directories, moving files out of directories if they are alone, and optionally deleting .url files.

I made this program because I was tired of all the clutter in my Start menu on Windows 10 and wanted a way to clean up some of the junk. Turns out, your start menu is basically just a folder of shortcut files that point to the actual applications, so organizing them has no real effect on how the programs run.

To run this program, simply download the lastest .exe (or compile it yourself in Visual Studio), and run it. If you want to also delete web links (`.url` files), run it from the command line and add the `-u` argument. The program does need admin rights in order to modify these folders, so you won't be able to run it if you don't have admin permissions.

If you find any issues or have any suggestions, feel free to open an issue or submit a pull request.
