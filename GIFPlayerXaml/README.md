### Here's another sample project, this time showcasing how to use a XAML Island inside a VB6 Form! This project is an **Animated GIF Player**, playing a list of some random animated GIFs directly from the Internet. You can easily expand this menu to include additional images (GIF or any other format) either from the Internet or from your computer:

<img width="516" height="278" alt="GIFPlayer" src="https://github.com/user-attachments/assets/18095153-4385-4797-b472-5ba5080859e4" />

A _XAML Island_ can be embedded in any VB6 control that has a **hWnd** property (including the Form itself), here we are using a vanilla **PictureBox** as a container for the _XAML Island_. As for the actual XAML controls, we are using a **Grid** control (for setting the background color to match the Form's color so that transparent GIFs don't play against a black background) and an **Image** control to display the actual GIF image (not to be confused with the classic VB6 Image control!).

You can compile and run the executable as it is but if you want to run this project from the VB6 IDE then there's an additional step that needs to be performed:

- If your VB6.EXE file is not manifested then simply copy the included **VB6.EXE.manifest** file (from the **Common** folder) in the same folder as VB6.EXE (usually: _C:\Program Files (x86)\Microsoft Visual Studio\VB98_). Probably requires a computer restart so that the manifest takes effect.

- If your VB6.EXE is already manifested then you need to edit its manifest (with Resource Hacker for example) and add the **\<maxversiontested Id="10.0.18362.0"\/\>** attribute as shown in the included manifest file.
