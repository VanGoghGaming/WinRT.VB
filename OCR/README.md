### This is an OCR Text Recognition project using the modern WinRT Canvas and ScrollViewer controls instead of the classic PictureBox of yonder years. Now you can Ctrl+MouseWheel to Zoom In/Out and also scroll around the picture (both horizontally and vertically).

The bounding rectangles for recognized words have attached tooltips showing the word they encompass and also benefit from built-in antialiasing (more visible for angled text). There's also some basic preprocessing done on the image to help with better recognizing text in lower quality images. The image is enlarged 1.5x times and converted to grayscale. This does improve results although nothing groundbreaking.

<img width="817" height="839" alt="OCR_WinRT" src="https://github.com/user-attachments/assets/29132321-1db8-46d8-bf8d-57f2cc36ad6d" />

It also showcases the WinRT **Clipboard** object that exposes a handy event which fires whenever the clipboard contents are changed. Whenever the clipboard contains an image with recognizable text, it is automatically processed.

Another novelty worth noting is that we don't need an **Image** control for this. The Clipboard image is painted directly on the background of the **Canvas** control using an **ImageBrush** object. Also the Canvas control acts as a container for any number of children controls like the bounding rectangles for recognized words and as such we can assign individual tooltips for each rectangle. Much of the trigonometry required to draw angled rectangles is abstracted away by a **RotateTransform** object that calculates everything automatically.

Finally the **ScrollViewer** control provides hardware-accelerated Zoom In/Out capabilities as well as scrolling in both directions.
