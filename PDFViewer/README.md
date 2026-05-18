### This is a PDF Viewer project. The user interface is created entirely as a XAML Island on an empty form, while the necessary plumbing is wired in code-behind. This serves as a great example of the features and flexibility of XAML code for creating a fluid and responsive UI in VB6.

Navigation is provided by a **MenuBar** at the top which includes a **File** menu for opening PDF files using the _FileOpenPicker_ class, a **Rotation** menu to rotate pages at different angles and a **Background** color menu for setting various background colors for the PDF pages. The **Background** menu template is also stylized in XAML code to show color swatches as visual cues for the color names.

Since there is no concept of grouping menu items in arrays (as you would in the classic VB6 menu editor) and we don't want a separate click event for each menu item, a single _XamlUICommand_ object is bound to each of them in XAML code with functionality provided by a single event in code-behind.

The main workhorse for displaying PDF pages is a **ListView** control with its _DataTemplate_ edited in XAML code to contain a **Border** and **Image** controls for each item container. One of the main selling points of XAML controls is their ability to **Bind** one or more properties to various data sources providing the ultimate flexibility. Here the ListView's _ItemsSource_ property is bound to a custom implementation of _IBindableVector_, while each container's Image control is bound to a _BitmapSource_ object rendered on-the-fly.

The whole functionality is entirely **asynchronous and event driven** providing silky-smooth performance even when scrolling through hundreds of PDF pages all rendered in real time. Old pages are gradually destroyed while new ones are rendered when scrolling through the document.

<img width="595" height="876" alt="PDFViewer" src="https://github.com/user-attachments/assets/70ed918b-af4e-4a51-8e1c-75b247957eb7" />

The project starts by loading a sample PDF document from a **URI** (an Icelandic dictionary, random choice I found online but great for testing) so it may take a couple of seconds while this large file is buffered. After that you can open other PDF documents from the **File** menu.

Scrolling can be either fine-tuned with the mouse wheel or in bulk chunks by dragging the vertical scroll bar handle within large documents. There's also the possibility to **Zoom In/Out** with the **Ctrl-MouseWheel** shortcut. Zooming is focused on the current position of the mouse pointer. While the page is zoomed in, a horizontal scroll bar at the bottom allows for horizontal scrolling.

The code is heavily commented to help whoever wants to dip their toes in the XAML world.
