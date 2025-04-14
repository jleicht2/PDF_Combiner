PDF_Combiner is a tool to combine multiple PDF files into a single output file. For each file, the program will default to adding all pages, but individual page(s) of the file can be selected. These pages can also be duplicated or placed "out of order" in the output file.

### **Requirements:**
<li>Python v. 3.8 or Greater <br /> </li>
<li>PyPDF v. 5.1.0 or Greater: This can be automatically installed by the script (using pip and the command line) if necessary <br /></li>
<li>pywin32 v. 310 or Greater: This is only needed to automatically install shortcuts to the script and is only available on Windows machines. It can also be automatically installed if necessary. <br /></li>

### **Notes:**
PDF_Combiner was built and tested on a Windows machine. The core functionality should be cross-platform, but some non-critical features (such as automatic shortcut installation) may not be available on all machines. <br />
Additionally, buttons may not appear in platform-specific format as an extension of the ttk.Label class was created to allow for "dark mode"-style buttons using the default ttk/Tkinter theme.
