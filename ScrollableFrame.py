from tkinter import *
from tkinter import ttk
import tkinter
from tkinter import font


class ScrollableFrame(ttk.Frame):
    """A frame with x and/or y scrollbars and a built-in dark mode."""
    def __init__(self, master: tkinter.Misc, dark_mode: bool = False, dims: str = "xy",
                 location: list = (0, 0), span: list = (1, 1), padding: list = ((0, 0), (0, 0)), **kwargs):
        """
        Initialize scrollable frame.

        Keyword arguments will be passed to frame that contains user-specified child widgets.

        :param master: Tkinter container in which to place scrollable frame
        :param dark_mode: Flag to specify whether the frame should be created in dark mode
        :param dims: A string to indicate which dimensions (x and/or y) should have a scrollbar
        :param location: The (x, y) grid coordinates of the scrollable frame in the parent container
        :param padding: The (x, y) padding of the scrollable frame in the parent container
        :param span: The (row span, column span) values to use for the container
        """

        # Passed parameters
        self.win = master
        self.dark_mode = dark_mode
        self.x_pos = location[0]
        self.y_pos = location[1]
        self.row_span = max(span[0], 1)
        self.column_span = max(span[1], 1)

        # List of headers
        self.header_list = []

        # Define main containers
        self.container = ttk.Frame(self.win)  # Frame to hold scrollbars and canvas
        self.canvas = Canvas(self.container, **kwargs)  # Scrollable component
        super().__init__(self.canvas)  # Frame for user-specified child widgets

        # Define scroll bars based on user specification
        if "y" in dims:
            self.y_scrollbar = ttk.Scrollbar(self.container, orient="vertical", command=self.canvas.yview)
            self.y_scrollbar.grid(row=0, column=1, rowspan=2, sticky="ns")
            self.canvas.configure(yscrollcommand=self.y_scrollbar.set)
        else:
            self.y_scrollbar = None

        if "x" in dims:
            self.x_scrollbar = ttk.Scrollbar(self.container, orient="horizontal", command=self.canvas.xview)
            self.x_scrollbar.grid(row=1, column=0, sticky="ew")
            self.canvas.configure(xscrollcommand=self.x_scrollbar.set)
        else:
            self.x_scrollbar = None

        # Set backgrounds to black if in dark mode
        if self.dark_mode:
            ttk.Style().configure("TFrame", background="black")
            self.canvas.configure(background="black")

        # Set up canvas
        self.canvas.create_window((0, 0), window=self, anchor="nw")
        self.canvas.grid(row=0, column=0, sticky="nsew")

        # Set up widgets
        self.win.rowconfigure(self.y_pos, weight=1)
        self.win.columnconfigure(self.x_pos, weight=1)

        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)
        self.container.grid(row=self.y_pos, column=self.x_pos, rowspan=self.row_span, columnspan=self.column_span,
                            padx=padding[0], pady=padding[1], sticky="nsew")

        # Bind configure
        self.bind("<Configure>", self._configure_interior)

    def _configure_interior(self, event: Event) -> None:
        """Resize the canvas and change canvas position if new widget added or if scroll event is called."""
        self.canvas.config(scrollregion=(0, 0, event.width, event.height))
        self.canvas.config(width=event.width)
        self.canvas.configure(height=event.height)

    def add_header(self, column: int = -1, text: str = "", style: str = "", **kwargs) -> None:
        """
        Add a hidden header label to the scrollable frame for proper column spacing.

        If column is not specified, it will be automatically determined based on the number of header labels present.

        :param column: Column position in which to add header label
        :param text: Text of header label
        :param style: Style name for header label
        """

        # Get style of label
        if style == "":
            font_info = ttk.Style.lookup(ttk.Style(), "TLabel", option="font")
        else:
            font_info = ttk.Style.lookup(ttk.Style(), style, option="font")

        # Determine font information
        if font_info == "TkDefaultFont":
            font_info = font.nametofont('TkTextFont').actual()
            font_type = font_info["family"]
            font_size = font_info["size"]
        else:
            font_info = font_info.replace("{", "")
            font_info = font_info.replace("}", "")
            *font_type, font_size = font_info.split(" ")
            font_size = font_size.strip()

        # Set up Header.TLabel style
        if ttk.Style().lookup("TFrame", option="background") == "black":
            ttk.Style().configure("Header.TLabel", foreground="black", background="black",
                                  font=(font_type, font_size))
        else:
            foreground = ttk.Style().lookup("TFrame", option="background")
            ttk.Style().configure("Header.TLabel", foreground=foreground, font=(font_type, font_size))

        # Determine number of rows to span (based on newline characters in "text")
        rowspan = len(text.split("\n"))

        # Determine column position if not specified
        if column == -1:
            column = len(self.header_list)

        # Delete row span argument if necessary
        if "rowspan" in kwargs.keys():
            kwargs.pop("rowspan")

        header = ttk.Label(self, text=text, style="Header.TLabel")
        header.grid(row=0, column=column, rowspan=rowspan, **kwargs)
        header.lower()

        self.header_list.append(header)
