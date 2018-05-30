# ShopGenerator
Random shop generator for DnD 5e, complete with Tkinter GUI and an option to output to a .xlsx file.

## Getting Started
Simply download the ShopGenerator.exe file for Windows; for Unix users, you'll have to download all the source
code and its dependencies and run it via shell script for the moment. This means you will have to install
[Python](https://www.python.org/downloads/), [Openpyxl](https://pypi.python.org/pypi/openpyxl), and possibly [Tkinter](https://www.python.org/download/mac/tcltk/), depending on your distribution.

## Deploymnet

Run the .exe file, and a terminal window should appear, do not close it as this will end the program. After about 10 seconds,
the main GUI for the program should appear, and while it's somewhat user-proof, I haven't tested exhaustively.
If the checkbox for 'Make .xlsx file' is not selected, it will only echo the shop to the terminal, making it
useful for creating shops on the fly. If a terminal window is not open, nothing will happen and you must export
the shop to a .xlsx file. These files are exported to a folder called 'shops' in the directory where the selected data file
was found. The .xlsx that data is read from can have at most one row for headers and must be formatted 
in same way as the sample file, including all the data being on a sheet labeled 'Sheet1', and the file extension must be
'.xlsx', not '.xls'. If no entry is given to field marked 'Number to Make' in the GUI, it will default to 1; if a decimal is given, it will be floored to the nearest integer.
