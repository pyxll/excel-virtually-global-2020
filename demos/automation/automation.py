"""
This example shows how to automate Excel using Excel's COM
interface and win32com.

This example requires the pywin32 package which can be installed using
>> pip install pywin32
"""
from win32com.client import Dispatch
import random

# Get the Excel Application COM object
xl = Dispatch("Excel.Application")

# Get the current active sheet
sheet = xl.ActiveSheet

# Clear all background colours
sheet.Cells.Interior.ColorIndex = 0

# Set the selected cells to random colours
cells = xl.Selection

for row in range(1, cells.Rows.Count+1):
    for col in range(1, cells.Columns.Count+1):
        cell = cells.Item(row, col)

        red = random.randint(0, 255)
        green = random.randint(0, 255)
        blue = random.randint(0, 255)
        color = red | (green << 8) | (blue << 16)
        cell.Interior.Color = color
