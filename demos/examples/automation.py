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



from functools import partial


class EventHandlerMetaClass(type):
    """
    A meta class for event handlers that don't repsond to all events.
    Without this an error would be raised by win32com when it tries
    to call an event handler method that isn't defined by the event
    handler instance.
    """
    @staticmethod
    def null_event_handler(event, *args, **kwargs):
        print(f"Unhandled event '{event}'")
        return None

    def __new__(mcs, name, bases, dict):
        # Construct the new class.
        cls = type.__new__(mcs, name, bases, dict)

        # Create dummy methods for any missing event handlers.
        cls._dispid_to_func_ = getattr(cls, "_dispid_to_func_", {})
        for dispid, name in cls._dispid_to_func_.items():
            func = getattr(cls, name, None)
            if func is None:
                setattr(cls, name, partial(EventHandlerMetaClass.null_event_handler, name))
        return cls


class WorksheetEventHandler(metaclass=EventHandlerMetaClass):

    def OnSelectionChange(self, target):
        print("Selection changed: " + self.Application.Selection.GetAddress())


from win32com.client import DispatchWithEvents

sheet = xl.ActiveSheet
sheet_with_events = DispatchWithEvents(sheet, WorksheetEventHandler)


# Process Windows messages periodically
import pythoncom
import time

while True:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.1)

