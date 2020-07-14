"""
This example shows how to subscribe to Excel events from Python
using win32com.client.DispatchWithEvents.

This example requires the pywin32 package which can be installed using
>> pip install pywin32
"""
from win32com.client import Dispatch, DispatchWithEvents
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


xl = Dispatch("Excel.Application")
sheet = xl.ActiveSheet
sheet_with_events = DispatchWithEvents(sheet, WorksheetEventHandler)


# Process Windows messages periodically
# NOTE: This isn't necessary if running in Excel using PyXLL.
import pythoncom
import time

while True:
    pythoncom.PumpWaitingMessages()
    time.sleep(0.1)

