"""
PyXLL Examples: Real time data

As well as returning static values from functions, PyXLL functions
can return special 'RTD' instances that can notify Excel of
updates to their value.

This could be used for any real time data feed, such as live
prices or the status of a service.
"""
from pyxll import RTD, xl_func
import asyncio
import logging

_log = logging.getLogger(__name__)


class ExampleRTD(RTD):

    def __init__(self, initial_value=0):
        super(ExampleRTD, self).__init__(value=initial_value)
        self.__running = True

    async def connect(self):
        # Called when Excel connects to this RTD instance, which occurs
        # shortly after an Excel function has returned an RTD object.
        _log.info("ExampleRTD Connected")

        while self.__running:
            await asyncio.sleep(1)

            # Setting self.value updates Excel
            self.value += 1

    async def disconnect(self):
       # Called when Excel no longer needs the RTD instance. This is
       # usually because there are no longer any cells that need it
       # or because Excel is shutting down.
       _log.info("ExampleRTD Disconnected")
       self.__running = False


@xl_func("int: rtd<int>")
def rtd_example(initial_value=0):
    return ExampleRTD(initial_value)

