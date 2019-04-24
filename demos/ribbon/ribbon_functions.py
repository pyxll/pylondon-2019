"""
Functions called from the Excel Ribbon
"""
from pyxll import get_config
import webbrowser
import os


def attach_to_pycharm(*args):
    """Attach the the PyCharm remote debugger"""
    import pydevd

    # disconnect if already connected
    pydevd.stoptrace()

    # (re)connect
    pydevd.settrace('localhost',
                    suspend=False,
                    port=5000,
                    stdoutToServer=True,
                    stderrToServer=True,
                    overwrite_prev_trace=True)


def open_logfile(*args):
    """Open the PyXLL Log file"""
    config = get_config()
    if config.has_option("LOG", "path") and config.has_option("LOG", "file"):
        path = os.path.join(config.get("LOG", "path"), config.get("LOG", "file"))
        webbrowser.open("file://%s" % path)

