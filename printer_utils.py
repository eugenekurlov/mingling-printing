import time
import os

import win32api
import win32print


class PrinterManager:
    def __init__(self, printer_name=None):
        """
        Initialize the PrinterManager with an optional printer name.
        Args:
            printer_name (str): The name of the printer. If not provided, defaults to the system default printer.
        """
        self.printer_handler = None
        self.printer_name = printer_name

    @staticmethod
    def list_printers():
        """
        Retrieves a list of available local printer names on the system.

        Returns:
        - list of str: A list containing the names of all locally available printers.
        """
        printers = []

        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):  # tuple of printer info
            printers.append(printer[2])  # 2 is index of printer name in the tuple
        return printers

    @staticmethod
    def get_default_printer_name():
        default_printer = win32print.GetDefaultPrinter()
        return default_printer

    def get_or_open_printer_handler(self):
        """Get or open a printer handler."""
        if not self.printer_handler:
            if not self.printer_name:
                self.printer_name = PrinterManager.get_default_printer_name()
                print(self.printer_name)
            print_defaults = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            self.printer_handler = win32print.OpenPrinter(self.printer_name, print_defaults)
        return self.printer_handler

    def is_file_in_printer_queue(self, file_name):
        jobs = win32print.EnumJobs(self.printer_handler, 0, -1, 1)
        return any(file_name in job["pDocument"] for job in jobs)

    def wait_for_file_in_queue(self, file_path, is_in_printer_spooler=False):
        file_name = os.path.basename(file_path)
        if is_in_printer_spooler:
            while self.is_file_in_printer_queue(file_name):
                print("Printer queue is not empty. Waiting...")
                time.sleep(1)
            else:
                print("Printer queue is now empty.")
            return None
        else:
            while not self.is_file_in_printer_queue(file_name):
                print("File hasn't gone to printer yet. Waiting...")
                time.sleep(0.1)
            else:
                print("File just arrived to printer.")
            return None

    def print_pdf(self, buffer_or_path, copies=1, orientation=1, duplex=True, flip_side="long", sort_copies=False):
        # Open printer and configure duplex
        self.get_or_open_printer_handler()
        printer_settings = win32print.GetPrinter(self.printer_handler, 2)
        devmode = printer_settings["pDevMode"]
        devmode.Copies = copies
        devmode.Orientation = orientation
        devmode.Duplex = 2 if duplex and flip_side == "long" else 3 if duplex and flip_side == "short" else 1
        # Set collate if sorting is required and copies are more than 1
        devmode.Collate = 1 if sort_copies and copies > 1 else 0
        time.sleep(1)

        win32print.SetPrinter(self.printer_handler, 2, printer_settings, 0)

        try:
            win32api.ShellExecute(0, "print", buffer_or_path, None, ".", 0)
            self.wait_for_file_in_queue(buffer_or_path, is_in_printer_spooler=False)
            print("checking is over")
            self.wait_for_file_in_queue(buffer_or_path, is_in_printer_spooler=True)
            print("ready to delete the file")

        finally:
            print("Document sent to printer successfully.")
            devmode.Duplex = 1  # 1 for simplex (no duplex), 2 for Long-edge, 3 for Short-edge,
            devmode.Orientation = 1  # 1 for Portrait, 2 for Landscape
            devmode.Copies = 1
            devmode.Collate = 0
            win32print.SetPrinter(self.printer_handler, 2, printer_settings, 0)
            time.sleep(2)
            win32print.ClosePrinter(self.printer_handler)
            self.printer_handler = None


if __name__ == "__main__":
    printer_manager = PrinterManager()
    file_path = "file1.pdf"
    try:
        printer_manager.print_pdf(file_path, copies=2, orientation=2, duplex=True, flip_side="long", sort_copies=True)
    except Exception as e:
        print(e)
