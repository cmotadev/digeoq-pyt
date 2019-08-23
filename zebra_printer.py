#
# https://github.com/mhammond/pywin32/releases/download/b224/pywin32-224.win32-py2.7.exe
#
import win32print
import io

printer_name = None

for flags, descr, name, comments in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
    if name.find("ZDesigner") >= 0:
        printer_name = name

if printer_name:
    hPrinter = win32print.OpenPrinter(printer_name)

    with io.open("teste.zpl", "r", encoding="utf-8") as raw_data:
        p = "".join(raw_data.readlines())

        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("ZPLII data from ArcMap", None, "RAW"))

            try:
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, p)
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
