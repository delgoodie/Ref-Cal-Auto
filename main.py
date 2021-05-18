from os.path import exists as os_path_exists
from os import remove as os_remove
from os import listdir as os_listdir
from os import system as os_system
from os import _exit as os__exit
from sys import exc_info as sys_exc_info
from re import search as re_search
from re import match as re_match
from shutil import copyfile as shutil_copyfile
import PySimpleGUI as sg
from docx import Document as docx_document
from docx.shared import Inches
import matplotlib.pyplot as plt
from docx2pdf import convert as docx2pdf_convert
from math import ceil, floor, log10
from threading import Timer
from time import perf_counter

"""
OVERVIEW:

This program is designed to automate the Reflectance Calibration Task in the RefLab.
Once a scan is taken, the scanning device creates a folder with data from the scan, found in /Equation1.Sample.Cycle1.Equation1.csv, and other intermediate csv files
Then five steps are required of the user, which are automated in this script:

* Correcting the raw data (achieved by copying data from /Equation1.Sample.Cycle1.Equation1.csv to \GrayReflectCalA.xls which functional corrects data)
* Creating a txt file which holds corrected data (must be named (last four of serial number)-(model name))
* Creating a certificate word doc and transfering metadata (model, serial number, date, ...), corrected data (from \GrayReflectCalA.xls), and a graph of the corrected data (from \GrayReflectCalA.xls)
* Creating a PDF cert of the word doc
* Copying the PDF and txt file to a USB stick

Also note that an important part of the automation process is locating the Stray Light Scan performed on the same day as the reflectance scan, which is used in correcting the raw data

CODE NOTES:

* When compiling this script, copy the /User Data/ folder into the same directory as main.exe

* methods are imported and renamed with underscores so as to keep naming clear and concise will still importing the mininmum number of methods

* Execute function is seperated into many parts so that each part can be wrapped with @debug.

* the debug function is a wrapper which logs the time taken by a function, handels errors, and logs status

* window and params were made global since debug and other functions require them (yeah, it had to be)

* this script depends on a relative folder /User Data which must contain rr.txt, clients.txt, config.txt, and template.docx

"""

# region SETUP

# checks for local folder User Data
if not os_path_exists("User Data\\rr.txt"):
    raise Exception("User Data\\rr.txt Does Not Exist")
if not os_path_exists("User Data\\config.txt"):
    raise Exception("User Data\\config.txt Does Not Exist")
if not os_path_exists("User Data\\template.docx"):
    raise Exception("User Data\\template.docx Does Not Exist")

# Reads config data for constant paths
config = open("User Data\\config.txt").read()

TARGET_DIR = re_search("TARGET_PATH\\s+=\\s+(.+)($|\\n)", config).group(1).strip()
STRAY_LIGHT_DIR = re_search("STRAY_LIGHT_PATH\\s+=\\s+(.+)($|\\n)", config).group(1).strip()
USB_PATH = re_search("USB_PATH\\s*=\\s*(.+)($|\\n)", config).group(1).strip()

del config

# reads client data to generate clients dropdown and populate CLIENTS requirements

cl_int = {}
if os_path_exists("User Data\\clients.txt"):
    last = None
    for l in open("User Data\\clients.txt").readlines():
        if re_match("\\s+.+", l):
            reg = re_search("(\\d+)\\s+(\\d+\\.*\\d*)\\-(\\d+\\.*\\d*)", l)
            if reg:
                cl_int[last][int(reg.group(1))] = (float(reg.group(2)), float(reg.group(3)))
            elif re_match("\\s+flatness", l):
                cl_int[last]["flatness"] = True
        else:
            last = l.strip()
            cl_int[last] = {}
    del last, reg, l
CLIENTS = cl_int
"""
CLIENTS dict follows structure:
{
    (str)"client name": { (client name is displayed in client dropdown)
        (int)wavelength: (tuple)((float)lower_limit, (float)upper_limit),
        ...,
        (optionally)
        (str)"flatness": (bool)True (!this value is ignored)
    },
    ...
}
"""
del cl_int

# deletes are used to clean up global scope

window = 0
params = 0

print("Constants & Globals initialized...")

# endregion

# region HELPERS


class Date:
    def __init__(self, month: int, day: int, year: int):
        self.month = month
        self.day = day
        self.year = year

    def valid(self):
        return self.month >= 1 and self.month <= 12 and self.day > 0 and self.day <= 31 and self.year > 2000 and self.year < 3000

    def __eq__(self, other):
        return type(other) is Date and self.month == other.month and self.day == other.day and self.year == other.year

    def __str__(self):
        return f"{self.month}/{self.day}/{self.year}"


class Parameters:
    def __init__(
        self,
        root_path: str = None,
        geometry: str = None,
        size: int = None,
        material: str = None,
        serial_number: str = None,
        reflectance: int = None,
        requirements: dict = None,
        instrument: str = None,
        date: Date = None,
        stray_light_path: str = None,
    ):
        if len(root_path) > 0 and not root_path[-1] in ["\\", "/"]:
            root_path += "\\"
        if type(reflectance) is str:
            search = re_search("(\\d+)\\%", reflectance)
            if search:
                reflectance = int(search.group(1))
        if type(date) is str:
            date = DateFromString(date)
        if len(stray_light_path) > 0 and not stray_light_path[-1] in ["\\", "/"]:
            stray_light_path += "\\"

        self.root_path = root_path
        self.geometry = geometry
        self.size = size
        self.material = material
        self.serial_number = serial_number
        self.reflectance = reflectance
        self.requirements = requirements
        self.instrument = instrument
        self.date = date
        self.stray_light_path = stray_light_path

        # derived property
        self.model = f"{'SRT' if self.geometry == 'Target' else 'SRS'}-{self.reflectance}-{self.size}"

    def isValid(self):
        if self.root_path and type(self.root_path) is str and os_path_exists(self.root_path):
            if self.material and type(self.material) is str and self.material in ["Spectraflect", "Permaflect"]:
                if self.geometry and type(self.geometry) is str and self.geometry in ["Target", "Puck"]:
                    if self.size and type(self.size) is int:
                        if self.serial_number and type(self.serial_number) is str and len(self.serial_number) > 4:
                            if self.reflectance and type(self.reflectance) is int and self.reflectance > 0 and self.reflectance < 100:
                                if True or self.requirements and type(self.requirements) is dict:
                                    if self.instrument and self.instrument in ["A", "B", "C"]:
                                        if self.date and type(self.date) is Date and self.date.valid():
                                            if self.stray_light_path and type(self.stray_light_path) is str and os_path_exists(self.stray_light_path):
                                                return True
                                            else:
                                                return "Invalid Stray Light Path"
                                        else:
                                            return "Invalid Date"
                                    else:
                                        return "Invalid Instrument"
                                else:
                                    return "Invalid Requirements (or client)"
                            else:
                                return "Invalid Reflectance"
                        else:
                            return "Invalid Serial Number"
                    else:
                        return "Invalid [Puck or Target] Size"
                else:
                    return "Invalid Type"
            else:
                return "Invalid Material"
        else:
            return "Invalid Root Path"


class CSV:
    def __init__(self, _path: str):
        self.path = _path
        file = open(self.path, "r")
        self._data = []
        for l in file.read().split("\n"):
            self._data.append(l.split(","))
        file.close()
        self.modified = False

    def _ParseTablePosition(self, position) -> tuple[int, int]:
        if type(position) is str:
            if ":" in position:
                c, r = position.split(":", 1)
            else:
                match = re_match("([a-zA-Z]+)(\\d+)", position)
                if len(match.regs) > 1:
                    c = match.group(1)
                    r = match.group(2)
            if c and r:
                position = (ord(c.upper()) - 65, int(r) - 1)
            else:
                raise ValueError(f"CSV._ParseTablePosition: c:{c}, r:{r} is invalid position")
        elif type(position) is tuple:
            c, r = position
            if type(c) is int or type(c) is float or re_match("\\d+", c):
                c = int(c)
            elif re_match("[A-Za-z]+", c):
                c = c.upper()
                col = 0
                i = 1
                for char in c:
                    col += (ord(char) - 65) * i
                    i *= 26
                c = col
            r = int(r)
            position = (c, r - 1)
        else:
            raise TypeError(f"CSV._ParseTablePosition: position: {type(position)} is not <class 'str'> or <class 'tuple'>")

        if type(position) is tuple and type(position[0]) is int and type(position[1]) is int:
            return position
        else:
            raise ValueError(f"CSV._ParseTablePosition: c:{c}, r:{r} is invalid position")

    def Read(self, position="*"):
        """
        Read from csv table cell(s)

        position -- cell position: '*', (0, 34) | ('A', 34) | ('A', '34') | 'A34' | 'A:34'
        if position is '*' then a 2D array of all cols & rows will be returned

        return -- data at position
        """
        if position == "*":
            return self._data
        else:
            parsed_pos = self._ParseTablePosition(position)
            if parsed_pos[1] < len(self._data) and parsed_pos[0] < len(self._data[parsed_pos[0]]):
                return self._data[parsed_pos[1]][parsed_pos[0]]
            else:
                raise IndexError(f"easyio.CSV.Read: position:{parsed_pos} not in cols: {len(self._data[0])}, rows: {len(self._data)}")

    def Write(self, position, data: str = ""):
        self.modified = True
        """
        Write string to single csv table cell

        position -- cell position: (0, 34) | ('A', 34) | ('A', '34') | 'A34' | 'A:34'
        data -- string of new cell contents: '34.213'

        return -- the overwriten data: 'data before Write()'
        """
        parsed_pos = self._ParseTablePosition(position)

        pre = self._data[parsed_pos[1]][parsed_pos[0]]
        self._data[parsed_pos[1]][parsed_pos[0]] = data
        return pre

    def __del__(self):
        if self.modified:
            string = ""
            for c in self._data:
                for r in c:
                    string += r + ","
                string += "\n"

            file = open(self.path, "w")
            file.write(string)
            file.close()


class DOCX:
    def __init__(self, _path):
        self.path = _path
        self.doc = docx_document(_path)

    def _docxOccurences(self, variable: str) -> None:
        occurences = []
        for p in self.doc.paragraphs:
            if re_search(f"<{variable}>", p.text):
                occurences.append(p)
        for t in self.doc.tables:
            for c in t._cells:
                if re_search(f"<{variable}>", c.text):
                    occurences.append(c)
        for s in self.doc.sections:
            for p in s.header.paragraphs:
                if re_search(f"<{variable}>", p.text):
                    occurences.append(p)
            for t in s.header.tables:
                for c in t._cells:
                    if re_search(f"<{variable}>", c.text):
                        occurences.append(c)
        return occurences

    def ReplaceText(self, variable: str, value: str):
        for e in self._docxOccurences(variable):
            e.text = e.text.replace(f"<{variable}>", value)

    def ReplacePicture(self, variable: str, path: str, size):
        for e in self._docxOccurences(variable):
            e.text = ""
            e.add_run().add_picture(path, width=Inches(size[0]), height=Inches(size[1]))

    def Save(self, path=None):
        if path:
            self.doc.save(path)
        else:
            self.doc.save(self.path)


def DateFromString(string: str):
    """
    Retrieves date from string containing a date in the following formats:

    MM/DD/YY(YY)

    MM-DD-YY(YY)

    NOTE: Shorthand and Full years will be converted to full years

    i.e. 21 => 2021, 1985 => 1985

    @return tuple containing (MM, DD, YYYY)
    """
    search = re_search("(\\d+)[\\/\\-](\\d+)[\\/\\-](\\d+)", string)
    if not search or len(search.regs) != 4:
        return -1
    else:
        month = int(search.group(1))
        day = int(search.group(2))
        year = int(search.group(3))
        # check if year is in shorthand, '21, or longhand, 2021 -> convert to longhand
        if log10(year) <= 3:
            year += 2000
    return Date(month, day, year)


def LeftPad(string, length):
    """
    Left Pad a string with ' ' (spaces) so that it takes on the length of @param length
    """
    while len(string) < length:
        string += " "
    return string


def debug(func):
    """
    Debug Wrapper to catch errors, display timestamp, and log status

    use:

    @debug

    def MyFunc
    """
    global window

    def wrap(*args, **kwargs):
        start = perf_counter()
        print(LeftPad(func.__name__, 40), end="")
        window["Log"].update(func.__name__)
        result = False
        try:
            result = func(*args, **kwargs)
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys_exc_info()
            print("\n\n[Exception Raised]")
            print("Send email to:    wdelgiudice@labsphere.com    with the following information:")
            print("* error message (seen below) \n* all form data (model, sn, date, client, etc)\n* anything different about this specific scan")
            print("Error Message:")
            print("TYPE: ", exc_type)
            print("LINE: ", exc_tb.tb_lineno)
            print("ERROR: ", e)
            window["Log"].update("ERROR: GO TO CONSOLE")
            return False
        end = perf_counter()
        print(f"√  {round(1000 * (end - start))} ms")
        window["Log"].update(f"{func.__name__}   √")
        return result

    return wrap


# endregion

# region EXECUTION STEPS


@debug
def GetStrayLightPaths(date: Date) -> list[str]:
    """
    Returns array of relative paths from STRAY_LIGHT_DIR to stray light folders written at dates that match the @param date

    @param date - must be tuple in (DD:int, MM:int, YYYY:int) format

    @return - List[str] of relative paths from STRAY_LIGHT_DIR
    """
    result = []
    for dir in os_listdir(STRAY_LIGHT_DIR):
        dirDate = DateFromString(dir)
        if dirDate != -1 and dirDate == date:
            result.append(dir)
    return result


@debug
def Get_rr():
    """
    Reads Rr data from \\User Data\\rr.txt
    """
    return [float(l) for l in open("User Data\\rr.txt").readlines()]


@debug
def CorrectData(raw: DOCX, strayLight: DOCX, Rr: list[float]) -> list[float]:
    """
    Calculates Corrected data from raw data, stray light data, and Rr
    """
    # MS% in column B
    Ms = {}
    i = 2
    while len(raw.Read(("A", i))) > 0:
        Ms[int(float(raw.Read(("A", i))))] = float(raw.Read(("B", i)))
        i += 1

    Mh = {}
    i = 2
    while len(strayLight.Read(("A", i))) > 0:
        s_l = float(strayLight.Read(("B", i)))
        s_l = abs(s_l)
        s_l = min(s_l, 100)
        Mh[int(float(strayLight.Read(("A", i))))] = s_l
        i += 1

    c_d = {}
    for w in Mh:
        if w >= 250 and w <= 2500:
            c_d[w] = round((Ms[w] * 0.01 - Mh[w] * 0.01) * (1 / (1 - Mh[w] * 0.01)) * Rr[w - 250], 4)
            if c_d[w] < 0:
                print(f"negative c_d: {c_d[w]} @ {w}")

    return c_d


@debug
def TestClientRequirements(corrected_data):
    """
    Tests Corrected Data against Client Requirements from selected Client

    @return - 0 if data passed tests, str if data failed for error msg
    """
    global params
    error = 0
    for w in params.requirements:
        if type(w) is int:
            if corrected_data[w] < params.requirements[w][0] or corrected_data[w] > params.requirements[w][1]:
                error = f"Corrected Data did not pass client requirements ({corrected_data[w]} @ {w} did not meet {params.requirements[w][0]} <= reflectance <= {params.requirements[w][1]} @ {w})"
        elif w == "flatness":
            dif = (corrected_data[1500] - corrected_data[1000]) * 100
            if dif < 1 or dif > 2:
                error = f"Corrected Data did not pass client requirements ({dif} did not meet flatness requirements 1 <= ([ref @ 1500] - [ref @ 1000]) * 100 <= 2)"
    return error


@debug
def SaveTextFile(corrected_data):
    """
    Saves corrected data as text file under (last four of sn)-(model name).txt
    """
    # print(f"Generating {params['serial number'][-4:len(params['serial number'])]}-{params['model']}.txt file    ")
    txt = open(f"{params.root_path}{params.serial_number[-4:len(params.serial_number)]}-{params.model}.txt", "w")
    stringdata = [f"{w}\t{corrected_data[w]}\n" for w in corrected_data]
    stringdata.insert(0, f"{params.serial_number}\nThis data is for reference only\n")
    txt.write("".join(stringdata))
    del stringdata
    return True


@debug
def WriteWordMeta(doc):
    """
    Writes metadata to word docx template

    * sn
    * date
    * model
    * instrument (A, B, C)
    """
    doc.ReplaceText("sn", params.serial_number)
    doc.ReplaceText("DATE", str(params.date))
    doc.ReplaceText("model", params.model)
    doc.ReplaceText("isA", "X" if params.instrument == "A" else "")
    doc.ReplaceText("isB", "X" if params.instrument == "B" else "")
    doc.ReplaceText("isC", "X" if params.instrument == "C" else "")
    return True


@debug
def WriteWordData(doc: DOCX, corrected_data: dict):
    """
    Writes corrected reflectance data to word docx cert in in table

    Rounds reflectance to nearest value in uncertainty chart and uses values for sig fig rounding
    """
    uncertainty_table = doc.doc.tables[2]._cells
    uncertainty = {}
    offset = -1
    uncertainty_ref = 0
    UNCERTAINTY_COLS = 9
    for i in range(1, UNCERTAINTY_COLS):
        col_ref = int(re_search("(\\d+)%", uncertainty_table[i].text).group(1))
        if params.reflectance == col_ref or not uncertainty_ref or abs(params.reflectance - col_ref) < abs(params.reflectance - uncertainty_ref):
            uncertainty_ref = col_ref
            offset = i

    for i in range(UNCERTAINTY_COLS + offset, len(uncertainty_table), UNCERTAINTY_COLS):
        w = int(uncertainty_table[i - offset].text)
        u = 0
        non_zero = False
        uncert = uncertainty_table[i].text
        for c in uncertainty_table[i].text:
            if c in "123456789":
                u += 1
                non_zero = True
            elif c == "0" and non_zero:
                u += 1
        uncertainty[w] = u

    for w in range(250, 2510, 50):
        if w in corrected_data:
            v = corrected_data[w]
            u = uncertainty[w]
            if v and u:
                # rounding to sig figs based off uncertainty table
                v = round(v, u - int(floor(log10(abs(v)))))
                v = str(v)
                while len(v) < u + 3:
                    v += "0"
                doc.ReplaceText(f"w{str(int(w / 10))}", v)
            else:
                print(f"No corrected_data at wavelength: {w}")
    return True


@debug
def WriteWordGraph(doc, corrected_data):
    """
    Writes graph to word docx cert

    * Creates graph with matlib from corrected data
    * Saves graph to temp.png
    * Writes image to docx cert
    * Deletes temp.png
    """
    plt_x = [w for w in corrected_data]
    plt_y = [corrected_data[w] for w in corrected_data]
    plt.plot(plt_x, plt_y, color="black")
    # plt.title("Graph I: 8°/Hemispherical Spectral Reflectance")
    plt.ylabel("Reflectance Factor")
    plt.xlabel("Wavelength (nm)")
    plt.xticks([i for i in range(250, 2501, 250)])
    plt.axis([250, 2500, 0, ceil(max(plt_y) * 4) * 0.25])
    plt.savefig("temp.png")
    doc.ReplacePicture("graph", "temp.png", (7, 5.5))
    os_remove("temp.png")
    return True


@debug
def SaveWord(doc):
    """
    Saves word doc cert

    This function exists simply to wrap DOCX.Save() method so the debug wrapper can be used
    """
    doc.Save(f"{params.root_path}DM-01400-010Rev04 {'99' if params.reflectance == '99%' else 'Gray'} cal cert non NVLAP.docx")
    return True


@debug
def SavePdf() -> None:
    """
    Copies word doc cert as pdf

    This function exists simply to wrap docx2pdf_covert() method so the debug wrapper can be used
    """
    print("")
    docx2pdf_convert(
        f"{params.root_path}DM-01400-010Rev04 {'99' if params.reflectance == '99%' else 'Gray'} cal cert non NVLAP.docx",
        f"{params.root_path}{params.serial_number}.pdf",
    )
    return True


@debug
def CopyToUsb():
    """
    Copies raw data txt file and final cert pdf file to USB path specified in User Data\\config.txt
    """
    if os_path_exists(USB_PATH):
        docxName = f"DM-01400-010Rev04 {'99' if params.reflectance == '99%' else 'Gray'} cal cert non NVLAP.docx"
        shutil_copyfile(f"{params.root_path}{docxName}", f"{USB_PATH}{docxName}")
        shutil_copyfile(f"{params.root_path}{params.serial_number}.pdf", f"{USB_PATH}{params.serial_number}.pdf")
    return True


# endregion


def Execute() -> bool:
    """
    Execute function sequentially calls steps and exits if any step fails:

    Execute is done in many steps so that each step can be wrapped with @debug for easy debugging if any point fails

    * Get Rr
    * Corrected Data
    * Test Client Requirements
    * Save Text File
    * Write Word Meta
    * Write Word Data
    * Write Word Graph
    * Save Word
    * Save Pdf
    * Copy To Usb
    """
    global params
    global window

    raw = CSV(f"{params.root_path}Equation1.Sample.Cycle1.Equation1.csv")
    strayLight = CSV(f"{params.stray_light_path}Equation1.Sample.Cycle1.Equation1.csv")
    doc = DOCX("User Data\\template.docx")

    rr = Get_rr()
    if not rr:
        return False

    corrected_data = CorrectData(raw, strayLight, rr)
    if not corrected_data:
        return False

    msg = TestClientRequirements(corrected_data)
    if type(msg) is str:
        window["Log"].update(msg)
        print(msg)
        return False

    if not SaveTextFile(corrected_data):
        return False

    if not WriteWordMeta(doc):
        return False

    if not WriteWordData(doc, corrected_data):
        return False

    if not WriteWordGraph(doc, corrected_data):
        return False

    if not SaveWord(doc):
        return False

    if not SavePdf():
        return False

    if not CopyToUsb():
        return False

    return True


def AsyncExecute():
    """
    This function exists so the Execute Function can occur on a seperate thread, while the GUI continues to receive updates
    """
    global window
    global params
    result = Execute()
    if result:
        window["Log"].update("Finished: SUCCESS")
        os_system(f'start "" "{params.root_path}"')
        if os_path_exists(USB_PATH):
            os_system(f'start "" "{USB_PATH}"')
        window.close()
        os__exit(0)


def main() -> None:
    """
    Main function:

    * Initializes GUI
    * Handles Events
    """
    global window
    global params

    layout = [
        # 0
        [sg.Text("Root Folder", font="30px")],
        # 1
        [sg.FolderBrowse(), sg.Input("(must have Equation1.Sample.Cycle1.Equation1.csv)", size=(82, 1), readonly=True)],
        # 2
        [sg.Text("")],
        # 3
        [sg.Text("Info", font="30px")],
        # 4
        [
            sg.Text("Geometry"),
            sg.DropDown(["Target", "Puck"], "Target", key="Geometry", size=(9, 1), readonly=True, enable_events=True),
            sg.Text("Material"),
            sg.DropDown(["Spectraflect", "Permaflect"], "Spectraflect", key="Material", size=(13, 1), readonly=True, enable_events=True),
            sg.Text("Target Length", size=(11, 0), key="Size Name"),
            sg.Input("", size=(6, 1), key="Size"),
            sg.Text("Reflectance"),
            sg.DropDown(["2%", "5%", "10%", "20%", "40%", "60%", "80%", "99%"], "99%", size=(7, 1), key="Reflectance", readonly=False),
        ],
        # 5
        [
            sg.Text("Serial Number", size=(10, 1)),
            sg.Input("", size=(20, 1), key="Serial Number"),
            sg.Text("Client", pad=((0, 0), 0)),
            sg.DropDown([c for c in CLIENTS], "no client selected", size=(49, 1), key="Client", readonly=True),
        ],
        # 6
        [sg.Text("")],
        # 7
        [sg.Text("Stray Light Scan", font="30px")],
        # 8
        [
            sg.Text("Instrument"),
            sg.DropDown(["A", "B", "C"], "A", size=(5, 1), key="Instrument", readonly=True),
            sg.CalendarButton(
                "Select Date",
                target=(8, 3),
                format="%m/%d/%Y",
                enable_events=True,
            ),
            sg.Input(key="Date", size=(10, 1), enable_events=True, readonly=True),
            sg.DropDown([], "No date selected", size=(45, 1), enable_events=True, key="Stray Light Dropdown", readonly=True),
        ],
        # 9
        [
            sg.FolderBrowse(button_text="Manual Browse", target=(9, 0)),
            sg.Input("Stray Light Path (not selected)", key="Stray Light Path", size=(75, 1), readonly=True),
        ],
        # 10
        [sg.Text("")],
        # 11
        [sg.Button("Execute", size=(20, 1), font="30px", pad=(210, 0))],
        # 12
        [sg.Text("Log", font="30px")],
        # 13
        [sg.Input(key="Log", size=(91, 5), readonly=True, enable_events=True)],
    ]

    window = sg.Window(title="Ref Cal Auto", layout=layout, margins=(0, 20))

    while True:
        event, values = window.read()

        if event == "Geometry":
            window["Size Name"].update("Target Size" if values["Geometry"] == "Target" else "Puck Diameter")
        elif event == "Material":
            0
        elif event == "Date":
            date = DateFromString(values["Date"])
            if date != -1:
                slps = GetStrayLightPaths(date)
                if len(slps) == 0:
                    window["Stray Light Dropdown"].update(values=[])
                else:
                    window["Stray Light Dropdown"].update(values=slps)
        elif event == "Stray Light Dropdown":
            window["Stray Light Path"].update(f"{STRAY_LIGHT_DIR}{values['Stray Light Dropdown']}")
        elif event == "Execute":
            params = Parameters(
                values["Browse"],
                values["Geometry"],
                int(values["Size"]),
                values["Material"],
                values["Serial Number"],
                values["Reflectance"],
                CLIENTS[values["Client"]] if values["Client"] in CLIENTS else {},
                values["Instrument"],
                values["Date"],
                values["Stray Light Path"],
            )
            error = params.isValid()
            if not type(error) is str:
                window["Log"].update("Executing...")
                Timer(1.0, AsyncExecute).start()
            else:
                window["Log"].update(error)
                print("Execution Conditions not met:     ", error, "\n")
        elif event == sg.WIN_CLOSED:
            break
    window.close()


# !ENTRY POINT
main()

# !TESTING
# Execute(
#     {
#         "path": "C:\\Users\\wdelgiudice\\Downloads\\18%PF-1020-4436 - Copy\\",
#         "reflectance": "99%",
#         "serial number": "PF-0921-4398",
#         "date": DateFromString("1/6/2021"),
#         "instrument": "B",
#         "stray light path": "\\\\lssvr-fs01\\Reflectance Lab\\Reflectance Calibrations\\stray light Summary.xls_files\\Stray Light 4-7-2021 C\\",
#         "client": {1000: (0.1, 0.21), 1500: (0.1, 0.21)},
#     }
# )
