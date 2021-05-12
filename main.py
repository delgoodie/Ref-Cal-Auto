from os.path import exists as os_path_exists
from os import remove as os_remove
from os import listdir as os_listdir
from os import system as os_system
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
NOTES:

* methods are imported and renamed with underscores so as to keep naming clear and concise will still importing the mininmum number of methods

* Execute function is seperated into many parts so that each part can be wrapped with @debug.

* the debug function is a wrapper which logs the time taken by a function, handels errors, and logs status

* window and params were made global since debug and other functions require them (yeah, it had to be)

* this script depends on a relative folder /User Data which must contain Rr.txt, Clients.txt, Constants.txt, and Ref-Cal-Cert-Template.docx

"""

constants = open("User Data\\Constants.txt").read()

TARGET_DIR = re_search("TARGET_PATH\\s+=\\s+(.+)\\n", constants).group(1).strip()
STRAY_LIGHT_DIR = re_search("STRAY_LIGHT_PATH\\s+=\\s+(.+)\\n", constants).group(1).strip()
USB_PATH = re_search("USB_PATH\\s+=\\s+(.+)\\n", constants).group(1).strip()

last = None
cl_int = {}
for l in open("User Data\\Clients.txt").readlines():
    if re_match("\\s+.+", l):
        reg = re_search("(\\d+)\\s+(\\d+\\.*\\d*)\\-(\\d+\\.*\\d*)", l)
        if reg:
            cl_int[last][int(reg.group(1))] = (float(reg.group(2)), float(reg.group(3)))
        elif re_match("\\s+flatness", l):
            cl_int[last]["flatness"] = True
    else:
        last = l.strip()
        cl_int[last] = {}
CLIENTS = cl_int
window = 0
params = {}

print("Constants & Globals initialized...")


# region HELPERS


class CSV:
    def __init__(self, _path: str):
        self.path = _path
        file = open(self.path, "r")
        self._data = []
        for l in file.read().split("\n"):
            self._data.append(l.split(","))
        file.close()
        self.modified = False

    def _ParseTablePosition(position) -> tuple[int, int]:
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
        Read from csv table cell[s]

        position -- cell position: '*', (0, 34) | ('A', 34) | ('A', '34') | 'A34' | 'A:34'
        if position == '*' then a 2D array of all cols & rows will be returned

        return -- data at position
        """
        if position == "*":
            return self._data
        else:
            position = self._ParseTablePosition(position)
            if position[1] < len(self._data) and position[0] < len(self._data[position[0]]):
                return self._data[position[1]][position[0]]
            else:
                raise IndexError(f"easyio.CSV.Read: position:{position} not in cols: {len(self._data[0])}, rows: {len(self._data)}")

    def Write(self, position, data: str = ""):
        self.modified = True
        """
        Write string to single csv table cell

        position -- cell position: (0, 34) | ('A', 34) | ('A', '34') | 'A34' | 'A:34'
        data -- string of new cell contents: '34.213'

        return -- the overwriten data: 'data before Write()'
        """
        position = self._ParseTablePosition(position)

        pre = self._data[position[1]][position[0]]
        self._data[position[1]][position[0]] = data
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


def DateFromString(string: str) -> tuple:
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
        return 0
    else:
        month = int(search.group(1))
        day = int(search.group(2))
        year = int(search.group(3))
        # check if year is in shorthand, '21, or longhand, 2021 -> convert to longhand
        if log10(year) > 3:
            return (month, day, year)
        else:
            return (month, day, 2000 + year)


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
            print("\n\n[Exception Raised]")
            print("Copy error message and send to:")
            print("wdelgiudice@labsphere.com\n")
            print("Error Message:")
            print(e)
            window["Log"].update("ERROR: GO TO CONSOLE")
            return False
        end = perf_counter()
        print(f"√  {round(1000*(end - start))} ms")
        window["Log"].update(f"{func.__name__}   √")
        return result

    return wrap


# endregion

# region EXECUTION STEPS


@debug
def GetStrayLightPaths(date) -> list[str]:
    """
    Returns array of relative paths from STRAY_LIGHT_DIR to stray light folders written at dates that match the @param date

    @param date - must be tuple in (DD:int, MM:int, YYYY:int) format

    @return - List[str] of relative paths from STRAY_LIGHT_DIR
    """
    result = []
    for dir in os_listdir(STRAY_LIGHT_DIR):
        dirDate = DateFromString(dir)
        if dirDate and dirDate[0] == date[0] and dirDate[1] == date[1] and dirDate[2] == date[2]:
            result.append(dir)
    return result


@debug
def TestExecuteConditions(values):
    """
    Tests for conditions before executing script

    * Makes sure all fields in the form have been filled out
    * Verifies paths of stray light scan and root directory
    * Verifies existance of dependant local files

    @return - 0 for Tests Passed, str for error message
    """
    result = 0
    if not values["Browse"]:
        result = "No Folder Selected"
    elif not values["Model"]:
        result = "No Model Specified"
    elif not values["Serial Number"]:
        result = "No Serial Number Specified"
    elif len(values["Serial Number"]) < 5:
        result = "Serial Number Invalid (too short)"
    elif not values["Nominal Reflectance"]:
        result = "No Nominal Reflectance Selected"
    elif not window["Date"]:
        result = "No Date Selected"
    elif not values["Client"]:
        result = "No Client Selected"
    elif not values["Instrument"]:
        result = "No Instrument Selected"
    elif values["Stray Light Path"] == "Stray Light Path (not selected)":
        result = "No Stray Light Path Specified"
    elif not os_path_exists(f"{values['Browse']}\\Equation1.Sample.Cycle1.Equation1.csv"):
        result = "Invalid Root Folder: Equation1.Sample.Cycle1.Equation1.csv DOES NOT EXIST"
    elif not os_path_exists(f"{values['Stray Light Path']}\\Equation1.Sample.Cycle1.Equation1.csv"):
        result = "No Stray Light Scan found"
    elif not os_path_exists("User Data\\Ref-Cal-Cert-Template.docx"):
        result = "No Ref-Cal-Cert-Template.docx found in script directory"
    return result


@debug
def GetRr():
    """
    Reads Rr data from \\User Data\\Rr.txt
    """
    return [float(l) for l in open("User Data\\Rr.txt").readlines()]


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
        Mh[int(float(strayLight.Read(("A", i))))] = float(strayLight.Read(("B", i)))
        i += 1

    c_d = {}
    for w in Mh:
        if w >= 250 and w <= 2500:
            c_d[w] = ((Ms[w] * 0.01 - Mh[w] * 0.01) / (1 - Mh[w] * 0.01)) * Rr[w - 250]

    return c_d


@debug
def TestClientRequirements(corrected_data):
    """
    Tests Corrected Data against Client Requirements from selected Client

    @return - 0 if data passed tests, str if data failed for error msg
    """
    global params
    error = 0
    for w in params["client"]:
        if type(w) is int:
            if corrected_data[w] < params["client"][w][0] or corrected_data[w] > params["client"][w][1]:
                error = f"Corrected Data did not pass client requirements ({corrected_data[w]} @ {w} did not meet {params['client'][w][0]} <= reflectance <= {params['client'][w][1]} @ {w})"
        elif w == "flatness":
            if corrected_data[1500] - corrected_data[1000] > 2 or corrected_data[1500] - corrected_data[1000] < 1:
                error = f"Corrected Data did not pass client requirements ({corrected_data[1500] - corrected_data[1000]} did not meet flatness requirements 1 <= [ref @ 1500] - [ref @ 1000] <= 2)"
    return error


@debug
def SaveTextFile(corrected_data):
    """
    Saves corrected data as text file under (last four of sn)-(model name).txt
    """
    # print(f"Generating {params['serial number'][-4:len(params['serial number'])]}-{params['model']}.txt file    ")
    txt = open(f"{params['path']}{params['serial number'][-4:len(params['serial number'])]}-{params['model']}.txt", "w")
    stringdata = [f"{w}\t{corrected_data[w]}\n" for w in corrected_data]
    stringdata.insert(0, f"{params['serial number']}\nThis data is for reference only\n")
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
    doc.ReplaceText("sn", params["serial number"])
    doc.ReplaceText("DATE", f"{params['date'][0]}/{params['date'][1]}/{params['date'][2]}")
    doc.ReplaceText("model", params["model"])
    doc.ReplaceText("isA", "X" if params["instrument"] in "aA" else "")
    doc.ReplaceText("isB", "X" if params["instrument"] in "bB" else "")
    doc.ReplaceText("isC", "X" if params["instrument"] in "cC" else "")
    return True


@debug
def WriteWordData(doc, corrected_data):
    """
    Writes corrected reflectance data to word docx cert in in table
    """
    for i in range(25, 251, 5):
        if i * 10 in corrected_data:
            v = corrected_data[i * 10]
            # rounding to 2 sig figs
            v = round(v, 2 - int(floor(log10(abs(v)))))
            v = str(v)
            while len(v) < 5:
                v += "0"
            doc.ReplaceText(f"w{i}", str(v))
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
    plt.axis([250, 2500, floor(min(plt_y) * 4) * 0.25, ceil(max(plt_y) * 4) * 0.25])
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
    doc.Save(f"{params['path']}DM-01400-010Rev04 {'99' if params['nominal reflectance'] == '99%' else 'Gray'} cal cert non NVLAP.docx")
    return True


@debug
def SavePdf() -> None:
    """
    Copies word doc cert as pdf

    This function exists simply to wrap docx2pdf_covert() method so the debug wrapper can be used
    """
    docx2pdf_convert(
        f"{params['path']}DM-01400-010Rev04 {'99' if params['nominal reflectance'] == '99%' else 'Gray'} cal cert non NVLAP.docx",
        f"{params['path']}{params['serial number']}.pdf",
    )
    return True


@debug
def CopyToUsb():
    """
    Copies raw data txt file and final cert pdf file to USB path specified in User Data\\Constants.txt
    """
    if os_path_exists(USB_PATH):
        docxName = f"DM-01400-010Rev04 {'99' if params['nominal reflectance'] == '99%' else 'Gray'} cal cert non NVLAP.docx"
        shutil_copyfile(f"{params['path']}{docxName}", f"{USB_PATH}{docxName}")
        shutil_copyfile(f"{params['path']}{params['serial number']}.pdf", f"{USB_PATH}{params['serial number']}.pdf")
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

    raw = CSV(f"{params['path']}Equation1.Sample.Cycle1.Equation1.csv")
    strayLight = CSV(f"{params['stray light path']}Equation1.Sample.Cycle1.Equation1.csv")
    doc = DOCX("User Data\\Ref-Cal-Cert-Template.docx")

    Rr = GetRr()
    if not Rr:
        return False

    corrected_data = CorrectData(raw, strayLight, Rr)
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
        os_system(f"start {params['path']} .")
        if os_path_exists(USB_PATH):
            os_system(f"start {USB_PATH} .")
        window.close()


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
        [sg.FolderBrowse(), sg.Input("(must have Equation1.Sample.Cycle1.Equation1.csv)", size=(108, 1), readonly=True)],
        # 2
        [sg.Text("")],
        # 3
        [sg.Text("Info", font="30px")],
        # 4
        [
            sg.Text("Model", size=(4, 1)),
            sg.Input("", size=(20, 1), key="Model"),
            sg.Text("Serial Number", size=(10, 1)),
            sg.Input("", size=(20, 1), key="Serial Number"),
            sg.Text("Nominal Reflectance"),
            sg.DropDown(["2%", "5%", "10%", "20%", "40%", "60%", "80%", "99%"], "99%", size=(7, 1), key="Nominal Reflectance", readonly=True),
            sg.Text("Client"),
            sg.DropDown([c for c in CLIENTS], "no client selected", size=(15, 1), key="Client", readonly=True),
        ],
        # 5
        [sg.Text("")],
        # 6
        [sg.Text("Stray Light Scan", font="30px")],
        # 7
        [
            sg.Text("Instrument"),
            sg.DropDown(["A", "B", "C"], "A", size=(5, 1), key="Instrument", readonly=True),
            sg.CalendarButton(
                "Select Date",
                target=(7, 3),
                format="%m/%d/%Y",
                enable_events=True,
            ),
            sg.Input(key="Date", size=(10, 1), enable_events=True, readonly=True),
            sg.DropDown([], "No date selected", size=(56, 1), enable_events=True, key="Stray Light Dropdown", readonly=True),
            sg.FolderBrowse(button_text="Manual Browse", target=(8, 0)),
        ],
        # 8
        [sg.Input("Stray Light Path (not selected)", key="Stray Light Path", size=(117, 1), readonly=True)],
        # 9
        [sg.Text("")],
        # 10
        [sg.Button("Execute", size=(20, 1), font="30px", pad=(290, 0))],
        # 11
        [sg.Text("Log", font="30px")],
        # 12
        [sg.Input(key="Log", size=(117, 5), readonly=True, enable_events=True)],
    ]

    window = sg.Window(title="Ref Cal Auto", layout=layout, margins=(0, 20))

    while True:
        event, values = window.read()

        if event == "Date":
            slps = GetStrayLightPaths(DateFromString(values["Date"]))
            if len(slps) == 0:
                window["Stray Light Dropdown"].update(values=[])
            else:
                window["Stray Light Dropdown"].update(values=slps)
        elif event == "Stray Light Dropdown":
            window["Stray Light Path"].update(f"{STRAY_LIGHT_DIR}{values['Stray Light Dropdown']}")
        elif event == "Execute":
            error = TestExecuteConditions(values)
            if error == 0:
                params = {
                    "path": f"{values['Browse']}\\",
                    "model": values["Model"],
                    "serial number": values["Serial Number"],
                    "nominal reflectance": values["Nominal Reflectance"],
                    "client": CLIENTS[values["Client"]],
                    "date": DateFromString(values["Date"]),
                    "instrument": values["Instrument"],
                    "stray light path": f"{values['Stray Light Path']}\\",
                }
                window["Log"].update("Executing...")
                Timer(1.0, AsyncExecute).start()
            else:
                window["Log"].update(error)
        elif event == sg.WIN_CLOSED:
            break
    window.close()


# !ENTRY POINT
main()

# !TESTING
# Execute(
#     {
#         "path": "C:\\Users\\wdelgiudice\\Downloads\\18%PF-1020-4436 - Copy\\",
#         "model": "CSRT-18-020",
#         "nominal reflectance": "99%",
#         "serial number": "PF-0921-4398",
#         "date": DateFromString("1/6/2021"),
#         "instrument": "B",
#         "stray light path": "\\\\lssvr-fs01\\Reflectance Lab\\Reflectance Calibrations\\stray light Summary.xls_files\\Stray Light 4-7-2021 C\\",
#         "client": {1000: (0.1, 0.21), 1500: (0.1, 0.21)},
#     }
# )
