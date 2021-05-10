from os.path import exists as os_path_exists
from os import remove as os_remove
from os import listdir as os_listdir
from os import system as os_system
from re import search as re_search
from re import match as re_match
from shutil import copyfile
import PySimpleGUI as sg
from docx import Document as docx_document
from docx.shared import Inches
import matplotlib.pyplot as plt
from docx2pdf import convert as docx2pdf_convert
from math import ceil, floor, log10

TARGET_DIR = "\\\\10.122.0.134\\Reflectance Lab\\Reflectance Calibrations\\PermaFlect Targets\\"
STRAY_LIGHT_DIR = "\\\\lssvr-fs01\\Reflectance Lab\\Reflectance Calibrations\\stray light Summary.xls_files\\"
USB_PATH = "D:\\"
PRODUCTION = True

log event is not firing and not causing execute to ever run

def ParseTablePosition(position) -> tuple[int, int]:
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
            raise ValueError(f"easyio.ParseTablePosition: c:{c}, r:{r} is invalid position")
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
        raise TypeError(f"easyio.ParseTablePosition: position: {type(position)} is not <class 'str'> or <class 'tuple'>")

    if type(position) is tuple and type(position[0]) is int and type(position[1]) is int:
        return position
    else:
        raise ValueError(f"easyio.ParseTablePosition: c:{c}, r:{r} is invalid position")


class CSV:
    def __init__(self, _path: str):
        self.path = _path
        file = open(self.path, "r")
        self._data = []
        for l in file.read().split("\n"):
            self._data.append(l.split(","))
        file.close()
        self.modified = False

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
            position = ParseTablePosition(position)
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
        position = ParseTablePosition(position)

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


def CorrectData(raw: DOCX, strayLight: DOCX, Rr: list[float]) -> list[float]:
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


def SavePDF(input, output) -> None:
    docx2pdf_convert(input, output)


def DateFromString(string: str) -> tuple:
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


def GetStrayLightPaths(date) -> list[str]:
    result = []
    for dir in os_listdir(STRAY_LIGHT_DIR):
        dirDate = DateFromString(dir)
        if dirDate and dirDate[0] == date[0] and dirDate[1] == date[1] and dirDate[2] == date[2]:
            result.append(dir)
    return result


def PopulateClients() -> dict:
    last = None
    clients = {}
    for l in open("Clients.txt").readlines():
        if re_match("\\s+.+", l):
            reg = re_search("(\\d+)\\s+(\\d+\\.*\\d*)\\-(\\d+\\.*\\d*)", l)
            if reg:
                clients[last][int(reg.group(1))] = (float(reg.group(2)), float(reg.group(3)))
            elif re_match("\\s+flatness", l):
                clients[last]["flatness"] = True
        else:
            last = l.strip()
            clients[last] = {}
    return clients


def Execute(params: dict) -> bool:
    print("Checking Conditions    ", end="")
    if not os_path_exists(f"{params['path']}Equation1.Sample.Cycle1.Equation1.csv"):
        raise Exception("Invalid Folder Selected: Equation1.Sample.Cycle1.Equation1.csv DOES NOT EXIST")
    if not os_path_exists(f"{params['stray light path']}Equation1.Sample.Cycle1.Equation1.csv"):
        raise Exception("No Stray Light Scan found")
    if not os_path_exists("Ref-Cal-Cert-Template.docx"):
        raise Exception("No Ref-Cal-Cert-Template.docx found in script directory")
    raw = CSV(f"{params['path']}Equation1.Sample.Cycle1.Equation1.csv")
    strayLight = CSV(f"{params['stray light path']}Equation1.Sample.Cycle1.Equation1.csv")
    doc = DOCX("Ref-Cal-Cert-Template.docx")
    docxName = f"DM-01400-010Rev04 {'99' if params['nominal reflectance'] == '99%' else 'Gray'} cal cert non NVLAP.docx"
    print("√")

    print("Retrieving Rr data from /Rr.txt    ", end="")
    Rr = [float(l) for l in open("Rr.txt").readlines()]
    print("√")

    print("Calculating Corrected Data    ", end="")
    corrected_data = CorrectData(raw, strayLight, Rr)
    print("√")

    print("Testing Client Requirements from /Client.txt    ", end="")
    for w in params["client"]:
        if type(w) is int:
            if corrected_data[w] < params["client"][w][0] or corrected_data[w] > params["client"][w][1]:
                raise Exception(
                    f"Corrected Data did not pass client requirements ({corrected_data[w]} @ {w} did not meet {params['client'][w][0]} <= reflectance <= {params['client'][w][1]} @ {w})"
                )
        elif w == "flatness":
            if corrected_data[1500] - corrected_data[1000] > 2 or corrected_data[1500] - corrected_data[1000] < 1:
                raise Exception(
                    f"Corrected Data did not pass client requirements ({corrected_data[1500] - corrected_data[1000]} did not meet flatness requirements 1 <= [ref @ 1500] - [ref @ 1000] <= 2)"
                )
    print("√")

    print(f"Generating {params['serial number'][-4:len(params['serial number'])]}-{params['model']}.txt file    ", end="")
    txt = open(f"{params['path']}{params['serial number'][-4:len(params['serial number'])]}-{params['model']}.txt", "w")
    stringdata = [f"{w}\t{corrected_data[w]}\n" for w in corrected_data]
    stringdata.insert(0, f"{params['serial number']}\nThis data is for reference only\n")
    txt.write("".join(stringdata))
    del stringdata
    print("√")

    print(f"Replacing docx template metadata (sn, model, instrument, date)    ", end="")
    doc.ReplaceText("sn", params["serial number"])
    doc.ReplaceText("DATE", f"{params['date'][0]}/{params['date'][1]}/{params['date'][2]}")
    doc.ReplaceText("model", params["model"])
    doc.ReplaceText("isA", "X" if params["instrument"] in "aA" else "")
    doc.ReplaceText("isB", "X" if params["instrument"] in "bB" else "")
    doc.ReplaceText("isC", "X" if params["instrument"] in "cC" else "")
    print("√")

    # print("Finding Uncertainty Values    ", end="")
    # uncertainty_table = doc.doc.tables[2]._cells
    # uncertainty = []
    # offset = -1
    # for i in range(9):
    #     if params["nominal reflectance"] == uncertainty_table[i].text:
    #         offset = i
    #         break
    # if offset == -1:
    #     return "Invalid Nominal Reflectance"
    # for i in range(9 + offset, len(uncertainty_table), 9):
    #     u = float(uncertainty_table[i].text)
    #     w = int(uncertainty_table[i - offset].text)
    #     j = 0
    #     while u < 1:
    #         u *= 10
    #         j += 1
    #     uncertainty.append((w, j))
    # print("√")

    print("Inserting Rounded Corrected Data into docx template table    ", end="")
    for i in range(25, 251, 5):
        if i * 10 in corrected_data:
            v = corrected_data[i * 10]
            # rounding to 2 sig figs
            v = round(v, 2 - int(floor(log10(abs(v)))))
            v = str(v)
            while len(v) < 5:
                v += "0"
            doc.ReplaceText(f"w{i}", str(v))
    print("√")

    print("Creating graph at /temp.png    ", end="")
    plt_x = [w for w in corrected_data]
    plt_y = [corrected_data[w] for w in corrected_data]
    plt.plot(plt_x, plt_y, color="black")
    # plt.title("Graph I: 8°/Hemispherical Spectral Reflectance")
    plt.ylabel("Reflectance Factor")
    plt.xlabel("Wavelength (nm)")
    plt.xticks([i for i in range(250, 2501, 250)])
    plt.axis(
        [
            250,
            2500,
            floor(min(plt_y) * 4) * 0.25,
            ceil(max(plt_y) * 4) * 0.25,
        ]
    )
    plt.savefig("temp.png")
    print("√")

    print("Replacing graph with /temp.png in docx template    ", end="")
    doc.ReplacePicture("graph", "temp.png", (7, 5.5))
    print("√")

    print("Removing /temp.png    ", end="")
    os_remove("temp.png")
    print("√")

    print(f"Saving docx to {params['path']}{docxName}    ", end="")
    doc.Save(f"{params['path']}{docxName}")
    print("√")

    print(f"Saving pdf of docx to {params['path']}{params['serial number']}.pdf", end="")
    SavePDF(f"{params['path']}{docxName}", f"{params['path']}{params['serial number']}.pdf")
    print("√")

    print(f"Copying data from {params['path']} to {USB_PATH}", end="")
    if os_path_exists(USB_PATH):
        copyfile(f"{params['path']}{docxName}", f"{USB_PATH}{docxName}")
        copyfile(f"{params['path']}{params['serial number']}.pdf", f"{USB_PATH}{params['serial number']}.pdf")
        print("√")
    else:
        print(f"\nNo USB located at {USB_PATH}")
    return True


def main() -> None:
    Params = {}
    executeOnNextFrame = False

    print("Retrieving Clients from /Clients.txt    ", end="")
    Clients = PopulateClients()
    print("√")

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
            sg.DropDown([c for c in Clients], "no client selected", size=(15, 1), key="Client", readonly=True),
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
    status = True

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
            else:
                result = "Executing..."
                Params = {
                    "path": f"{values['Browse']}\\",
                    "model": values["Model"],
                    "serial number": values["Serial Number"],
                    "nominal reflectance": values["Nominal Reflectance"],
                    "client": Clients[values["Client"]],
                    "date": DateFromString(values["Date"]),
                    "instrument": values["Instrument"],
                    "stray light path": f"{values['Stray Light Path']}\\",
                }
                executeOnNextFrame = True
            window["Log"].update(result)
        elif event == "Log":
            if executeOnNextFrame:
                try:
                    status = Execute(Params)
                    window["Log"].update("Finished: SUCCESS")
                    os_system(f"start {values['Browse']}\\ .")
                    if os_path_exists(USB_PATH):
                        os_system(f"start {USB_PATH} .")
                    break
                except Exception as e:
                    result = e
                    print(f"\ncopy-paste error message and send to:\n\twdelgiudice@labsphere.com\n\nError Message:\n{e}")
                    window["Log"].update("SEE TERMINAL FOR INSTRUCTIONS")
        elif event == sg.WIN_CLOSED:
            break
    window.close()
    if not status:
        print("Press enter to exit")
        input("")


if PRODUCTION:
    main()
else:
    Execute(
        {
            "path": "C:\\Users\\wdelgiudice\\Downloads\\18%PF-1020-4436 - Copy\\",
            "model": "CSRT-18-020",
            "nominal reflectance": "99%",
            "serial number": "PF-0921-4398",
            "date": DateFromString("1/6/2021"),
            "instrument": "B",
            "stray light path": "\\\\lssvr-fs01\\Reflectance Lab\\Reflectance Calibrations\\stray light Summary.xls_files\\Stray Light 4-7-2021 C\\",
            "client": {1000: (0.1, 0.21), 1500: (0.1, 0.21)},
        }
    )
