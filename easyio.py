import docx
import xlrd
import re


class CSV:
    def __init__(self, _path: str):
        self.path: _path
        self.file = open(_path)
        self.data = []
        for l in self.file.readlines():
            self.data.append(l.split(","))

    def _parse_position(self, position) -> tuple[int, int]:
        if type(position) is str:
            if ":" in position:
                c, r = position.split(":", 1)
            else:
                match = re.match("^([a-zA-Z]+)(\\d+)&", position)
                if len(match.regs) > 1:
                    c = position[match.regs[0][0] : match.regs[0][1]]
                    r = position[match.regs[1][0] : match.regs[1][1]]
            if c and r:
                position = (ord(c.trim().upper()) - 65, r.trim())
            else:
                raise ValueError("easyio.CSV._parse_position: c:{0}, r:{1} is invalid position".format(c, r))
        elif type(position) is tuple:
            c, r = position
            if type(c) is int or type(c) is float or re.match("^\\d+$", c):
                c = int(c)
            elif re.match("^[A-Za-z]+$", c):
                c = c.upper()
                col = 0
                i = 1
                for char in c:
                    col += (ord(char) - 65) * i
                    i *= 26
                c = col
            r = int(r)
            position = (c, r)
        else:
            raise TypeError("easyio.CSV._parse_position: position: {0} is not <class 'str'> or <class 'tuple'>".format(str(type(position))))

        if type(position) is tuple and type(position[0]) is int and type(position[1]) is int:
            return position
        else:
            raise ValueError("easyio.CSV._parse_position: c:{0}, r:{1} is invalid position".format(c, r))

    def Read(self, position="*"):
        """
        Read from csv table cell[s]

        position -- cell position: '*', (0, 34) | ('A', 34) | ('A', '34') | 'A34' | 'A:34'
        if position == '*' then a 2D array of all cols & rows will be returned

        return -- the overwriten data: 'data before Write()'
        """
        if position == "*":
            return self.data
        else:
            position = self._parse_position(position)
            if position[1] < len(self.data) and position[0] < len(self.data[position[0]]):
                return self.data[position[1]][position[0]]
            else:
                raise IndexError("easyio.CSV.Read: position:{0} not in cols: {1}, rows: {2}".format(str(position), len(self.data[0]), len(self.data)))

    def Write(self, position, data: str = ""):
        """
        Write string to single csv table cell

        position -- cell position: (0, 34) | ('A', 34) | ('A', '34') | 'A34' | 'A:34'
        data -- string of new cell contents: '34.213'

        return -- the overwriten data: 'data before Write()'
        """
        position = self._parse_position(position)

        pre = self.data[position[0], position[1]]
        self.data[position[0], position[1]] = data
        return pre

    def __del__(self):
        string = ""
        for c in self.data:
            for r in c:
                string += r + ","
            string += "\n"
        if self.file.writable():
            self.file.write(string)
        self.file.close()


class XLSX:
    def __init__():
        0


class XLS:
    def __init__():
        0


# TODO: .docx support
class DOCX:
    def __init__():
        0


# TODO: .txt support
class TXT:
    def __init__():
        0


# Constant list of supported file types
FILES = {"csv": CSV, "xlsx": XLSX, "xls": XLS, "docx": DOCX, "txt": TXT}

# Returns the file extension
def FileType(path: str) -> str:
    match = re.match("^.*\\.(\\w+)$", path)
    if len(match.regs) > 1:
        return path[match.regs[-1][0] : match.regs[-1][1]]
    else:
        print("ERROR [" + path + "] has no File Type")
        return ""


# Used to read once
def Read(path: str):
    filetype = FileType(path)
    if len(filetype) == 0:
        return -1

    if filetype == "csv":
        table = []
        for l in open(path).readlines():
            table.append(l.split(","))
        return table
    elif "xls" in filetype:
        sheets = {}
        excel = xlrd.open_workbook(path) if filetype == "xlsx" else xlrd.open_workbook_xls(path)
        for s in excel.sheets():
            sheets[s.name] = []
            for r in range(s.nrows):
                sheets[s.name].append([])
                for c in s.row(r):
                    sheets[s.name][-1].append(c.value)
        return sheets
    elif filetype == "docx":
        0
    elif filetype == "txt":
        0
    else:
        return print("easyio.Read: ERROR [" + filetype + "] Not Supported")


# TODO: Write Function
# Used to write once
def Write(path: str, position, data: str):
    0


def File(path: str):
    type = FileType(path)
    if len(type) > 0:
        for t in FILES.keys():
            if type == t:
                return FILES[t](path)
    else:
        return None