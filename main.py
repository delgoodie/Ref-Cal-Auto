# myFile = open("\\\\10.122.0.134\\Reflectance Lab\\Reflectance Calibrations\\PermaFlect Targets\\18%PF-0921-4400\\Info.txt")
import os
import easyio

folderPath = "C:\\Users\\wdelgiudice\\Downloads\\18%PF-1020-4436\\"


def getTables(dirPath: str) -> list[list[list[str]]]:
    csvTables = []
    for f in os.listdir(dirPath):
        if f[len(f) - 3 : len(f)] == "csv":
            csvTables.append(easyio.File(dirPath + f))
        # elif f[len(f) - 3 : len(f)] == "xls":
        #     csvTables.append(easyio.Read(dirPath + f))
    return csvTables


tables = getTables(folderPath)

print(tables[0].Read(("\n", 4)))
