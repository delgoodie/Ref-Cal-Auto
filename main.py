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


# tables = getTables(folderPath)

# print(tables[0].Write(("B", 4), "ok"))
#
#       * FILES
#       !PDF
#       {NAME}.pdf                                          PDF
#
#       100% or 0 Absorbance Basline.Correction.Raw.csv     CSV
#
#       CSTR-18-020.txt                                     TXT
#
#       Gray cal Cert.doc                                   DOC
#
#       Equation1.Sample.Cycle1.Equation1.csv               CSV
#
#       !XLS
#       GrayReflectCalA.xls                                 XLS
#
#       Info.txt                                            TXT
#
#       3x Sample552.Sample.Cycle{n}.Raw.csv                CSV
#
#
#       * PROCEDURE
#
#       - WRITE  GrayReflectCalA.Correctdata.A1 = Serial Name
#       - WRITE GrayReflectCalA.Correctdata[B][6:2250] = (100% or 0 Absorbance...).(100% or 0 Absorbance Baseline.C).[B][2:2262]
#       !Need access to deleted files
#       - WRITE GrayReflectCalA.Correctdata[H][6:2250] = (Internal CSV FILE).?.?
#       !Need access to deleted files
#       - WRITE GrayReflectCalA.Correctdata[J][6:2250] = (Internal CSV FILE).?.?
#
#       - +WRITE {ProductName}.txt = GrayReflectCalA.Data to disk.[M:N][6:2250]
#
#       - (Copy generated files to directory)
#
#       - Make Certificate

xls = easyio.File("C:\\Users\\wdelgiudice\\Downloads\\GrayReflectCalA.xls")

print(xls.Write("Data to disk", "B12", "okok"))

del xls