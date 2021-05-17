# Reflectance Calibration Automation

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

// Nominal Reflectance
round to uncertainty table
12, 15, 18, 25, 50, 75

append -FAIL to Root folder on client req fail

puck diameter dropdown
target width dropdown

select puck or target

select material

derived model name