# Reflectance Calibration Automation
#### Overview
#### User Manual
### Code Manual
## OVERVIEW:

This program is designed to automate the process of taking raw data from 
a Reflectance Calibration result in the RefLab and converting it to a Calibration Certificate.

Once a scan is taken, the scanning device creates a folder with data from the scan. In the folder is Equation1.Sample.Cycle1.Equation1.csv, which is used in conjunction with the stray light scan file, to calculate the corrected data.
To understand these calculations, see the 99%ReflectCalA.xls sheet, which was the old method of calculating corrected data.

Once the data is corrected, it is also tested under internal and external (customer) requirements. This was previously done in the 99%ReflectCalA.xls sheet, but is now automatically done in this script.

If the requirements are not met, the script halts and must be manually exited, and the root folder is renamed to {serial number}-FAIL, to indicate a non-passing calibration.

Otherwise, if the requirements are met, the script will next make a text file names {model name}.txt containing the corrected data at every nanometer for customer reference.

Next the Certificate is made, first as a word doc. This certificate is generated using a pre-existing word doc called template.docx, which is located in the User Data folder, which in the root directory of the executable. Opening up the template.docx file reveals that there are several variables or flags (encased in angled brackets < >) which will be replaced by the script with data from the scan.
This includes:
* serial number (sn)
* model name (model)
* device (isA, isB, isC)
* date (DATE)
* reflectance table
* reflectance graph

Then once this data has been filled in the Template, the word doc is saved and a pdf named {serial number} is created from the word doc.

Then, the root folder is renamed to the serial number and opened in file explorer.
Lastly, the script will attempt to copy the pdf cert and raw data txt file onto the USB drive if it is present. The USB path can be specified in the config.txt file (in /User Data/config.txt)

## USER MANUAL
To run the script, locate the Ref-Cal-Auto folder, which is most likely in Program Files, and double click on main.exe. Alternatively you can locate the shortcut on your desktop or in your taskbar.
Once the script starts running, you should see two windows, a console (which will display the progress / errors of the program), and the gui. You don't need to pay attention to the console for now.

In the gui, there are several field you must fill out, before clicking the Execute button to run the tool.

The first field is Root Folder. Click Browse and use the pop up file explorer to select the folder that contains the scan data you are making a certificate for. This folder must contain the Equation1.Sample.Cycle1.Equation1.csv file produced by the calibration machines.

The Geometry field defines wether this calibration is a Target or Puck / Standard.

The Material field defines the material you calibrated, either Spectralon or Permaflect.

The Target Size / Puck Diameter field defines the size / diameter of your geometry, and is used solely in the model name. Conventionally, a 2 inch puck is written as 020 or a 10 inch target is written as 100. This field can be manually entered or selected from a dropdown.

The Reflectance field defines the % Reflectance of the calibration and can be manually entered or selected from a dropdown.

The Serial Number field is used as the serial number for the product.

The Requirements dropdown is used to select additional requirements. Note that Labsphere's internal requirements will still always be applied, even if No Additional requirements is selected.

The next section is for determining the stray light scan, which is used to correct the reflectance data.

The instrument field is for defining which instrument was used to measure the reflectance.

The Select Date button, and the Select Scan dropdown, are used in conjunction to find a stay light scan from the server. By choosing the date of the stray light scan, the tool will automatically browse for stray light scans with the given date, and the dropdown will show a list of each stray light folder from the server. If you can't find a scan from the server this way, or your stray light scan is not yet on the server, you can skip the date & scan dropdown and click manual browse on the row below, and select your stray light scan folder (containing the Equation1.Sample.Cycle1.Equation1.csv file). Note that no matter which method is chosen, the stray light scan path will be displayed in the white text field at the bottom of the Stray Light Scan section.

Once all of the forms have been completed, you can click the Execute button to run the tool. If the is an error in one of your fields, the Log Field will tell you what you need to fix. 

If the Log shows a requirements failure. This means the data did not pass one of the internal or customer requirements, and the program will stop, renaming the root folder to FAIL, you must manually close out of the program, and handle the failed calibration. Note that the log will tell you exactly where the requirements were not met.

If the Log shows an ERROR. You must open up the console which appeared at the start of the program. This will display a all the information about the error, as well as the processes which previously succeeded. Select all of the text in the console and hit ctrl + shift + C to copy the console text. Send this to a developer who understand this software tool.

If there are no errors, the tool will automatically close itself and open up the root folder in file explorer containing the certificate in both docx and pdf form, as well as the stray light scan and raw data text file.

If the tool finds a USB it will also copy over the pdf cert and text file to the usb, and open up the usb in file explorer as well.

## CODE MANUAL

This tool is written in Python, and relies on several dependencies, imported as modules:
* os (operating system)     - handle files
* sys (system)              - handle files
* re (regular expression)   - regex searching
* shutil                    - copy files
* PySimpleGUI               - create gui
* docx                      - handle microsoft word docx files
* matplotlib                - create graphs
* docx2pdf                  - convert word doc to pdf
* math                      - extra math functions
* threading                 - asynchronous functions
* time                      - measure function time
* date                      - get current date
* pyinstaller               - used to compile script into executable

The script also depends on a folder located in the same directory called 'User Data' which contains three files:
* config.txt
* rr.txt
* template.docx

How it works:

First, there are four global variables:
* window        - holds the PySimpleGUI window object
* params        - hold the Parameters object
* config        - dictionary of configuration data
* internal_reqs - holds the static list of internal requirements 

The script starts by running setup, which populates the config dictionary from the /User Data/config.txt file. It also creates the layout list which is used to then initialize the PySimpleGUI. The window variable then holds on to the gui.

Then the main function runs. This function is in charge of maintaining the gui loop, and eventually executing all the script processes.

In the main while loop, we listen for events and respectively handle them in their own functions.

When the 'Execute' event is called, we jump into the ExecuteEvent function, which is the gateway to our program.

In the ExecuteEvent, we first summarize all our form fields into a Parameters object. Then we call params.isValid() to make sure the user did not forget or incorrectly fill out a field. This method will return a string containing the error message, which is relayed back to the log field, if the user did something incorrectly, otherwise it will return true and proceed to execution.

We specifically use the line:

    849|     Timer(1.0, AsyncExecute).start() 

to asynchronously execute our tool. By doing this we can update the gui and console as each step in the execution takes place. This avoids our application freezing and becoming unresponsive, which might cause the user to be unsure if their tool is actually working or not.

Jumping into the AsyncExecute, we are simply wrapping the Execute function, and then handling its output. If Execute returns true, 
we open up the root folder of the calibration, and attempt to open the usb folder (in file explorer). Then we forcibly stop the program, so as to stop all threads.

In the Execute function, we essentially go down a list of subtasks, linking up their results to produce the final cert, since this tool is the automation of many subtasks.

Tasks:

Get_rr
Retrieves RR data from /User Data/rr.txt. This is used to calculate corrected data

CorrectData
Combines the raw data, stray light scan, and rr data to produce corrected reflectance data

TestRequirements
Test the internal and customer requirements on the corrected data. If the data passes the requirements, it returns true, if it doesn't it returns a string explaining where the data failed to meet the requirements, which is then bubbled back to the Log and terminal.

RenameRootFolder
Renames root folder to the serial number


SaveStrayLight
Copies stray light file to the same directory of the cert so as to provide a way of re-creating the cert

SaveTextFile
Generates a text file with the name of the model name containing the corrected reflectance data at every nanometer

WriteWordMeta
Populates the word doc with metadata such as serial number, date, and instrument
The way the word doc is generated is through populating a template.
The template is located in /User Data/template.docx. Inside the template you can see variables inside angled brackets <> which represent data that needs to be populated. The function then uses regexs to find each variable and replaces it with the according data

WriteWordData
Populates the word doc reflectance table at every 50nms and uses significant figures drawn from the uncertainty table in the word doc


WriteWordGraph
Uses matlibplot to create a graph of the reflectance data, then saves the graph as a png, then places the png into the word doc, and then deletes the png.

SaveWord
Saves the word doc to the root folder

SavePdf
Generates a pdf from the word doc

CopyToUsb
Attempts to copy the pdf cert and data txt file onto the usb (the usb path is specified in /User Data/config.txt)


Notes:
* this script depends on a relative folder /User Data which must contain rr.txt, config.txt, and template.docx

* When compiling this script, copy the /User Data/ folder into the same directory as main.exe

* methods are imported and renamed with underscores so as to keep naming clear and concise will still importing the minimum number of methods / dependencies

* Execute function is separated into many parts so that each part can be wrapped with @debug.

* the debug function is a wrapper which measures the duration of the function, handles exceptions, and logs status


Contact delwillgiudice@gmail.com for questions regarding this tool






Nvlap checkbox

99% should say white





