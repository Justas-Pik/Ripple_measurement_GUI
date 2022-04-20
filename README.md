# Power Supply Ripple measurement GUI
This repository contains Ripple_measurement_GUI.py Python file which can by built to GUI application for ripple measurement system.

## Description
##### Devices:
1. Oscilloscope
2. Electronic Load
3. Testing Board
4. Gateway
5. LAN Switch
6. 4 LAN Cables
7. Oscilloscope Cable
8. 2 Electronic Load cables with banana plugs
9. Power supplies for Gateway, Switch and Testing board
10. Wire to connect Gateway to Testing board (3 pins) 

##### System connections
1. Connect oscilloscope, electronic load, gateway and computer to the switch.
2. Plug gateway to the testing board using 3 pin wire, and also connect switch, gateway and testing board to mains.
3. Connect oscilloscope and electronic load to testing board using oscilloscope and electronic load cables.
4. Connect power supplies for testing (Max. 4).
5. Turn on Ripple_measurement_GUI application. (For building application from .py file look below).
6. Enter and set IP adresses of electronic load, oscilloscope and gateway, also gateway password. (If addresses have not been changed, the initial values should work).
7. Select channels of testing board and measurement type.
8. Enter current for every channel separatly.
9. Start measurements using "Start" button.
10. After the first measurement, you need to choose in which folder to save the document.*
11. After taking the measurements, disconnect from the devices using "Close" button. 
12. Close the program using "X" or "Quit" buttons.

*If you don't select where to save the document, it will only be saved in the folder that contains the GUI application file. If you select an additional folder, the data will be saved to your additional directory and to folder that contains your application.
##### Building application from code

1. Then open the file in "Visual Studio Code", or other similar application.
2. Rename Ripple_measurement_Gui.py to Ripple_measurement_Gui.pyw.** 
3. Write "pip install pyinstaller" to the terminal window.
4. After "pyinstaller" module has been installed, write "pyinstaller --onefile Ripple_measurement_GUI.pyw" to the terminal window.
5. Open "dist" folder and there you have your GUI.***

** We do this to prevent the unnecessary console from opening. It can be done by opening "Explorer" at left upper corner, then pressing "Open Folder", selecting the folder where the file is located, then again opening "Explorer", right clicking on "Ripple_measurement_GUI.py" and adding "w" to the end of filename.
*** There will be more files created during compilation of .exe file, and they can be deleted.


## Comments
- Measurements and channels can be selected after first connection to devices. 
- For enableing current entry field or submit button, use "Cancel" button. 
- When changing dynamic parameters, press "Change" button, enter the values and confirm with the "Submit" button.
- After changing dynamic settings and pressing the Change button, the initial values are listed.
- In dynamic measurement, level B corresponds to the level entered next to the measured channel.

## Advantages
- Possibility to select channels and measurements.
- Current input for each channel.
- Ability to save the document to a preselected folder, if not selected when asked after first measurement.
- For No Load, Full Load, and Dynamic Load measurements, an additional table with the transformed data is generated.
- Ability to change dynamic settings.
- Messages are output after the measurement, as well as if no current is applied and full or dynamic load measurements are desired, also by entering dynamic parameters that do not correspond to the electronic load range or entering letters instead of the address.

## Disadvantages

- Document is automatically saved in the folder that contains .exe file. If additional location is selected the document is saved in both the specified folder and the .exe file folder. If "No load" measurement additional path is not specified (i.e., click Cancel in the pop-up window), each measurement is split into a separate document and placed in a folder that contains the user interface. When for example "No Load" and "Full Load" measurements are selected and no path is specified, the measurement is stopped after "Full Load", the current remains on, and the document is placed in the GUI folder.
- The current input fields (Current CH:1,2,3,4) are not protected against entering letters.
- Messages are not output in real time.
- Closing the application does not close the device connections. You can log out by pressing the Close button in the user interface, only then you can close the program by pressing the X or Quit buttons in the top corner.
- "Long term" measurement is commented, after removing comments and running the command calibration would take place, then after 1h measurement and data recording would be done.
- Disconnection from devices does not reported.
