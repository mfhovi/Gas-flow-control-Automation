# Gas-flow-control-Automation
This stand alone program is developed for Mass flow controller (GMC1200) for gas flow automation/scheduling <br>
1) The system (e.g. PC or laptop) has to be connected via USB to RS232 adapter (FTDI) with the flow controller (GMC1200). Normally, both RS232 and GMC1200 have male connectors, in this case connection can be made directly with jumper wires in male pin no 2,3,5 (see male DB9-attached image) with the corresponding pins (no cross connection needed between Tx and Rx) of flow controller. <br>
<img width="1010" height="358" alt="RS-232_DE-9_Connector_Pinouts" src="https://github.com/user-attachments/assets/24bcb0da-279a-4117-a27c-4e526c2d967a" />

2) When appropriate connector is pluged into the system (e.g. PC), the active COM port can be viewed from "Device Manager". <br>
<img width="502" height="557" alt="Screenshot 2025-12-03 201614" src="https://github.com/user-attachments/assets/f7a1ee66-214a-4c5c-abe2-3f8c25b0ca62" />

3) From the flow controller, control mode has to be changed to RS232 by MENU->Control Mode->RS232->Ent.<br>

4) There are 2 python files, one for two channel flow and another for multichannel flow control, they can be run directly with python. Otherwise, after the python files are downloaded or copied in local drive (e.g. C:\Downloads) the following command has to be given in Command Prompt to get windows executable program, <br>
   a) cd "C:\download_folder"<br>
   b) pyinstaller --noconfirm --windowed --onefile --icon=NONE "gmc1200_gui_4 channel.py" <br>
   c) the executable program can be found in download_folder/dist <br>

5) After selecting appropriate COM port and Baud rate (9600 by default), commands can be given to the GMC1200 <br>
   a) With "gmc1200_gui_2 channel", total gas flow rate (carrier gas+ target gas) and concentration of target gas (e.g. 2% = 20,000ppm) has to be selected. after that manual ppm value can be input, the system will automatically calculate the flow rate according to total flow rate. Flow sequence can be given. <br>
<img width="1907" height="1020" alt="Screenshot 2025-12-03 201011" src="https://github.com/user-attachments/assets/400fe05c-2f1f-422f-b756-21c47ec390c9" />
   b) Similarly "gmc1200_gui_4 channel" can be used to control maximum 4 channel simultaneously with manual and automatic sequence. <br>
<img width="1918" height="1017" alt="Screenshot 2025-12-03 200743" src="https://github.com/user-attachments/assets/20dd9734-17af-44c3-ac73-d7ab59eb11bb" />
