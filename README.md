# Bindulator

<p align="center">
       <a href="https://github.com/norekdcs2020/Bindulator/blob/master/LICENSE">
       <img src="https://img.shields.io/badge/License-GPLv2-red.svg" alt="Bindulator License">
       </a>
</p>


Bindulator is a MS Excel based tool for Digital Combat Simulator (DCS) keybind visualization. The Bindulator uses a mix of Excel, PowerQuery and VBA actions, macros and scripts and works together with a free PDF editor. 

After countless binding, rebinding and unbinding binds in DCS World (+formats +uninstalls +DCS bind bugs) I lost track what function on which airplane is bound on which button of my controllers. Unfortunately I have never found any tool for DCS World bind visualization (nor PDF nor kneeboard) so I made one for myself and my friends. Supposing it could also make some of the DCS World community happier, I have decided to make it public. I am not a proffesional programmer so please bear that in mind and use the Bindulator at your own risk. If my effort inspire you to make a better, standalone program that does not need DCS HTML exports, MS Excel, PDF editor - I am already a fan of your work!

All rights to the Bindulator are reserved by 'norekdcs2020' (GPLv2). I do not consent to any paid or commercial use of the Bindulator. Contact me via Github, norekdcs2020@gmail.com or the ED forum thread <a href="LINK DO WATKU NA FORUM ED">LINK DO WATKU NA FORUM ED</a>.

Also, taking my friends' suggestions' into consideration, here is a link if you would like to support my work: <a href="LINK DO PAYPALA">LINK DO PAYPALA</a>.

<BR>

## Compatibility
The Bindulator has been checked to work in this software configuration:
- At the moment it has been tested and confirmed to work only with <a href = https://www.tracker-software.com/product/pdf-xchange-editor> PDF-Xchange® Editor </a> Version 8.0 compilation 336 Portable (download free __"Editor Portable Version"__ and proceed accordingly to the program license). The Bindulator could use any other free PDF editor but it needs to be modified in such manner,
- DCS OpenBeta 2.5.6.45915,
- MS Office 2019, 
- MS Windows 10 Pro.

<BR>
       
## Features & limitations
### Whole workbook:
- Visualize any DCS World airplane binds - on PDF and on kneeboard (VR tested!),
- Support to visualize binds of up to two Controllers per a Template,
- The Bindulator relies on DCS World HTML exports because it was easier for me to make the tool work,
- Binding only one function to a button is accepted + modifier 1 + modifier 2 (no support for more modifiers, sorry!),
- Big red/green/yellow text fields giving you feedback about current Bindulator state,
- DUAL Templates (two controllers on one page) will not be readable in VR (too much stuff going on),
- Remember it's just an amateur MS Excel file - please be patient and restart if necessary.

### ORANGE and YELLOW worksheets:
**TODO YELLOW** - Marking in red not supported multiple functions per button binds,
**TODO** - Marking in violet not supported multiple buttons per function binds,

### ORANGE worksheet:
- (Advanced) After initialisation you can edit the names of functions before the PDF creation to suit your needs.

### Support for controllers + created templates:
- Thrustmaster Warthog Stick,
- Thrustmaster Warthog Throttle,
- Thrustmaster T.16000M Stick,
- Thrustmaster T.16000M FCS Stick,
- Thrustmaster TWCS Throttle,
- Virpil (VPC) MongoosT-50CM2 Stick,
- Logitech Force 3D Pro Stick,
- Logitech Extreme 3D Pro Stick,
- Saitek X52 Stick,
**TODO** - Saitek X52 Throttle,
- Saitek X52 Pro Stick,
= Saitek X52 Pro Throttle.

<BR>
       
## How to use
- **TODO** Check the brief tutorial video <a href="LINK DO FILMU YOUTUBE">LINK DO FILMU YOUTUBE</a>,

OR:

### PDF bind visualization creation
- Export DCS airplane binds via the "HTML Export" (fortunately, at the moment, it is one click per airplane),
- Open Bindulator. The **GREEN** worksheet should appear,
- Choose your controller setup and airplane type via drop down lists at No. 1,
- Choose the exported DCS HTML files for controller 1 and controller 2 (if you need a controller 2 on the same Template) at. No. 2 and No. 3 with according buttons. The paths to the HTML files will appear in the `F` column,
- Press the **Initialise data** button at No. 4 and wait for data refresh start and end prompts. If a `DUPLICATES PRESENT` text appears next to No. 4 button, correct your binds,
- Now you have imported your data to the **YELLOW** (raw import data) and **ORANGE** (processed import data) worksheets. If you don't like the DCS function names present in **ORANGE** worksheet, this is the time to change them at `Funkcja` field (the `Kategoria funkcji` is only a helper column),
- Choose your **compatible** Template at No. 5 button. The path to the Template will appear in the `F` column,
- Choose the folder for saving bind visualization PDF files at No. 6 button. The path to the folder will appear in the `F` column,
- Point to the .exe file of your PDF editor at No. 7 button. The path to the .exe file will appear in the `F` column,

**WARNING BEFORE NO. 8! AFTER YOU PRESS THE BUTTON DO NOT TOUCH THE MOUSE, NOR THE KEYBOARD - UNLESS PROMPTED ON THE SCREEN OTHERWISE. IF YOU CHANGE THE FOCUS OF ACTIVE WINDOW, BAD THINGS COULD HAPPEN. USE AT YOUR OWN RISK.**

- Press the No. 8 button and watch how the Bindulator does the job,
- Now you can either use the PDF created (if something is wrong or not optimal, you can correct it now manually) or go further and export the file as a kneeboard in DCS,

### Kneeboard export from PDF files
- Check if the path to your newly created PDF file at No. 9 in the `F` column is correct,
- Choose the DCS World main kneeboard folder at No. 10. If the folder `Kneeboard` does not exist in the `/Saved Games/DCS.openbeta/` path - create it. After choosing the kneeboard folder with the Bindulator, the path to the folder will appear in the `F` column,

**WARNING BEFORE NO. 11! AFTER YOU PRESS THE BUTTON DO NOT TOUCH THE MOUSE, NOR THE KEYBOARD - UNLESS PROMPTED ON THE SCREEN OTHERWISE. IF YOU CHANGE THE FOCUS OF ACTIVE WINDOW, BAD THINGS COULD HAPPEN. USE AT YOUR OWN RISK.**

- Press the No. 11 button and watch how the Bindulator does the job,
- Check if everything works in DCS World. Do not forget to bind keys for Kneeboard ON/OFF or GLANCE and NEXT/PREVIOUS kneeboard pages. Your kneeboards should appear after pressing the NEXT button for a few times.
- Profit?
- Repeat previous steps if needed for different controllers and airships.

<BR>
       
## Changelog
v0.3 - initial version of the Bindulator that was made public.

<BR>
       
## Known bugs
1) Only one color available (other than black) while using button modifiers - cannot change font colour in PDF-Xchange Editor by keyboard shortcut,
2) Sometimes Excel loses focus after the Bindulator initialisation or bind visualization (the icon of background task does not appear) - you just need to press the mouse on the open Excel file,
3) Often the initialisation process takes some time (~30 s), maybe too much but hey - it works,
4) There is a chance that you will experience different axis naming - please let me know if you do not have feedback for Z axis (because on my computer it might be RZ axis).
