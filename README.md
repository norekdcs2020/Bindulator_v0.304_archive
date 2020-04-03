# Bindulator

<p align="center">
       <a href="https://github.com/norekdcs2020/Bindulator/blob/master/LICENSE">
       <img src="https://img.shields.io/badge/License-GPLv2-red.svg" alt="Bindulator License">
       </a>
</p>

After countless binding, rebinding and unbinding binds in DCS World (+formats +uninstalls +DCS bind bugs) I lost track what function on which airplane is bound on which of my controllers. Unfortunately I have never found any tool for DCS World bind visualization (nor PDF nor kneeboard) so I made one for myself and my friends. Supposing it could also make some of the DCS World community happier, I have decided to make it public. I am not a proffesional programmer so please bear that in mind. If my effort inspire you to make a better, standalone program that does not need DCS HTML exports, MS Excel, PDF editor - I am already a fan of your work!

Bindulator is a MS Excel based tool for Digital Combat Simulator (DCS) keybind visualization. The Bindulator uses a mix of Excel, PowerQuery and VBA actions, macros and scripts and works together with a free PDF editor. At the moment **it has been tested and confirmed to work only with <a href = https://www.tracker-software.com/product/pdf-xchange-editor> PDF-Xchange® Editor </a> Version 8.0 compilation 336 Portable** (download free __"Editor Portable Version"__ and proceed accordingly to the program license). The Bindulator could use any other free PDF editor but it needs to be modified in such manner.

All rights to the Bindulator are reserved by 'norekdcs2020' (GPLv2). I do not consent to any paid or commercial use of the Bindulator. Contact me via the ED forum thread via link:
## **LINK**

Also, taking my friends' suggestions' into consideration, here is a link if you would like to support my work:
## **LINK**

<BR>
       
## Features & limitations
**Whole workbook:**
- Visualize any DCS World airplane binds - on PDF and on kneeboard (VR tested!),
- Support to visualize binds of up to two Controllers per a Template,
- Binding only one function to a button is accepted + modifier 1 + modifier 2,

**ORANGE** worksheet:
- **TODO** Marking duplicate button binds in red,
- (Advanced) After initialisation you can edit the names of functions before the PDF creation to suit your needs,

<BR>
       
## How to use
- **TODO** Check the brief <a href="LINK DO FILMU">youtube video</a> OR:
- Export DCS airplane binds via the "HTML Export" (fortunately, at the moment, it is one click per airplane),
- Open Bindulator,
- Choose your controller setup and airplane type via drop down lists at No. 1,
- Choose the exported DCS HTML files for controller 1 and controller 2 at. No. 2 and No. 3 accordingly,
- Press the **Initialise data** button at No. 4,
- Now you have imported your data to the YELLOW (raw import data) and ORANGE (processed import data) worksheets. If you don't like the DCS function names present in ORANGE worksheet, this is the time to change them,

<BR>
       
## Changelog
v0.3 - initial version of the Bindulator that was made public.

<BR>
       
## Known bugs
1) Only one color available (other than black) while using button modifiers - cannot change font color in PDF-Xchange Editor by keyboard shortcut.
