# VBFlexGrid64

This is an x64-compatible port of Krool's VBFlexGrid for twinBASIC

https://github.com/Kr00l/VBFLXGRD

https://www.vbforums.com/showthread.php?848839-VBFlexGrid-Control-(Replacement-of-the-MSFlexGrid-control)/page17

![image](https://github.com/fafalone/VBFlexGrid64/assets/7834493/9c4b9fe5-9cb9-4831-958e-189d8f15c497)


**Notes**

-This is currently set up like the StdExe version.

-Property pages have been prepared for x64 but tB does not yet support them for in-project controls (.ctl as opposed to .ocx). 

-For this project I made an x64 version of OLEGuids as a tB Package (due to longstanding issue with low-level COM/OLE interface redefinitions and midl making a TLB impossible); the project files are included here, but it's built into VBFlexGridDemo.twinproj too, so are included only for use in other projects.
