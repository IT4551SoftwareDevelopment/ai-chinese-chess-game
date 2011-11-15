<html>
<body>
<pre>
<h1>Build Log</h1>
<h3>
--------------------Configuration: MfcXq - Win32 (WCE ARMV4I) Release--------------------
</h3>
<h3>Command Lines</h3>
Creating temporary file "C:\DOCUME~1\xxzx1\LOCALS~1\Temp\RSPFA.tmp" with contents
[
/nologo /W3 /D "ARM" /D "_ARM_" /D "ARMV4I" /D UNDER_CE=500 /D _WIN32_WCE=500 /D "WCE_PLATFORM_xnj1" /D "UNICODE" /D "_UNICODE" /D "NDEBUG" /FR"E:\Wince模拟器\共享文件夹/" /Fp"E:\Wince模拟器\共享文件夹/MfcXq.pch" /Yu"stdafx.h" /Fo"E:\Wince模拟器\共享文件夹/" /QRarch4T /QRinterwork-return /O2 /MC /c 
"G:\SP\eVC\MfcXq_Vert(GPS用)\MfcXqDlg.cpp"
]
Creating command line "clarm.exe @C:\DOCUME~1\xxzx1\LOCALS~1\Temp\RSPFA.tmp" 
Creating temporary file "C:\DOCUME~1\xxzx1\LOCALS~1\Temp\RSPFB.tmp" with contents
[
/nologo /base:"0x00010000" /stack:0x10000,0x1000 /entry:"wWinMainCRTStartup" /incremental:no /pdb:"E:\Wince模拟器\共享文件夹/MfcXqVert.pdb" /out:"E:\Wince模拟器\共享文件夹/MfcXqVert.exe" /subsystem:windowsce,5.00 /MACHINE:THUMB 
"E:\Wince模拟器\共享文件夹\MfcXq.obj"
"E:\Wince模拟器\共享文件夹\MfcXqDlg.obj"
"E:\Wince模拟器\共享文件夹\SearchInfo.obj"
"E:\Wince模拟器\共享文件夹\StdAfx.obj"
"E:\Wince模拟器\共享文件夹\MfcXq.res"
]
Creating command line "link.exe @C:\DOCUME~1\xxzx1\LOCALS~1\Temp\RSPFB.tmp"
<h3>Output Window</h3>
Compiling...
MfcXqDlg.cpp
Linking...
   Creating library E:\Wince模拟器\共享文件夹/MfcXqVert.lib and object E:\Wince模拟器\共享文件夹/MfcXqVert.exp
LINK : warning LNK4089: all references to 'WININET.dll' discarded by /OPT:REF
LINK : warning LNK4089: all references to 'commdlg.dll' discarded by /OPT:REF
Creating command line "bscmake.exe /nologo /o"E:\Wince模拟器\共享文件夹/MfcXq.bsc"  E:\Wince模拟器\共享文件夹\StdAfx.sbr E:\Wince模拟器\共享文件夹\MfcXq.sbr E:\Wince模拟器\共享文件夹\MfcXqDlg.sbr E:\Wince模拟器\共享文件夹\SearchInfo.sbr"
Creating browse info file...
<h3>Output Window</h3>




<h3>Results</h3>
MfcXqVert.exe - 0 error(s), 2 warning(s)
</pre>
</body>
</html>
