# Microsoft Developer Studio Project File - Name="Sword" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Application" 0x0101

CFG=Sword - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "Sword.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "Sword.mak" CFG="Sword - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "Sword - Win32 Release" (based on "Win32 (x86) Application")
!MESSAGE "Sword - Win32 Debug" (based on "Win32 (x86) Application")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "Sword - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 5
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /FR /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x404 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x404 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /machine:I386
# ADD LINK32 /nologo /subsystem:windows /machine:I386

!ELSEIF  "$(CFG)" == "Sword - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_AFXDLL" /D "_MBCS" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x404 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x404 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept
# ADD LINK32 /nologo /subsystem:windows /debug /machine:I386 /pdbtype:sept

!ENDIF 

# Begin Target

# Name "Sword - Win32 Release"
# Name "Sword - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Group "CA-SDK Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\_cacontrol.cpp
# End Source File
# Begin Source File

SOURCE=.\ca200srvr.cpp
# End Source File
# Begin Source File

SOURCE=.\EnterDlg.cpp
# End Source File
# End Group
# Begin Group "MFC Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\DlgProxy.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\Sword.odl
# End Source File
# Begin Source File

SOURCE=.\Sword.rc
# End Source File
# End Group
# Begin Group "UnUse Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Group "ExcelForm"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\Exl13Form.cpp
# End Source File
# Begin Source File

SOURCE=.\Exl25Form.cpp
# End Source File
# Begin Source File

SOURCE=.\ExlLB2Source.cpp
# End Source File
# Begin Source File

SOURCE=.\ExlLightForm.cpp
# End Source File
# Begin Source File

SOURCE=.\ExlOQCForm.cpp
# End Source File
# Begin Source File

SOURCE=.\ExlSECForm.cpp
# End Source File
# End Group
# Begin Source File

SOURCE=.\CheckBox.cpp
# End Source File
# Begin Source File

SOURCE=.\DefMsrPoints.cpp
# End Source File
# Begin Source File

SOURCE=.\excel.cpp
# End Source File
# Begin Source File

SOURCE=.\HideAndShow.cpp
# End Source File
# Begin Source File

SOURCE=.\InitialDlg.cpp
# End Source File
# Begin Source File

SOURCE=.\MsrItem.cpp
# End Source File
# Begin Source File

SOURCE=.\PatterntoType.cpp
# End Source File
# Begin Source File

SOURCE=.\xlef.cpp
# End Source File
# End Group
# Begin Source File

SOURCE=.\Sword.cpp
# End Source File
# Begin Source File

SOURCE=.\SwordDlg.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Group "MFC Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\DlgProxy.h
# End Source File
# Begin Source File

SOURCE=.\EnterDlg.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# End Group
# Begin Group "CA-SDK Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\_cacontrol.h
# End Source File
# Begin Source File

SOURCE=.\ca200srvr.h
# End Source File
# End Group
# Begin Group "UnUse Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\excel.h
# End Source File
# Begin Source File

SOURCE=.\InitialDlg.h
# End Source File
# Begin Source File

SOURCE=.\xlef.h
# End Source File
# End Group
# Begin Source File

SOURCE=.\Sword.h
# End Source File
# Begin Source File

SOURCE=.\SwordDlg.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\res\008.bmp
# End Source File
# Begin Source File

SOURCE=.\res\bitmap3.bmp
# End Source File
# Begin Source File

SOURCE=.\res\exit.ico
# End Source File
# Begin Source File

SOURCE=.\logo.bmp
# End Source File
# Begin Source File

SOURCE=.\res\max.ico
# End Source File
# Begin Source File

SOURCE=.\res\Sword.ico
# End Source File
# Begin Source File

SOURCE=.\res\Sword.rc2
# End Source File
# Begin Source File

SOURCE=".\res\¥¼©R¦W -2«þ¨©.bmp"
# End Source File
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# Begin Source File

SOURCE=.\Sword.reg
# End Source File
# End Target
# End Project
# Section Sword : {EF009364-CDD3-43F0-BE96-E6A920B29D60}
# 	2:21:DefaultSinkHeaderFile:_xycontrol.h
# 	2:16:DefaultSinkClass:C_xyControl
# End Section
# Section Sword : {A77E545F-15E1-47D0-882E-1CFD59FF43B2}
# 	2:5:Class:C_CaControl
# 	2:10:HeaderFile:_cacontrol.h
# 	2:8:ImplFile:_cacontrol.cpp
# End Section
# Section Sword : {745B81BE-B2E5-4655-A8AF-C88C867BCA39}
# 	2:5:Class:C_xyControl
# 	2:10:HeaderFile:_xycontrol.h
# 	2:8:ImplFile:_xycontrol.cpp
# End Section
# Section Sword : {1F2F52A5-CE61-45F1-B737-75A56ED77422}
# 	2:21:DefaultSinkHeaderFile:_cacontrol.h
# 	2:16:DefaultSinkClass:C_CaControl
# End Section
