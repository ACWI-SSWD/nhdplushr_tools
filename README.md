# NHDPlus HR VAA Navigator

The NHDPlusHR VAA Navigator performs navigation on the NHDPlusHR surface water network using NHDPlusHR Value Added Attributes found in the NHDPlusFlowlineVAA table.  The Navigator may be used through ArcMap via a special ArcMap toolbar or it may be called from user-written program code.  The Navigator will work on any NHDPlusHR gdb.  The gdb may contain a single Vector Processing Unit (VPU) or several VPUs that have been appended together.

The Navigator performs four types of navigation: upstream mainstem, upstream with tributaries, downstream mainstem, and downstream with divergences.   Navigations can begin and end on full NHDFlowline features or may begin and end as points along features.

Any of the four types of navigation may be stopped based on a user-supplied distance from the starting point.  Navigation results may also be filtered based on certain user specified criteria. 

## Disclaimer

This repository contains provisional or preliminary software and software specifications developed by Horizon Systems Corporation, under contract to the U.S. Geological Survey (USGS). This software is in the public domain. It has been used operationally within USGS, but has not been reviewed or approved by the USGS. The software is provided “AS IS”, to meet the need for timely best science. Posting this software does not imply any obligation by USGS or Horizon Systems Corporation for future support or further development of the software. This software is provided on the condition that neither the USGS nor the U.S. Government nor Horizon Systems shall be held liable for any damages resulting from the authorized or unauthorized use of the information. See Disclaimer.md. 

[Download](https://github.com/ACWI-SSWD/nhdplushr_tools/raw/master/docs/NHDPlusV2_VAA_Navigator_InstallGuide.docx) NHDPlusHR VAA Navigator Installation Guide 

[Download](https://github.com/ACWI-SSWD/nhdplushr_tools/raw/master/docs/NHDPlusHR_VAA_Navigator_UserGuide.docx) NHDPlusHR VAA Navigator Toolbar User Guide

## Windows Executable installation
System Requirements: Windows 7 Service Pack 1 64-bit; ArcGIS 10.5.1; Microsoft .NET Framework 4.0.3 or higher; Microsoft SQL Server 2012 Express LocalDB 64-bit; Microsoft SQL Server 2012 Management Studio
The primary testing environment has been ArcGIS 10.5.1; Windows 7 64-bit Service Pack 1; Microsoft .NET Framework 4.5.2.

Installation Order: Microsoft .NET, Microsoft SQL Server 2012 Express LocalDB, Microsoft SQL Server 2012 Management Studio, VAA Navigator

[Link to Microsoft .Net Framework version 4.5.2 Offline installer](http://www.microsoft.com/en-us/download/details.aspx?id=42642)

[Link to Microsoft .Net Framework version 4.5.2 Web installer](http://www.microsoft.com/en-us/download/details.aspx?id=42643)

[Download Microsoft SQL Server Install for NHDPlusHR Tools](http://www.horizon-systems.com/NHDPlusData/NHDPlusV21/Tools/NHDPlusTools_MSSQLServer2012ExpressLocalDB_x64_Install.7z)

[Download Microsoft SQL Server Management Studio Install](http://www.horizon-systems.com/NHDPlusData/NHDPlusV21/Tools/SQLManagementStudio_x64_ENU.7z)

## Source Code

Source code is available from the 'src' directory of this repository.