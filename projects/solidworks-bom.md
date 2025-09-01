# 📊 SolidWorks BOM Export Automation

## Overview
A custom **VBA macro** for SolidWorks that extracts **Bill of Materials (BOM)** data directly into Excel.  
The tool automates repetitive tasks in assembly documentation.

## Features
- Export structured BOM with part numbers, quantities, and materials  
- Custom sorting: weldments, sheet metal, and others  
- Integration with Excel for reporting  

## Tools & Skills
- SolidWorks API (VBA)  
- Excel automation  
- Assembly documentation  

## Preview
vba
' Example snippet
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

[⬅️ Back to Projects Page](../README.md#-solidworks-macros-&-market-automation)
