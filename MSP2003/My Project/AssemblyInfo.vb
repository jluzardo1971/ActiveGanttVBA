#Const VS_2005IDE = True
'//#Const VS_2008IDE = True
'//#Const VS_2010IDE = True

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

<Assembly: AssemblyTitle("ActiveGanttVBA Microsoft Project 2003 Integration Library")> 
<Assembly: AssemblyDescription("RELEASE VERSION .Net Framework 1.0")> 
<Assembly: AssemblyCompany("The Source Code Store LLC")> 
<Assembly: AssemblyProduct("ActiveGanttVBA Scheduler Web Server Control")> 
<Assembly: AssemblyCopyright("Copyright (c) 2002-2013 The Source Code Store LLC")> 
<Assembly: AssemblyTrademark("")> 
<Assembly: CLSCompliant(True)> 
<Assembly: ComVisible(False)> 
<Assembly: Guid("a5f5340d-a2c0-4618-8d08-2c8288a72d62")> 
<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyKeyFile("C:\SCS\SN\AG\VBA\MSP2003.snk")> 
<Assembly: AssemblyDelaySign(False)> 
<Assembly: AllowPartiallyTrustedCallers()> 
#If VS_2010IDE = True Then
    <Assembly: System.Security.SecurityRules(System.Security.SecurityRuleSet.Level1)>
#End If