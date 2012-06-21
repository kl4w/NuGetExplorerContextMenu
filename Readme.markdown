# NuGetContextMenuHandler
Creates context menu in Windows Explorer which calls NuGet get extension command provided by https://github.com/BenPhegan/NuGet.Extensions.git


#### Remarks
1. Based on CSShellExtContextMenuHandler : http://1code.codeplex.com/wikipage?title=WinShell
2. Expects a NUGET_EXE environment variable to be set, the directory holding the nuget executable


#### Setup and Removal:

A. Setup

Run 'Visual Studio Command Prompt (2010)' (or 'Visual Studio x64 Win64 
Command Prompt (2010)' if you are on a x64 operating system) in the Microsoft 
Visual Studio 2010 \ Visual Studio Tools menu as administrator. Navigate to 
the folder that contains the build result VBShellExtContextMenuHandler.dll 
and enter the command:

    Regasm.exe NuGetContextMenuHandler.dll /codebase

The context menu handler is registered successfully if the command prints:

    "Types registered successfully"

B. Removal

Run 'Visual Studio Command Prompt (2010)' (or 'Visual Studio x64 Win64 
Command Prompt (2010)' if you are on a x64 operating system) in the Microsoft 
Visual Studio 2010 \ Visual Studio Tools menu as administrator. Navigate to 
the folder that contains the build result NuGetContextMenuHandler.dll 
and enter the command:

    Regasm.exe NuGetContextMenuHandler.dll /unregister

The context menu handler is unregistered successfully if the command prints:

    "Types un-registered successfully"
