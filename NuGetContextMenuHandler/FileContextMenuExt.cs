/********************************** Module Header **********************************\
Module Name:  FileContextMenuExt.cs
Project:      CSShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The FileContextMenuExt.cs file defines a context menu handler by implementing the 
IShellExtInit and IContextMenu interfaces.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***********************************************************************************/

#region Using directives
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NuGetContextMenuHandler.Properties;
using System.Drawing;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
#endregion


namespace NuGetContextMenuHandler
{
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("B1F1405D-94A1-4692-B72F-FC8CAF8B8700"), ComVisible(true)]
    public class FileContextMenuExt : IShellExtInit, IContextMenu
    {
        // The name of the selected file.
        private string selectedFile;

        private string menuText = "&Call NuGet Get";
        private IntPtr menuBmp = IntPtr.Zero;
        private string verb = "csdisplay";
        private string verbCanonicalName = "CSDisplayFileName";
        private string verbHelpText = "Call NuGet Get";
        private uint IDM_DISPLAY = 0;
        private ConfigHandler _config = new ConfigHandler();

        public FileContextMenuExt()
        {
            // Load the bitmap for the menu item.
            Bitmap bmp = Resources.download;
            bmp.MakeTransparent(bmp.GetPixel(0, 0));
            this.menuBmp = bmp.GetHbitmap();
        }

        ~FileContextMenuExt()
        {
            if (this.menuBmp != IntPtr.Zero)
            {
                NativeMethods.DeleteObject(this.menuBmp);
                this.menuBmp = IntPtr.Zero;
            }
        }


        void OnVerbDisplayFileName(IntPtr hWnd, int pos)
        {
            if (pos < 3)
            {
                _config.SwitchOptionValue(pos);
            }
            else
            {
                var sourcePos = pos - 3;
                Process myProcess = new Process();
                myProcess.EnableRaisingEvents = false;
                myProcess.StartInfo.FileName = @"cmd.exe";
                var nugetLoc = Environment.GetEnvironmentVariable("NUGET_EXE");
                var excludeVersion = bool.Parse(_config.ConfigElement("useVersionedPackages").Value) ? string.Empty : " -x";
                var latest = bool.Parse(_config.ConfigElement("getLatest").Value) ? " -latest" : string.Empty;
                var clean = bool.Parse(_config.ConfigElement("cleanPackages").Value) ? " -clean" : string.Empty;
                myProcess.StartInfo.Arguments = string.Format("/k {0}\\NuGet.exe get \"{1}\" -s \"{2}\"{3}{4}{5}", nugetLoc, this.selectedFile, _config.Sources.ElementAt(sourcePos).Value, excludeVersion, latest, clean);
                myProcess.Start();
            }
        }


        #region Shell Extension Registration

        [ComRegisterFunction()]
        public static void Register(Type t)
        {
            try
            {
                ShellExtReg.RegisterShellExtContextMenuHandler(t.GUID, "Directory",
                    "NuGetContextMenuHandler.FileContextMenuExt Class");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                throw;  // Re-throw the exception
            }
        }

        [ComUnregisterFunction()]
        public static void Unregister(Type t)
        {
            try
            {
                ShellExtReg.UnregisterShellExtContextMenuHandler(t.GUID, "Directory");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); // Log the error
                throw;  // Re-throw the exception
            }
        }

        #endregion


        #region IShellExtInit Members

        /// <summary>
        /// Initialize the context menu handler.
        /// </summary>
        /// <param name="pidlFolder">
        /// A pointer to an ITEMIDLIST structure that uniquely identifies a folder.
        /// </param>
        /// <param name="pDataObj">
        /// A pointer to an IDataObject interface object that can be used to retrieve 
        /// the objects being acted upon.
        /// </param>
        /// <param name="hKeyProgID">
        /// The registry key for the file object or folder type.
        /// </param>
        public void Initialize(IntPtr pidlFolder, IntPtr pDataObj, IntPtr hKeyProgID)
        {
            if (pDataObj == IntPtr.Zero)
            {
                throw new ArgumentException();
            }

            FORMATETC fe = new FORMATETC();
            fe.cfFormat = (short)CLIPFORMAT.CF_HDROP;
            fe.ptd = IntPtr.Zero;
            fe.dwAspect = DVASPECT.DVASPECT_CONTENT;
            fe.lindex = -1;
            fe.tymed = TYMED.TYMED_HGLOBAL;
            STGMEDIUM stm = new STGMEDIUM();

            // The pDataObj pointer contains the objects being acted upon. In this 
            // example, we get an HDROP handle for enumerating the selected files 
            // and folders.
            IDataObject dataObject = (IDataObject)Marshal.GetObjectForIUnknown(pDataObj);
            dataObject.GetData(ref fe, out stm);

            try
            {
                // Get an HDROP handle.
                IntPtr hDrop = stm.unionmember;
                if (hDrop == IntPtr.Zero)
                {
                    throw new ArgumentException();
                }

                // Determine how many files are involved in this operation.
                uint nFiles = NativeMethods.DragQueryFile(hDrop, UInt32.MaxValue, null, 0);

                // This code sample displays the custom context menu item when only 
                // one file is selected. 
                if (nFiles == 1)
                {
                    // Get the path of the file.
                    StringBuilder fileName = new StringBuilder(260);
                    if (0 == NativeMethods.DragQueryFile(hDrop, 0, fileName,
                        fileName.Capacity))
                    {
                        Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                    }
                    this.selectedFile = fileName.ToString();
                }
                else
                {
                    Marshal.ThrowExceptionForHR(WinError.E_FAIL);
                }
            }
            finally
            {
                NativeMethods.ReleaseStgMedium(ref stm);
            }
        }

        #endregion


        #region IContextMenu Members

        /// <summary>
        /// Add commands to a shortcut menu.
        /// </summary>
        /// <param name="hMenu">A handle to the shortcut menu.</param>
        /// <param name="iMenu">
        /// The zero-based position at which to insert the first new menu item.
        /// </param>
        /// <param name="idCmdFirst">
        /// The minimum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="idCmdLast">
        /// The maximum value that the handler can specify for a menu item ID.
        /// </param>
        /// <param name="uFlags">
        /// Optional flags that specify how the shortcut menu can be changed.
        /// </param>
        /// <returns>
        /// If successful, returns an HRESULT value that has its severity value set 
        /// to SEVERITY_SUCCESS and its code value set to the offset of the largest 
        /// command identifier that was assigned, plus one.
        /// </returns>
        public int QueryContextMenu(
            IntPtr hMenu,
            uint iMenu,
            uint idCmdFirst,
            uint idCmdLast,
            uint uFlags)
        {
            // If uFlags include CMF_DEFAULTONLY then we should not do anything.
            if (((uint)CMF.CMF_DEFAULTONLY & uFlags) != 0)
            {
                return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0, 0);
            }

            #region sources dropdown
            IntPtr sourceMenu = NativeMethods.CreatePopupMenu();
            var sources = _config.Sources;
            for (uint wId = 0; wId < sources.Count(); wId++)
            {
                // Use either InsertMenu or InsertMenuItem to add menu items.
                MENUITEMINFO mii = new MENUITEMINFO();
                mii.cbSize = (uint)Marshal.SizeOf(mii);
                mii.fMask = MIIM.MIIM_STRING | MIIM.MIIM_FTYPE |
                    MIIM.MIIM_ID | MIIM.MIIM_STATE;
                mii.wID = idCmdFirst + wId + 3;
                mii.fType = MFT.MFT_STRING;
                mii.dwTypeData = sources.ElementAt((int)wId).Key;
                mii.fState = MFS.MFS_ENABLED;
                if (!NativeMethods.InsertMenuItem(sourceMenu, wId, true, ref mii))
                {
                    return Marshal.GetHRForLastWin32Error();
                }
                IDM_DISPLAY = wId + 3;
            }
            #endregion

            #region options dropdown
            var options = _config.Options;
            IntPtr optionMenu = NativeMethods.CreatePopupMenu();
            MENUITEMINFO excludeVersion = new MENUITEMINFO();
            excludeVersion.cbSize = (uint)Marshal.SizeOf(excludeVersion);
            excludeVersion.fMask = MIIM.MIIM_STRING | MIIM.MIIM_FTYPE |
                MIIM.MIIM_ID | MIIM.MIIM_STATE | MIIM.MIIM_CHECKMARKS;
            excludeVersion.wID = idCmdFirst + 0;
            excludeVersion.fType = MFT.MFT_STRING;
            excludeVersion.dwTypeData = "Use versioned packages";
            var versionChecked = options.First(x => x.Key == "useVersionedPackages").Value == "True" ? MFS.MFS_CHECKED : MFS.MFS_UNCHECKED;
            excludeVersion.fState = MFS.MFS_ENABLED | versionChecked;
            if (!NativeMethods.InsertMenuItem(optionMenu, 0, true, ref excludeVersion))
            {
                return Marshal.GetHRForLastWin32Error();
            }

            MENUITEMINFO latest = new MENUITEMINFO();
            latest.cbSize = (uint)Marshal.SizeOf(latest);
            latest.fMask = MIIM.MIIM_STRING | MIIM.MIIM_FTYPE |
                MIIM.MIIM_ID | MIIM.MIIM_STATE | MIIM.MIIM_CHECKMARKS;
            latest.wID = idCmdFirst + 1;
            latest.fType = MFT.MFT_STRING;
            latest.dwTypeData = "Always get latest version";
            var latestChecked = options.First(x => x.Key == "getLatest").Value == "True" ? MFS.MFS_CHECKED : MFS.MFS_UNCHECKED;
            latest.fState = MFS.MFS_ENABLED | latestChecked;
            if (!NativeMethods.InsertMenuItem(optionMenu, 1, true, ref latest))
            {
                return Marshal.GetHRForLastWin32Error();
            }

            MENUITEMINFO clean = new MENUITEMINFO();
            clean.cbSize = (uint)Marshal.SizeOf(clean);
            clean.fMask = MIIM.MIIM_STRING | MIIM.MIIM_FTYPE |
                MIIM.MIIM_ID | MIIM.MIIM_STATE | MIIM.MIIM_CHECKMARKS;
            clean.wID = idCmdFirst + 2;
            clean.fType = MFT.MFT_STRING;
            clean.dwTypeData = "Clean packages directory";
            var cleanChecked = options.First(x => x.Key == "cleanPackages").Value == "True" ? MFS.MFS_CHECKED : MFS.MFS_UNCHECKED;
            clean.fState = MFS.MFS_ENABLED | cleanChecked;
            if (!NativeMethods.InsertMenuItem(optionMenu, 2, true, ref clean))
            {
                return Marshal.GetHRForLastWin32Error();
            }
            #endregion

            IntPtr getMenu = NativeMethods.CreatePopupMenu();
            if (!NativeMethods.InsertMenu(getMenu, 0, (uint)MF.MF_BYPOSITION | (uint)MF.MF_POPUP, (uint)optionMenu.ToInt32(), "Options"))
            {
                return Marshal.GetHRForLastWin32Error();
            }
            if (!NativeMethods.InsertMenu(getMenu, 1, (uint)MF.MF_BYPOSITION | (uint)MF.MF_POPUP, (uint)sourceMenu.ToInt32(), "Sources"))
            {
                return Marshal.GetHRForLastWin32Error();
            }

            MENUITEMINFO parentmii = new MENUITEMINFO();
            parentmii.cbSize = (uint)Marshal.SizeOf(parentmii);
            parentmii.fMask = MIIM.MIIM_BITMAP | MIIM.MIIM_STRING | MIIM.MIIM_FTYPE |
                MIIM.MIIM_ID | MIIM.MIIM_STATE | MIIM.MIIM_SUBMENU;
            parentmii.wID = idCmdFirst;
            parentmii.fType = MFT.MFT_STRING;
            parentmii.hSubMenu = getMenu;
            parentmii.dwTypeData = this.menuText;
            parentmii.fState = MFS.MFS_ENABLED;
            parentmii.hbmpItem = this.menuBmp;
            if (!NativeMethods.InsertMenuItem(hMenu, iMenu, true, ref parentmii))
            {
                return Marshal.GetHRForLastWin32Error();
            }

            // Add a separator.
            MENUITEMINFO sep = new MENUITEMINFO();
            sep.cbSize = (uint)Marshal.SizeOf(sep);
            sep.fMask = MIIM.MIIM_TYPE;
            sep.fType = MFT.MFT_SEPARATOR;
            if (!NativeMethods.InsertMenuItem(hMenu, iMenu + 1, true, ref sep))
            {
                return Marshal.GetHRForLastWin32Error();
            }

            // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
            // Set the code value to the offset of the largest command identifier 
            // that was assigned, plus one (1).
            return WinError.MAKE_HRESULT(WinError.SEVERITY_SUCCESS, 0,
                IDM_DISPLAY + 1);
        }

        /// <summary>
        /// Carry out the command associated with a shortcut menu item.
        /// </summary>
        /// <param name="pici">
        /// A pointer to a CMINVOKECOMMANDINFO or CMINVOKECOMMANDINFOEX structure 
        /// containing information about the command. 
        /// </param>
        public void InvokeCommand(IntPtr pici)
        {
            bool isUnicode = false;

            // Determine which structure is being passed in, CMINVOKECOMMANDINFO or 
            // CMINVOKECOMMANDINFOEX based on the cbSize member of lpcmi. Although 
            // the lpcmi parameter is declared in Shlobj.h as a CMINVOKECOMMANDINFO 
            // structure, in practice it often points to a CMINVOKECOMMANDINFOEX 
            // structure. This struct is an extended version of CMINVOKECOMMANDINFO 
            // and has additional members that allow Unicode strings to be passed.

            try
            {
                CMINVOKECOMMANDINFO ici = (CMINVOKECOMMANDINFO)Marshal.PtrToStructure(
                    pici, typeof(CMINVOKECOMMANDINFO));
                CMINVOKECOMMANDINFOEX iciex = new CMINVOKECOMMANDINFOEX();
                if (ici.cbSize == Marshal.SizeOf(typeof(CMINVOKECOMMANDINFOEX)))
                {
                    if ((ici.fMask & CMIC.CMIC_MASK_UNICODE) != 0)
                    {
                        isUnicode = true;
                        iciex = (CMINVOKECOMMANDINFOEX)Marshal.PtrToStructure(pici,
                            typeof(CMINVOKECOMMANDINFOEX));
                    }
                }
                OnVerbDisplayFileName(ici.hwnd, ici.verb.ToInt32());
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Get information about a shortcut menu command, including the help string 
        /// and the language-independent, or canonical, name for the command.
        /// </summary>
        /// <param name="idCmd">Menu command identifier offset.</param>
        /// <param name="uFlags">
        /// Flags specifying the information to return. This parameter can have one 
        /// of the following values: GCS_HELPTEXTA, GCS_HELPTEXTW, GCS_VALIDATEA, 
        /// GCS_VALIDATEW, GCS_VERBA, GCS_VERBW.
        /// </param>
        /// <param name="pReserved">Reserved. Must be IntPtr.Zero</param>
        /// <param name="pszName">
        /// The address of the buffer to receive the null-terminated string being 
        /// retrieved.
        /// </param>
        /// <param name="cchMax">
        /// Size of the buffer, in characters, to receive the null-terminated string.
        /// </param>
        public void GetCommandString(
            UIntPtr idCmd,
            uint uFlags,
            IntPtr pReserved,
            StringBuilder pszName,
            uint cchMax)
        {
            switch ((GCS)uFlags)
            {
                case GCS.GCS_VERBW:
                    if (this.verbCanonicalName.Length > cchMax - 1)
                    {
                        Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                    }
                    else
                    {
                        pszName.Clear();
                        pszName.Append(this.verbCanonicalName);
                    }
                    break;

                case GCS.GCS_HELPTEXTW:
                    if (this.verbHelpText.Length > cchMax - 1)
                    {
                        Marshal.ThrowExceptionForHR(WinError.STRSAFE_E_INSUFFICIENT_BUFFER);
                    }
                    else
                    {
                        pszName.Clear();
                        pszName.Append(this.verbHelpText);
                    }
                    break;
            }
        }

        #endregion
    }
}