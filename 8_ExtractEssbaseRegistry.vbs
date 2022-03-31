'#----------------------------------------------------------------------------------
'#- extractEssbaseRegistry.vbs
'#- Author: Charles Beyer
'#- Date Created: 6/24/2014
'#- Date Last Modified: 6/27/2014
'#- Description: This script will iterate through all users in the HKEY_USERS hive 
'#- and extract any \Software\Hyperion Solutions\Essbase\CSL Global Options branches
'#- Inputs: 
'#-    - [None]
'#- Output: 
'#-    - Registry extract files (one per user) stored in %TEMP%\EssAddInRegistryExtract\
'#-    - Return Code indicating completion status where 0 is complete with no errors
'#- Notes:
'#-    - This script should be run with the wscript engine, not cscript.  
'#----------------------------------------------------------------------------------


OPTION EXPLICIT

'#--------Alter Error Logic to Prevent App Termination------------------------------

Dim boolDebug: boolDebug=0  'Debug Mode Disabled, enable by setting this to 1 

if boolDebug <> 0 then
  On Error Resume Next
else
  On Error Goto 0
end if 

'#------------ Determine Windows Temp Folder --------------------------------
Const TemporaryFolder = 2
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")

Dim tempFolder: tempFolder = fso.GetSpecialFolder(TemporaryFolder)
  If Err.Number <> 0  then Wscript.Quit(Err.Number)
  If boolDebug <> 0 then Msgbox "Temp Path = " & tempFolder

Dim essExtractFolder: essExtractFolder = tempFolder & "\EssAddInRegistryExtract\"
  If Err.Number <> 0  then Wscript.Quit(Err.Number)
  If boolDebug <> 0 then Msgbox "essExtractFolder = " & essExtractFolder

'#------------ Check for existence of Essbase Temp folder and create if needed -------
If NOT (fso.FolderExists(essExtractFolder)) Then
   fso.CreateFolder(essExtractFolder)
       If Err.Number <> 0  then Wscript.Quit(Err.Number)
       If boolDebug <> 0 then Msgbox "Folder Created Successfully (" & essExtractFolder & ")"
end if 

'#-------------- Iterate Registry Users in HKEY_USERS ------------------------
Dim strIPDKeyPath: strIPDKeyPath = "\Software\Hyperion Solutions\Essbase\CSL Global Options"
Dim objRegistry
Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv") 
   If Err.Number <> 0  then Wscript.Quit(Err.Number)
   If boolDebug <> 0 then Msgbox "Registry Reference Acquired"

'#------------------- Grab a list of All User SIDs ----------------------------
Dim arrSubkeys
Const HKEY_USERS = &H80000003
objRegistry.EnumKey HKEY_USERS, "", arrSubkeys
   If Err.Number <> 0  then Wscript.Quit(Err.Number)
   If boolDebug <> 0 then Msgbox "Registry subkeys retrieved"

'#------------------- Loop Through all User SIDs and extract -------------------------
Dim oShell
Set oShell = CreateObject("WScript.Shell")
   If Err.Number <> 0  then Wscript.Quit(Err.Number)
   If boolDebug <> 0 then Msgbox "Shell Object instantiated"

Dim intFileCounter: intFileCounter = 1
Dim sCmd 
Dim sRegistryFile
Dim sKeyBranch
Dim strSubKey

For Each strSubKey in arrSubKeys
  sKeyBranch = "HKEY_USERS\" & strSubkey & strIPDKeyPath
    If boolDebug <> 0 then Msgbox "SubKey Path = [" & sKeyBranch & "]"
  sRegistryFile = essExtractFolder & "EssbaseGlobal_User_" & intFileCounter & ".reg"
    If boolDebug <> 0 then Msgbox "Registry Output File = [" & sRegistryFile & "]"
  sCmd = "regedit.exe /S /E:A """ & sRegistryFile & """ " & """" & sKeyBranch & """"
    If boolDebug <> 0 then Msgbox "Registry Command Line = [" & sCmd & "]"

  'Execute regedit extract command : regedit.exe /S /E:A "File Name" "Key Name"
   oShell.Run sCmd, 0, True
     If Err.Number <> 0  then Wscript.Quit(Err.Number)
     If boolDebug <> 0 then Msgbox "Executed Registry Command Line"

   intFileCounter = intFileCounter + 1

Next

If boolDebug <> 0 then Msgbox "Finished executing script."



