<div align="center">

## VB\-friendly SpecialFolderPath Function


</div>

### Description

This is a VB-friendly interpretation of a reliable method for accessing the myriad of Windows special folder paths in their resolved states. If you need to know the full path to the Desktop Folder, Startup Folder, and many many others, this function will quickly get you there.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kerry B\. Rogers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kerry-b-rogers.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kerry-b-rogers-vb-friendly-specialfolderpath-function__1-39475/archive/master.zip)





### Source Code

```
Private Type SHELLITEMID
  cb As Long
  abID As Byte
End Type
Private Type ITEMIDLIST
  mkid As SHELLITEMID
End Type
Public Enum SpecialFolderTypes
  sftCDBurningCache = 59&
  sftCommonAdminTools = 47&
  sftCommonApplicationData = 35&
  sftCommonDesktop = 25&
  sftCommonDocumentTemplates = 45&
  sftCommonFavorites = 31&
  sftCommonMyDocuments = 46&
  sftCommonMyPictures = 54&
  sftCommonProgramFiles = 43&
  sftCommonStartMenu = 22&
  sftCommonStartMenuPrograms = 23&
  sftCommonStartup = 24&
  sftFonts = 20&
  sftProgramFiles = 38&
  sftSystem32Folder = 41&
  sftSystemFolder = 37&
  sftThemes = 56&
  sftUserAdminTools = 48&
  sftUserApplicationData = 26&
  sftUserCookies = 33&
  sftUserDesktop = 16&
  sftUserDocumentTemplates = 21&
  sftUserFavorites = 6&
  sftUserHistory = 34&
  sftUserLocalApplicationData = 28&
  sftUserMyDocuments = 5&
  sftUserMyMusic = 13&
  sftUserMyPictures = 39&
  sftUserNetHood = 19&
  sftUserPrintHood = 27&
  sftUserProfileFolder = 40&
  sftUserRecentDocuments = 8&
  sftUserSendTo = 9&
  sftUserStartMenu = 11&
  sftUserStartMenuPrograms = 2&
  sftUserStartup = 7&
  sftUserTempInternetFiles = 32&
  sftWindowsFolder = 36&
End Enum
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Function SpecialFolderPath(ByVal lngFolderType As SpecialFolderTypes) As String
  Dim strPath As String
  Dim IDL As ITEMIDLIST
  SpecialFolderPath = ""
  If SHGetSpecialFolderLocation(0&, lngFolderType, IDL) = 0& Then
    strPath = Space$(MAX_PATH)
    If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath) Then
      SpecialFolderPath = Left$(strPath, InStr(strPath, vbNullChar) - 1&) & "\"
    End If
  End If
End Function
```

