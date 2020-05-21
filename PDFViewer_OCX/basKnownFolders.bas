Attribute VB_Name = "basKnownFolders"
Option Explicit
'© Ellis Dee VB-Forums CodeBank - Visual Basic 6 and earlier

#Const IncludeVirtualFolders = False
#Const IncludeDebugListing = True

Public Enum KnownFolderEnum
    kfUserProfiles
    kfUser
    kfUserDocuments
    kfUserContacts
    kfUserDesktop
    kfUserDownloads
    kfUserMusic
    kfUserPictures
    kfUserSavedGames
    kfUserVideos
    kfUserAppDataRoaming
    kfUserAppDataLocal
    kfUserAppDataLocalLow
    kfUserCDBurning
    kfUserCookies
    kfUserFavorites
    kfUserGameTasks
    kfUserHistory
    kfUserInternetCache
    kfUserLinks
    kfUserNetHood
    kfUserPrintHood
    kfUserQuickLaunch
    kfUserRecent
    kfUserSavedSearches
    kfUserSendTo
    kfUserStartMenu
    kfUserStartMenuAdminTools
    kfUserStartMenuPrograms
    kfUserStartMenuStartup
    kfUserTemplates
    kfPublic
    kfPublicDesktop
    kfPublicDocuments
    kfPublicDownloads
    kfPublicMusic
    kfPublicPictures
    kfPublicVideos
    kfPublicStartMenu
    kfPublicStartMenuAdminTools
    kfPublicStartMenuPrograms
    kfPublicStartMenuStartup
    kfPublicGameTasks
    kfPublicTemplates
    kfProgramData
    kfWindows
    kfSystem
    kfSystemX86
    kfSystemFonts
    kfSystemResourceDir
    kfProgramFilesX86
    kfProgramFilesCommonX86
    kfProgramFiles
    kfProgramFilesCommon
#If IncludeVirtualFolders = True Then
    kfAddNewPrograms
    kfAppUpdates
    kfChangeRemovePrograms
    kfCommonOEMLinks
    kfComputerFolder
    kfConflictFolder
    kfConnectionsFolder
    kfControlPanelFolder
    kfGames
    kfInternetFolder
    kfLocalizedResourcesDir
    kfNetworkFolder
    kfOriginalImages
    kfPhotoAlbums
    kfPlaylists
    kfPrintersFolder
    kfProgramFilesX64
    kfProgramFilesCommonX64
    kfRecordedTV
    kfRecycleBinFolder
    kfSampleMusic
    kfSamplePictures
    kfSamplePlaylists
    kfSampleVideos
    kfSEARCH_CSC
    kfSEARCH_MAPI
    kfSearchHome
    kfSidebarDefaultParts
    kfSidebarParts
    kfSyncManagerFolder
    kfSyncResultsFolder
    kfSyncSetupFolder
    kfTreeProperties
    kfUsersFiles
#End If
#If IncludeDebugListing = True Then
    kfKnownFolders
#End If
End Enum

Private Type GUIDType
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
    
Private Declare Function SHGetKnownFolderPath Lib "shell32" (rfid As Any, ByVal dwFlags As Long, ByVal hToken As Long, ppszPath As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal ptr As Long) As Long

Public Function KnownFolder(penKnownFolder As KnownFolderEnum) As String
    Dim strGUID As String
    Dim typGUID As GUIDType
    Dim lngPath As Long
    Dim bytArray() As Byte
    Dim lngBytes As Long
    
    strGUID = KnownFolderGUID(penKnownFolder)
    If Len(strGUID) = 0 Then Exit Function
    If CLSIDFromString(StrPtr(strGUID), typGUID) = 0 Then
        If SHGetKnownFolderPath(typGUID, 0, 0, lngPath) = 0 Then
            If lngPath <> 0 Then
                lngBytes = lstrlenW(ByVal lngPath) * 2
                If lngBytes <> 0 Then
                    ReDim bytArray(0 To lngBytes - 1)
                    CopyMemory bytArray(0), ByVal lngPath, lngBytes
                    KnownFolder = bytArray
                End If
            End If
            Call CoTaskMemFree(lngPath)
        End If
    End If
End Function

Public Function KnownFolderGUID(penKnownFolder As KnownFolderEnum) As String
    Dim strReturn As String
    
    Select Case penKnownFolder
        Case kfUserProfiles: strReturn = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
        Case kfUser: strReturn = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
        Case kfUserDocuments: strReturn = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
        Case kfUserContacts: strReturn = "{56784854-C6CB-462b-8169-88E350ACB882}"
        Case kfUserDesktop: strReturn = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
        Case kfUserDownloads: strReturn = "{374DE290-123F-4565-9164-39C4925E467B}"
        Case kfUserMusic: strReturn = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
        Case kfUserPictures: strReturn = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
        Case kfUserSavedGames: strReturn = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
        Case kfUserVideos: strReturn = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
        Case kfUserAppDataRoaming: strReturn = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
        Case kfUserAppDataLocal: strReturn = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
        Case kfUserAppDataLocalLow: strReturn = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
        Case kfUserCDBurning: strReturn = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
        Case kfUserCookies: strReturn = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
        Case kfUserFavorites: strReturn = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
        Case kfUserGameTasks: strReturn = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
        Case kfUserHistory: strReturn = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
        Case kfUserInternetCache: strReturn = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
        Case kfUserLinks: strReturn = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
        Case kfUserNetHood: strReturn = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
        Case kfUserPrintHood: strReturn = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
        Case kfUserQuickLaunch: strReturn = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
        Case kfUserRecent: strReturn = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
        Case kfUserSavedSearches: strReturn = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
        Case kfUserSendTo: strReturn = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
        Case kfUserStartMenu: strReturn = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
        Case kfUserStartMenuAdminTools: strReturn = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
        Case kfUserStartMenuPrograms: strReturn = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
        Case kfUserStartMenuStartup: strReturn = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
        Case kfUserTemplates: strReturn = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
        Case kfPublic: strReturn = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
        Case kfPublicDesktop: strReturn = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
        Case kfPublicDocuments: strReturn = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
        Case kfPublicDownloads: strReturn = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
        Case kfPublicMusic: strReturn = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
        Case kfPublicPictures: strReturn = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
        Case kfPublicVideos: strReturn = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
        Case kfPublicStartMenu: strReturn = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
        Case kfPublicStartMenuAdminTools: strReturn = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
        Case kfPublicStartMenuPrograms: strReturn = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
        Case kfPublicStartMenuStartup: strReturn = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
        Case kfPublicGameTasks: strReturn = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
        Case kfPublicTemplates: strReturn = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
        Case kfProgramData: strReturn = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
        Case kfWindows: strReturn = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
        Case kfSystem: strReturn = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
        Case kfSystemX86: strReturn = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
        Case kfSystemFonts: strReturn = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
        Case kfSystemResourceDir: strReturn = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
        Case kfProgramFilesX86: strReturn = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
        Case kfProgramFilesCommonX86: strReturn = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
        Case kfProgramFiles: strReturn = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
        Case kfProgramFilesCommon: strReturn = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
#If IncludeVirtualFolders Then
        Case kfAddNewPrograms: strReturn = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
        Case kfAppUpdates: strReturn = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
        Case kfChangeRemovePrograms: strReturn = "{df7266ac-9274-4867-8d55-3bd661de872d}"
        Case kfCommonOEMLinks: strReturn = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
        Case kfComputerFolder: strReturn = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
        Case kfConflictFolder: strReturn = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
        Case kfConnectionsFolder: strReturn = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
        Case kfControlPanelFolder: strReturn = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
        Case kfGames: strReturn = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
        Case kfInternetFolder: strReturn = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
        Case kfLocalizedResourcesDir: strReturn = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
        Case kfNetworkFolder: strReturn = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
        Case kfOriginalImages: strReturn = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
        Case kfPhotoAlbums: strReturn = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
        Case kfPlaylists: strReturn = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
        Case kfPrintersFolder: strReturn = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
        Case kfProgramFilesX64: strReturn = "{6D809377-6AF0-444b-8957-A3773F02200E}"
        Case kfProgramFilesCommonX64: strReturn = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
        Case kfRecordedTV: strReturn = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
        Case kfRecycleBinFolder: strReturn = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
        Case kfSampleMusic: strReturn = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
        Case kfSamplePictures: strReturn = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
        Case kfSamplePlaylists: strReturn = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
        Case kfSampleVideos: strReturn = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
        Case kfSEARCH_CSC: strReturn = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
        Case kfSEARCH_MAPI: strReturn = "{98ec0e18-2098-4d44-8644-66979315a281}"
        Case kfSearchHome: strReturn = "{190337d1-b8ca-4121-a639-6d472d16972a}"
        Case kfSidebarDefaultParts: strReturn = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
        Case kfSidebarParts: strReturn = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
        Case kfSyncManagerFolder: strReturn = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
        Case kfSyncResultsFolder: strReturn = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
        Case kfSyncSetupFolder: strReturn = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
        Case kfTreeProperties: strReturn = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
        Case kfUsersFiles: strReturn = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
#End If
    End Select
    KnownFolderGUID = strReturn
End Function

#If IncludeDebugListing = True Then
Public Function KnownFolderList()
    Dim i As Long
    
    For i = 0 To kfKnownFolders - 1
        Debug.Print KnownFolderName(i) & ": " & KnownFolder(i)
    Next
End Function

Private Function KnownFolderName(penKnownFolder As KnownFolderEnum) As String
    Dim strReturn As String
    
    Select Case penKnownFolder
        Case kfUserProfiles: strReturn = "UserProfiles"
        Case kfUser: strReturn = "User"
        Case kfUserDocuments: strReturn = "UserDocuments"
        Case kfUserContacts: strReturn = "UserContacts"
        Case kfUserDesktop: strReturn = "UserDesktop"
        Case kfUserDownloads: strReturn = "UserDownloads"
        Case kfUserMusic: strReturn = "UserMusic"
        Case kfUserPictures: strReturn = "UserPictures"
        Case kfUserSavedGames: strReturn = "UserSavedGames"
        Case kfUserVideos: strReturn = "UserVideos"
        Case kfUserAppDataRoaming: strReturn = "UserAppDataRoaming"
        Case kfUserAppDataLocal: strReturn = "UserAppDataLocal"
        Case kfUserAppDataLocalLow: strReturn = "UserAppDataLocalLow"
        Case kfUserCDBurning: strReturn = "UserCDBurning"
        Case kfUserCookies: strReturn = "UserCookies"
        Case kfUserFavorites: strReturn = "UserFavorites"
        Case kfUserGameTasks: strReturn = "UserGameTasks"
        Case kfUserHistory: strReturn = "UserHistory"
        Case kfUserInternetCache: strReturn = "UserInternetCache"
        Case kfUserLinks: strReturn = "UserLinks"
        Case kfUserNetHood: strReturn = "UserNetHood"
        Case kfUserPrintHood: strReturn = "UserPrintHood"
        Case kfUserQuickLaunch: strReturn = "UserQuickLaunch"
        Case kfUserRecent: strReturn = "UserRecent"
        Case kfUserSavedSearches: strReturn = "UserSavedSearches"
        Case kfUserSendTo: strReturn = "UserSendTo"
        Case kfUserStartMenu: strReturn = "UserStartMenu"
        Case kfUserStartMenuAdminTools: strReturn = "StartMenuAdminTools"
        Case kfUserStartMenuPrograms: strReturn = "UserStartMenuPrograms"
        Case kfUserStartMenuStartup: strReturn = "UserStartMenuStartup"
        Case kfUserTemplates: strReturn = "UserTemplates"
        Case kfPublic: strReturn = "Public"
        Case kfPublicDesktop: strReturn = "PublicDesktop"
        Case kfPublicDocuments: strReturn = "PublicDocuments"
        Case kfPublicDownloads: strReturn = "PublicDownloads"
        Case kfPublicMusic: strReturn = "PublicMusic"
        Case kfPublicPictures: strReturn = "PublicPictures"
        Case kfPublicVideos: strReturn = "PublicVideos"
        Case kfPublicStartMenu: strReturn = "PublicStartMenu"
        Case kfPublicStartMenuAdminTools: strReturn = "PublicStartMenuAdminTools"
        Case kfPublicStartMenuPrograms: strReturn = "PublicStartMenuPrograms"
        Case kfPublicStartMenuStartup: strReturn = "PublicStartMenuStartup"
        Case kfPublicGameTasks: strReturn = "PublicGameTasks"
        Case kfPublicTemplates: strReturn = "PublicTemplates"
        Case kfProgramData: strReturn = "ProgramData"
        Case kfWindows: strReturn = "Windows"
        Case kfSystem: strReturn = "System"
        Case kfSystemX86: strReturn = "SystemX86"
        Case kfSystemFonts: strReturn = "SystemFonts"
        Case kfSystemResourceDir: strReturn = "SystemResourceDir"
        Case kfProgramFilesX86: strReturn = "ProgramFilesX86"
        Case kfProgramFilesCommonX86: strReturn = "ProgramFilesCommonX86"
        Case kfProgramFiles: strReturn = "ProgramFiles"
        Case kfProgramFilesCommon: strReturn = "ProgramFilesCommon"
#If IncludeVirtualFolders Then
        Case kfAddNewPrograms: strReturn = "AddNewPrograms"
        Case kfAppUpdates: strReturn = "AppUpdates"
        Case kfChangeRemovePrograms: strReturn = "ChangeRemovePrograms"
        Case kfCommonOEMLinks: strReturn = "CommonOEMLinks"
        Case kfComputerFolder: strReturn = "ComputerFolder"
        Case kfConflictFolder: strReturn = "ConflictFolder"
        Case kfConnectionsFolder: strReturn = "ConnectionsFolder"
        Case kfControlPanelFolder: strReturn = "ControlPanelFolder"
        Case kfGames: strReturn = "Games"
        Case kfInternetFolder: strReturn = "InternetFolder"
        Case kfLocalizedResourcesDir: strReturn = "LocalizedResourcesDir"
        Case kfNetworkFolder: strReturn = "NetworkFolder"
        Case kfOriginalImages: strReturn = "OriginalImages"
        Case kfPhotoAlbums: strReturn = "PhotoAlbums"
        Case kfPlaylists: strReturn = "Playlists"
        Case kfPrintersFolder: strReturn = "PrintersFolder"
        Case kfProgramFilesX64: strReturn = "ProgramFilesX64"
        Case kfProgramFilesCommonX64: strReturn = "ProgramFilesCommonX64"
        Case kfRecordedTV: strReturn = "RecordedTV"
        Case kfRecycleBinFolder: strReturn = "RecycleBinFolder"
        Case kfSampleMusic: strReturn = "SampleMusic"
        Case kfSamplePictures: strReturn = "SamplePictures"
        Case kfSamplePlaylists: strReturn = "SamplePlaylists"
        Case kfSampleVideos: strReturn = "SampleVideos"
        Case kfSEARCH_CSC: strReturn = "SEARCH_CSC"
        Case kfSEARCH_MAPI: strReturn = "SEARCH_MAPI"
        Case kfSearchHome: strReturn = "SearchHome"
        Case kfSidebarDefaultParts: strReturn = "SidebarDefaultParts"
        Case kfSidebarParts: strReturn = "SidebarParts"
        Case kfSyncManagerFolder: strReturn = "SyncManagerFolder"
        Case kfSyncResultsFolder: strReturn = "SyncResultsFolder"
        Case kfSyncSetupFolder: strReturn = "SyncSetupFolder"
        Case kfTreeProperties: strReturn = "TreeProperties"
        Case kfUsersFiles: strReturn = "UsersFiles"
#End If
    End Select
    KnownFolderName = strReturn
End Function
#End If
