
' encoding=gb18030

On Error Resume Next

Function GetFileDirectory(filePath)
    Dim objFSO
    Dim fileDirectory

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    fileDirectory = objFSO.GetParentFolderName(filePath)

    Set objFSO = Nothing

    GetFileDirectory = fileDirectory
End Function


Sub DeleteFile(filePath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If objFSO.FileExists(filePath) Then
        objFSO.DeleteFile(filePath)
    End If

    Set objFSO = Nothing
End Sub


Sub CopyFile(sourceFilePath, destFilePath)
    Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(sourceFilePath) Then
        objFSO.CopyFile sourceFilePath, destFilePath, true
    End If
    Set objFSO = Nothing

End Sub


Function CreatePath(path)
    Dim objFSO
    Dim parentPath

    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    parentPath = objFSO.GetParentFolderName(path)
    If Not objFSO.FolderExists(parentPath) Then
        CreatePath(parentPath)
    End If
    objFSO.CreateFolder path

    CreatePath = objFSO.FolderExists(path)
End Function


Function CreateParentPath(filePath)
    Dim objFSO
    Dim parentPath

    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    parentPath = objFSO.GetParentFolderName(filePath)
    If Not objFSO.FolderExists(parentPath) Then
        CreatePath(parentPath)
    End If
End Function


Function GetImageUrlList(xmlUrl)
    Dim objXML, imageUrlList: imageUrlList = Array()

    Set objXML = CreateObject("Msxml2.DOMDocument")
    objXML.setProperty "ServerHTTPRequest", True
    objXML.async =  False
    objXML.Load xmlUrl

    If objXML.parseError.errorCode = 0 Then
        Set us = objXML.documentElement.getElementsByTagName("urlBase")
        Redim imageUrlList(us.length-1)
        For i = 0 To us.length-1
            imageUrlList(i) = "http://www.bing.com" & us(i).text & "_1920x1200.jpg"
        Next

        xmlLocalPath = GetScriptDirectory() & "\bing.xml"
        objXML.Save(xmlLocalPath)

        Set us = Nothing
    'Else
    '    MsgBox "Request XML errorCode:" & objXML.parseError.errorCode
    End If

    Set objXML = Nothing

    GetImageUrlList = imageUrlList
End Function


Function GetScriptDirectory
    GetScriptDirectory = GetFileDirectory(Wscript.ScriptFullName)
End Function


Function GetImageDirectory
    Dim scriptDirectory
    scriptDirectory = GetScriptDirectory()
    GetImageDirectory = scriptDirectory & "\wallpapers"
End Function


Function GetImage(imageUrlList)
    Dim imageUrl, imagePath, imageLocalFullName
    imagePath = GetImageDirectory()

    If UBound(imageUrlList) > 0 Then
        imageUrl = imageUrlList(0)
        imageLocalFullName = Wget(imageUrl, imagePath)
    End If

    GetImage = imageLocalFullName
End Function


Function Wget(imageUrl, imagePath)
    Dim imageFileName, imageFullPath, objFSO, objXMLHTTP, retries
    imageFileName = Mid(imageUrl, InStrRev(imageUrl, "/")+1)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(imagePath) Then
        CreatePath(imagePath)
    End If

    imageFullPath = imagePath & "\" & imageFileName

    If Not objFSO.FileExists(imageFullPath) Then
        Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
        retries = 3
        Do
            objXMLHTTP.open "GET", imageUrl, false
            objXMLHTTP.send()

            If objXMLHTTP.Status = 200 Then
                Set objStream = CreateObject("ADODB.Stream")
                objStream.Open
                objStream.Type = 1
                objStream.Write objXMLHTTP.ResponseBody
                objStream.SaveToFile imageFullPath
                objStream.Close

                Set objStream = Nothing

                Exit Do
            End If

            retries = retires - 1
        Loop While retries > 0

        Set objXMLHTTP = Nothing
    End If

    Set objFSO = Nothing

    Wget = imageFullPath
End Function


Function GetOSLanguage()
    Dim osLanguageStr, uiLanguages

    Set objShell = WScript.CreateObject("WScript.Shell")

    ' Detect language for windows 7 with multiple language packs installed
    If IsEmpty(osLanguageStr) Then
        uiLanguages = objShell.RegRead("HKCU\Control Panel\Desktop\PreferredUILanguages")

        If UBound(uiLanguages) >= 0 Then
            Select Case uiLanguages(0)
                Case "en-US"
                    osLanguageStr = "en-US"
                Case "zh-CN"
                    osLanguageStr = "zh-CN"
            End Select
        End If
    End If

    ' Detect language using WMI
    If IsEmpty(osLanguageStr) Then
        ' TODO
    End If

    Set objShell = Nothing

    GetOSLanguage = osLanguageStr
End Function


Function GetSetBackgroundStr()
    Dim objShell
    Dim setBackgroundStr, osLanguageStr

    setBackgroundStr = "Set as desktop background"
    osLanguageStr = GetOSLanguage()

    Select Case osLanguageStr
        Case "en-US"
            setBackgroundStr = "Set as desktop background"
        Case "zh-CN"
            setBackgroundStr = "…Ë÷√Œ™◊¿√Ê±≥æ∞"
    End Select

    GetSetBackgroundStr = setBackgroundStr
End Function


Sub SetWallpaper(imageFullPath)
    Dim objFSO, objFile, imagePath, imageName, setBackgroundStr
    Dim objShell, objFolder, objFolderItem, colVerbs

    If IsNull(imageFullPath) Or IsEmpty(imageFullPath) Or imageFullPath = "" Then
        Exit Sub
    End If

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.GetFile(imageFullPath)
    imagePath = objFile.ParentFolder
    imageName = objFile.Name

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.NameSpace(imagePath)
    Set objFolderItem = objFolder.ParseName(imageName)
    Set colVerbs = objFolderItem.Verbs

    setBackgroundStr = GetSetBackgroundStr()

    For Each objVerb in colVerbs
        If InStr(Replace(objVerb, "&", ""), setBackgroundStr) <> 0 Then
            objVerb.DoIt
            wscript.sleep(2000)
            Exit For
        End If
    Next

    Set colVerbs = Nothing
    Set objFolderItem = Nothing
    Set objFolder = Nothing
    Set objShell = Nothing

    Set objFile = Nothing
    Set objFSO = Nothing
End Sub


Function GetFileSize(filePath)
    Dim objFSO, objFile
    Dim fileSize

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(filePath) Then
        Set objFile = objFSO.GetFile(filePath)
        fileSize = objFile.Size
        Set objFile = Nothing
    End If

    Set objFSO = Nothing

    GetFileSize = fileSize
End Function


Sub CompressJPEG(originalImagePath, newImagePath, maxImageSize)
    Dim objImage, objIP
    Dim objNewImage

    If GetFileSize(originalImagePath) <= maxImageSize Then
        CopyFile originalImagePath, newImagePath
        Exit Sub
    End If

    Set objImage = CreateObject("WIA.ImageFile")
    Set objIP = CreateObject("WIA.ImageProcess")

    objIP.Filters.Add objIP.FilterInfos("Convert").FilterID
    'objIP.Filters(1).Properties("FormatID").Value = wiaFormatJPEG
    objIP.Filters(1).Properties("FormatID").Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"

    quality=90
    Do While quality > 0
        objImage.LoadFile originalImagePath
        objIP.Filters(1).Properties("Quality").Value = quality

        Set objNewImage = objIP.Apply(objImage)

        DeleteFile(newImagePath)
        objNewImage.SaveFile newImagePath

        If GetFileSize(newImagePath) <= maxImageSize Then
            Exit Do
        End If

        If quality > 10 Then
            quality = quality - 10
        Else
            quality = quality - 3
        End If
    Loop

    Set objIP = Nothing
    Set objImage = Nothing

End Sub


Function GetOEMImagePath()
    Dim objFSO
    Dim oemImagePath

    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    oemImagePath = objFSO.GetSpecialFolder(SystemFolder) & "\System32\oobe\Info\backgrounds\backgroundDefault.jpg"
    Set objFSO = Nothing

    GetOEMImagePath = oemImagePath
End Function


Sub EnableOobeBackground()
    Dim objShell
    Set objShell = WScript.CreateObject("WScript.Shell")
    objShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background\OEMBackground", "1", "REG_DWORD"
    Set objShell = Nothing
End Sub


Sub SetLogonBackground(imagePath)
    Dim logonBackgroundPath, oemImagePath
    Const MAX_IMAGE_SIZE = 256000

    If IsNull(imagePath) Or IsEmpty(imagePath) Or imagePath = "" Then
        Exit Sub
    End If

    oemImagePath = GetOEMImagePath()
    CreateParentPath(oemImagePath)
    logonBackgroundPath = GetImageDirectory() & "\logonBackground.jpg"
    CompressJPEG imagePath, logonBackgroundPath, MAX_IMAGE_SIZE
    DeleteFile(oemImagePath)
    CopyFile logonBackgroundPath, oemImagePath
    EnableOobeBackground
End Sub


xmlUrl = "http://www.bing.com/hpimagearchive.aspx?format=xml&idx=0&n=23&mbl=1&mkt=en-ww"
imageUrlList = GetImageUrlList(xmlUrl)
imageLocalPath = GetImage(imageUrlList)
SetWallpaper imageLocalPath
SetLogonBackground imageLocalPath


' vim: st=4 sw=4 sta et
