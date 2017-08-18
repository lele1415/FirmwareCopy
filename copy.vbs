Option Explicit

Sub Window_OnLoad
Dim width,height
width = CreateObject("HtmlFile").ParentWindow.Screen.AvailWidth
Window.MoveTo width-400,100
Window.ResizeTo 400,750
End Sub

Dim ws, Fso
set ws = CreateObject("wscript.shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim allCodePathTxt, allCodePathListId
allCodePathTxt = "allCodePath.txt"
allCodePathListId = "allCodePathList_id"

Dim uTargetFolderPath, uTargetDbFolderPath, uTargetOtaFolderPath, uCodeOutPath, uCodeOutProjectPath, sNewFolderName, sProjectName, uScatterFilePath

Dim aAllNeedFilesName

Dim bNotCopySoftware, bContinue, bTargetFolderFileExist, bNewFolder, bCopyOta

Dim sKK_AP, sKK_BP, sL1_AP, sL1_BP
sKK_AP = "obj\CODEGEN\cgen\"
sKK_BP = "obj\CUSTGEN\custom\modem\"
sL1_AP = "obj\CGEN\"
sL1_BP = "obj\ETC\"

Dim sDbFolderPath_AP, aDbFolderPath_BP(), sDbFilePath_AP, aDbFilePath_BP()
Dim count_BPFolder, count_BPFile
count_BPFolder = -1
count_BPFile = -1
Dim otaFilePath, sOtaFileName_tf, sOtaFileName_ota, sOtaFileName_objtf

ReadCodePath allCodePathTxt, allCodePathListId


Function swapTwoStrings(s1, s2)
    Dim sTmp
    sTmp = s1
    s1 = s2
    s2 = sTmp
End Function

Function sortStringArray(aArray)
    Dim aTmp, i, j, lenArray
    aTmp = aArray
    lenArray = UBound(aTmp)
    If lenArray > 0 Then
        For i = 0 to lenArray - 1
            For j = i + 1 to lenArray
                If StrComp(aTmp(i), aTmp(j), 1) > 0 Then
                    swapTwoStrings aTmp(i), aTmp(j)
                End If
            Next
        Next
    End If
    sortStringArray = aTmp
End Function

Function checkIsExistInArray(sStr, aArray)
    Dim flag, i
    flag = False
    For i = 0 To UBound(aArray)
        If aArray(i) = sStr Then
            flag = True
            Exit For
        End If
    Next
    checkIsExistInArray = flag
End Function

Function getNeedFilesName(uPath)
    Dim oTxt, sReadLine, sTmpFileName, aFilesName(), iSeq, iInFileName, iInNone
    Set oTxt = Fso.OpenTextFile(uPath, 1)
    ReDim Preserve aFilesName(0)
    iSeq = 0
    Do Until oTxt.AtEndOfStream
        sReadLine = oTxt.ReadLine
        iInFileName = InStr(sReadLine, "file_name")
        iInNone = InStr(sReadLine, "NONE")
        If iInFileName > 0 And iInNone = 0 Then
            sTmpFileName = Mid(sReadLine, iInFileName + 11)
            If Not checkIsExistInArray(sTmpFileName, aFilesName) Then
                ReDim Preserve aFilesName(iSeq)
                aFilesName(iSeq) = Mid(sReadLine, iInFileName + 11)
                iSeq = iSeq + 1
            End If
        End If
    Loop
    getNeedFilesName = sortStringArray(aFilesName)
End Function


'//////////////////////////////////////////////////////////////////

Function getCheckedRadio(name)
    Dim radioObj, i
    Set radioObj = document.getElementsByName(name)
    For i = 0 To radioObj.length
        If radioObj(i).checked Then
            getCheckedRadio =  radioObj(i).value
            Exit For
        End If
    Next
End Function

Sub Sleep(MSecs)  
    Dim fso
    Dim objOutputFile

    Set fso = CreateObject("Scripting.FileSystemObject") 
    If Fso.FileExists("sleeper.vbs") = False Then 
        Set objOutputFile = fso.CreateTextFile("sleeper.vbs", True) 
        objOutputFile.Write "wscript.sleep WScript.Arguments(0)" 
        objOutputFile.Close 
    End If 
    CreateObject("WScript.Shell").Run "sleeper.vbs " & MSecs,1 , True 
End Sub

Function getCustomBuildVersion(which)
    initBoolean()
    getInputValue()
    If Fso.FolderExists(uCodeOutPath) Then
        getProjectName()
        Dim oTxt, sTmp
        Set oTxt = fso.opentextfile(uCodeOutProjectPath & "system\build.prop", 1)
        Do Until oTxt.AtEndOfStream
            sTmp = oTxt.ReadLine
            If which = "display.id" Then
                If InStr(sTmp,"ro.build.display.id") > 0 Then
                    document.getElementById("folder_name").value = Trim(Mid(sTmp,InStr(sTmp,"=")+1))
                    Exit Do
                End If
            Else
                If InStr(sTmp,"ro.custom.build.version") > 0 Then
                    document.getElementById("folder_name").value = Trim(Mid(sTmp,InStr(sTmp,"=")+1))
                    Exit Do
                End If
            End If
        Loop
    Else
        msgbox("代码路径不存在out目录，请确认后重试")
    End If

End Function

Function runCopyFile(uCopyFilePath, uTargetFolderPath)
    If Fso.FileExists(uCopyFilePath) Then
        Fso.copyfile uCopyFilePath, uTargetFolderPath, false
    Else
       MsgBox(uCopyFilePath & " is not exist!")
    End If
End Function


'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////

Function ReadCodePath(codePathTxtName, codePathListId)
    Dim txtPath, oTxt, codePathCount, sTmp, continue
    txtPath = ws.CurrentDirectory & "\" & codePathTxtName
    codePathCount = 0
    continue = True
    If Fso.FileExists(txtPath) Then
        Set oTxt = fso.OpenTextFile(txtPath,1)
    Else
        MsgBox("代码路径文件不存在！")
        continue = False
    End If

    If continue Then
        Do Until oTxt.AtEndOfStream
            sTmp = oTxt.ReadLine
            If sTmp <> "" Then
                addOption codePathListId, sTmp
            End If
        Loop
    End If
End Function

Function cleanInputPath()
    document.getElementById("codePathFromInput").value = ""
End Function

Function initBoolean()
    bNotCopySoftware = False
    bContinue = True
    bTargetFolderFileExist = False
    bNewFolder = True
    bCopyOta = False
    count_BPFolder = -1
    count_BPFile = -1
    ReDim Preserve aDbFolderPath_BP(0)
    ReDim Preserve aDbFilePath_BP(0)
End Function

Function getInputValue()
    Dim sCodePathInput
    uTargetFolderPath = document.getElementById("destination_link").value & "\"
    sCodePathInput = document.getElementById("codePathFromInput").value

    If sCodePathInput <> "" Then
        uCodeOutPath = sCodePathInput
    Else
        uCodeOutPath = document.getElementById(allCodePathListId).value & "\out"
    End If

    If getCheckedRadio("new_folder") = "new_folder" Then  
        sNewFolderName = document.getElementById("folder_name").value
        bNewFolder = True
    Else
        bNewFolder = False
    End If
End Function

Function checkInputInfo()

    '/////////////检测输入路径的环境/////////////////////////////////////////////////////

    If (Not Fso.FolderExists(uCodeOutPath)) Then
        msgbox("代码路径不存在out目录，请确认后重试")
        bContinue = False
    ElseIf (Not Fso.FolderExists(uTargetFolderPath)) Then
        msgbox("目标路径不存在，请确认后重试")
        bContinue = False
    ElseIf bNewFolder Then
        If sNewFolderName = "" Then
            msgbox("请输入文件夹名")
            bContinue = False
        Else
            uTargetFolderPath = uTargetFolderPath & sNewFolderName & "\"
            If Fso.FolderExists(uTargetFolderPath) Then
                msgbox("输入的文件夹名与现有文件夹重名，请修改后重试")
                bContinue = False
            Else    
                Fso.CreateFolder(uTargetFolderPath) 
            End If
        End If  
    Else
        Dim oFolder, oSubFolders, Folder, oFiles, File
        Set oFolder = Fso.GetFolder(uTargetFolderPath)
        Set oFiles = oFolder.Files
        Set oSubFolders = oFolder.SubFolders
        bTargetFolderFileExist = False
        For Each File in oFiles
            msgbox("目标路径已存在文件，请删除后重试")
            bContinue = False
            bTargetFolderFileExist = True
            Exit For
        Next
        If bTargetFolderFileExist Then
            For Each Folder in oSubFolders
                msgbox("目标路径已存在文件夹，请删除后重试")             
                bContinue = False
                Exit For
            Next
        End If    
    End If
End Function

Function getProjectName()
    '/////////////获取工程名/////////////////////////////////////////////////////

    If bContinue Then
        Dim oFolder, oSubFolders, Folder, sTmp
        Set oFolder = Fso.GetFolder(uCodeOutPath & "\target\product\")
        Set oSubFolders = oFolder.SubFolders
        For Each Folder in oSubFolders
            sTmp = Folder.Name
            If Fso.FileExists(uCodeOutPath & "\target\product\" & sTmp & "\system.img") Then
                sProjectName = sTmp
                Exit For
            End If
        Next

        If sProjectName = "" Then
            bContinue = False
            msgbox("代码路径下不存在system.img，确认后重试")
        Else
            uCodeOutProjectPath = uCodeOutPath & "\target\product\" & sProjectName & "\"
        End If
    End If
End Function

Function checkDbPath()
    If bContinue Then

    '/////////////检测DB文件路径/////////////////////////////////////////////////////  

        If Fso.FolderExists(uCodeOutProjectPath & sKK_AP) And _
                Fso.FolderExists(uCodeOutProjectPath & sKK_BP) Then
            sDbFolderPath_AP = sKK_AP
            count_BPFolder = count_BPFolder + 1
            ReDim Preserve aDbFolderPath_BP(count_BPFolder)
            aDbFolderPath_BP(count_BPFolder) = sKK_BP
        Else
            If Fso.FolderExists(uCodeOutProjectPath & sL1_AP) Then
                sDbFolderPath_AP = sL1_AP
            End If

            If Fso.FolderExists(uCodeOutProjectPath & sL1_BP) Then
                Dim oFolder, oSubFolders, Folder, sTmp
                Set oFolder = Fso.GetFolder(uCodeOutProjectPath & sL1_BP)
                Set oSubFolders = oFolder.SubFolders
                For Each Folder in oSubFolders
                    sTmp = Folder.Name
                    If InStr(sTmp, "BPLGU") > 0 Then
                        count_BPFolder = count_BPFolder + 1
                        ReDim Preserve aDbFolderPath_BP(count_BPFolder)
                        aDbFolderPath_BP(count_BPFolder) = sL1_BP & sTmp & "\"
                    End If
                Next
            End If

        End If

        If sDbFolderPath_AP = "" And _
                count_BPFolder = -1 Then
            bContinue = False
            MsgBox("未找到DB文件")
        End If

    End If
End Function

Function getDbFilePath()
    If bContinue Then

    '/////////////获取DB文件名/////////////////////////////////////////////////////

        Dim oFolder, oFiles, File, sTmp
        If sDbFolderPath_AP <> "" Then
            Set oFolder = Fso.GetFolder(uCodeOutProjectPath & sDbFolderPath_AP)
            Set oFiles = oFolder.Files
            For Each File in oFiles
                sTmp = File.Name
                If Fso.FileExists(uCodeOutProjectPath & sDbFolderPath_AP & sTmp & "_ENUM") Then
                    sDbFilePath_AP = sDbFolderPath_AP & sTmp    
                    Exit For
                End If
            Next
        End If   

        If count_BPFolder > -1 Then
            Dim i
            For i = 0 To count_BPFolder
                Set oFolder = Fso.GetFolder(uCodeOutProjectPath & aDbFolderPath_BP(i))
                Set oFiles = oFolder.Files
                For Each File in oFiles
                    sTmp = File.Name
                    If InStr(sTmp, "BPLGU") > 0 Then
                        count_BPFile = count_BPFile + 1
                        ReDim Preserve aDbFilePath_BP(count_BPFile)
                        aDbFilePath_BP(count_BPFile) = aDbFolderPath_BP(i) & sTmp
                        Exit For
                    End If
                Next    
            Next    
        End If

        If sDbFilePath_AP = "" And _
                count_BPFile = -1 Then
            bContinue = False
            MsgBox("未找到DB文件")
        End If

    End If
End Function

Function copyDbFile()
    If bContinue Then
        uTargetDbFolderPath = uTargetFolderPath & "DB"
        Fso.CreateFolder(uTargetDbFolderPath)

        If sDbFilePath_AP <> "" And _
                count_BPFile > -1 Then
            document.getElementById("db_count1").innerHTML = 2 + count_BPFile
        Else
            document.getElementById("db_count1").innerHTML = 1
        End If

        If sDbFilePath_AP <> "" Then
            Fso.copyfile uCodeOutProjectPath & sDbFilePath_AP, uTargetDbFolderPath & "\", false
            dbCount()
        End If

        If count_BPFile > -1 Then
            Dim i
            For i = 0 To count_BPFile
                Fso.copyfile uCodeOutProjectPath & aDbFilePath_BP(i), uTargetDbFolderPath & "\", false
                dbCount()
            Next
        End If

    End If
End Function

Function checkSoftwareFiles()
    If bContinue And (NOT bNotCopySoftware) then
        Dim oFolder, oFiles, File, sTmp
        Set oFolder = Fso.GetFolder(uCodeOutProjectPath)
        Set oFiles = oFolder.Files        
        
        For Each File in oFiles
            sTmp = File.Name
            If InStr(sTmp, "_Android_scatter.txt") > 0 Then
                uScatterFilePath = uCodeOutProjectPath & sTmp
                Exit For
            End If
        Next
        
        aAllNeedFilesName = getNeedFilesName(uScatterFilePath)

        Dim iSoftwareFileCount
        iSoftwareFileCount = UBound(aAllNeedFilesName) + 1

        If iSoftwareFileCount > 0 Then
            document.getElementById("software_count1").innerHTML = iSoftwareFileCount
        Else
            bContinue = False
        End If
    End If
End Function

Function copySoftware()
    If bContinue And (NOT bNotCopySoftware) then      
    '/////////////拷贝软件/////////////////////////////////////////////////////
        runCopyFile uScatterFilePath, uTargetFolderPath

        Dim uCopyFilePath, i
        For i = 0 To UBound(aAllNeedFilesName)
            uCopyFilePath = uCodeOutProjectPath & aAllNeedFilesName(i)
            If aAllNeedFilesName(i) <> "system.img" Then
                runCopyFile uCopyFilePath, uTargetFolderPath
                softwareCount()
            End If
        Next

        Sleep(1)

        uCopyFilePath = uCodeOutProjectPath & "system.img"
        runCopyFile uCopyFilePath, uTargetFolderPath
        softwareCount()
    End If
End Function

Function checkOtaPath()
    If bContinue Then
        If getCheckedRadio("otaFilePicker") <> "no_ota" Then
            Select Case getCheckedRadio("otaFilePicker")
                Case "target_files-package"
                    otaFilePath = getFilePath(uCodeOutProjectPath, "target_files-package.zip")
                Case "-ota-"
                    otaFilePath = getFilePath(uCodeOutProjectPath, "-ota-")
                Case "obj/-target_files-"
                    otaFilePath = getFilePath(uCodeOutProjectPath & "obj\PACKAGING\target_files_intermediates\", "-target_files-")
            End Select

            'MsgBox(otaFilePath)

            If otaFilePath <> "" Then
                bCopyOta = True
            Else
                bCopyOta = False
                MsgBox("OTA文件不存在")
            End If

        End If

    Else
        bCopyOta = False
    End If
End Function

Function copyOtaFile()
    If bContinue And bCopyOta Then
    '/////////////拷贝OTA文件/////////////////////////////////////////////////////
        uTargetOtaFolderPath = uTargetFolderPath & "OTA"
        otaFilePath = otaFilePath

        Fso.CreateFolder(uTargetOtaFolderPath)
        document.getElementById("ota_count1").innerHTML = 1
        Sleep(1)

        runCopyFile otaFilePath, uTargetOtaFolderPath & "\"

        otaCount()    
    End If
End Function

Function getFilePath(uFolderPath, sFileNamePart)
    If Fso.FolderExists(uFolderPath) Then
        Dim oFolder, oFile, File, sTmp
        Set oFolder = Fso.GetFolder(uFolderPath)
        Set oFile = oFolder.Files          
        For Each File in oFile
            sTmp = File.Name
            If InStr(sTmp, sFileNamePart) > 0 Then
                Exit For
            End If
        Next
    End If

    getFilePath = uFolderPath & sTmp
End Function




Function runCopy()
    initBoolean()
    cleanCount()
    getInputValue()
    checkInputInfo()
    getProjectName()
    checkDbPath()
    getDbFilePath()

    copyDbFile()
    checkSoftwareFiles()
    copySoftware()
    checkOtaPath()
    copyOtaFile()

    If bContinue Then
        MsgBox("Copy done!")
    End If

End Function