Function getFileName(filePath)
	Dim oFile, sFileName
	Set oFile = oFso.GetFile(filePath)
	sFileName = oFile.Name
	Set oFile = Nothing
	getFileName = sFileName
End Function

Sub setFileName(filePath, sName)
	Dim oFile
	Set oFile = oFso.GetFile(filePath)
	oFile.Name = sName
	Set oFile = Nothing
End Sub

Sub deleteFile(filePath)
	Dim oFile
	Set oFile = oFso.GetFile(filePath)
	Call oFile.Delete
	Set oFile = Nothing
End Sub

Sub modifyScatterFile()
    If Not oFso.FileExists(uScatterFilePath) Then Exit Sub
    
    Dim oldFile, newFile, sFileName, newFilePath, sReadLine, replaceFlag
    Set oldFile = oFso.OpenTextFile(uScatterFilePath, FOR_READING)
    sFileName = getFileName(uScatterFilePath)
    newFilePath = uScatterFilePath + ".bak"

    Call initTxtFile(newFilePath)
    Set newFile = oFso.OpenTextFile(newFilePath, FOR_APPENDING)
    replaceFlag = False

    Do Until oldFile.AtEndOfStream
        sReadLine = oldFile.ReadLine

        If Not replaceFlag Then
	        If InStr(sReadLine, "preloader") > 0 And InStr(sReadLine, "file_name") > 0 Then
	        	newFile.WriteLine(sReadLine)
	        	sReadLine = oldFile.ReadLine

	        	If InStr(sReadLine, "is_download") > 0 And InStr(sReadLine, "true") > 0 Then
	        		sReadLine = Replace(sReadLine, "true", "false")
	        		replaceFlag = True
	        	End If
	        End If
		End If

		newFile.WriteLine(sReadLine)
    Loop

    oldFile.Close
    newFile.Close
    Set oldFile = Nothing
    Set newFile = Nothing

    Call deleteFile(uScatterFilePath)
    Call setFileName(newFilePath, sFileName)
End Sub