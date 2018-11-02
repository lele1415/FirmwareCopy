Dim pFile_AP, vaFilePath_BP

Const ID_DIV_DB_WAIT = "db_wait"
Const KK_AP_FOLDER = "\obj\CODEGEN\cgen"
Const L1_AP_FOLDER = "\obj\CGEN"
Const All_BP_FOLDER = "\system\etc\mddb"
Const N0_BP_FOLDER = "\system\vendor\etc\mddb"
Const O1_BP_FOLDER = "\vendor\etc\mddb"

Sub getDbFiles()
    Dim pFolder_AP, pFolder_BP
    
    pFolder_AP = getFolderOfAP()
    If pFolder_AP <> "" Then
        pFile_AP = getFileOfAP(pFolder_AP)
    Else
        pFile_AP = ""
    End If

    pFolder_BP = getFolderOfBP()
    If pFolder_BP <> "" Then
        Set vaFilePath_BP = getFileOfBP(pFolder_BP)
    Else
        Set vaFilePath_BP = New VariableArray
    End If

    If pFile_AP <> "" Then
        Call setElementInnerHTML(ID_DIV_DB_WAIT, vaFilePath_BP.Bound + 2)
    ELse
        Call setElementInnerHTML(ID_DIV_DB_WAIT, vaFilePath_BP.Bound + 1)
    End If
End Sub

        Function getFolderOfAP()
            Dim str
            Select Case True
                Case oFso.FolderExists(mOutSoftwarePath & KK_AP_FOLDER)
                    str = mOutSoftwarePath & KK_AP_FOLDER
                Case oFso.FolderExists(mOutSoftwarePath & L1_AP_FOLDER)
                    str = mOutSoftwarePath & L1_AP_FOLDER
                Case Else
                    str = ""
            End Select

            getFolderOfAP = str
        End Function

        Function getFileOfAP(folderPath)
            Dim str
            str = searchFolder(folderPath, "_ENUM", _
                    SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ONE, SEARCH_RETURN_PATH)
            If str <> "" Then
                str = Replace(str, "_ENUM", "")
            Else
                str = ""
                MsgBox("AP file is not found")
            End If

            getFileOfAP = str
        End Function

        Function getFolderOfBP()
            Dim str
            Select Case True
                Case oFso.FolderExists(mOutSoftwarePath & All_BP_FOLDER)
                    str = mOutSoftwarePath & All_BP_FOLDER
                Case oFso.FolderExists(mOutSoftwarePath & N0_BP_FOLDER)
                    str = mOutSoftwarePath & N0_BP_FOLDER
                Case oFso.FolderExists(mOutSoftwarePath & O1_BP_FOLDER)
                    str = mOutSoftwarePath & O1_BP_FOLDER
                Case Else
                    str = ""
            End Select

            getFolderOfBP = str
        End Function

        Function getFileOfBP(folderPath)
            Dim vaTmp
            Set vaTmp = searchFolder(folderPath, "BPLGU", _
                    SEARCH_FILE, SEARCH_ROOT, SEARCH_PART_NAME, SEARCH_ALL, SEARCH_RETURN_PATH)

            If vaTmp.Bound = -1 Then MsgBox("BP file is not found")

            Set getFileOfBP = vaTmp
        End Function