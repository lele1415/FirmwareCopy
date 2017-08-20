Const ID_INPUT_CODE_PATH = "input_code_path"

Const ID_LIST_CODE_PATH_SELECT_VER = "list_code_path_select_ver"
Const ID_UL_CODE_PATH_SELECT_VER = "ul_code_path_select_ver"

Const ID_LIST_CODE_PATH_KK = "list_code_path_kk"
Const ID_UL_CODE_PATH_KK = "ul_code_path_kk"

Const ID_LIST_CODE_PATH_L1 = "list_code_path_l1"
Const ID_UL_CODE_PATH_L1 = "ul_code_path_l1"

Const ID_LIST_CODE_PATH_M0 = "list_code_path_m0"
Const ID_UL_CODE_PATH_M0 = "ul_code_path_m0"

Const ID_LIST_CODE_PATH_N0 = "list_code_path_n0"
Const ID_UL_CODE_PATH_N0 = "ul_code_path_n0"

Const ANDROID_VERSION_N0 = "N0"
Const ANDROID_VERSION_M0 = "M0"
Const ANDROID_VERSION_L1 = "L1"
Const ANDROID_VERSION_KK = "KK"

Dim pConfigText : pConfigText = oWs.CurrentDirectory & "\config.ini"

Dim vaCodePath_N0 : Set vaCodePath_N0 = New VariableArray
Dim vaCodePath_M0 : Set vaCodePath_M0 = New VariableArray
Dim vaCodePath_L1 : Set vaCodePath_L1 = New VariableArray
Dim vaCodePath_KK : Set vaCodePath_KK = New VariableArray

Call addVerForSelect()
Call readConfigText(pConfigText)
Call addLiOfCodePath(vaCodePath_N0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_N0, ID_UL_CODE_PATH_N0)
Call addLiOfCodePath(vaCodePath_M0, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_M0, ID_UL_CODE_PATH_M0)
Call addLiOfCodePath(vaCodePath_L1, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_L1, ID_UL_CODE_PATH_L1)
Call addLiOfCodePath(vaCodePath_KK, ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_KK, ID_UL_CODE_PATH_KK)

Sub addVerForSelect()
    Call addAfterLi("N0", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLi("M0", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLi("L1", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
    Call addAfterLi("KK", ID_INPUT_CODE_PATH, ID_LIST_CODE_PATH_SELECT_VER, ID_UL_CODE_PATH_SELECT_VER)
End Sub

Sub readConfigText(DictPath)
    If Not oFso.FileExists(DictPath) Then Exit Sub
    
    Dim oText, sReadLine, sAndroidVer
    sAndroidVer = ""
    Set oText = oFso.OpenTextFile(DictPath, FOR_READING)

    Do Until oText.AtEndOfStream
        Call handleReadLine(oText, sReadLine, sAndroidVer)
    Loop

    oText.Close
    Set oText = Nothing
End Sub

Sub handleReadLine(oText, sReadLine, sAndroidVer)
    sReadLine = oText.ReadLine

    If sAndroidVer = "" Then
        Select Case Trim(sReadLine)
            Case "N0 {"
                sAndroidVer = ANDROID_VERSION_N0
                sReadLine = oText.ReadLine
            Case "M0 {"
                sAndroidVer = ANDROID_VERSION_M0
                sReadLine = oText.ReadLine
            Case "L1 {"
                sAndroidVer = ANDROID_VERSION_L1
                sReadLine = oText.ReadLine
            Case "KK {"
                sAndroidVer = ANDROID_VERSION_KK
                sReadLine = oText.ReadLine
        End Select
    End If

    If sReadLine = "}" Then sAndroidVer = ""

    If sAndroidVer <> "" Then
        sReadLine = Trim(sReadLine)
        Select Case sAndroidVer
            Case ANDROID_VERSION_N0
                Call vaCodePath_N0.Append(sReadLine)
            Case ANDROID_VERSION_M0
                Call vaCodePath_M0.Append(sReadLine)
            Case ANDROID_VERSION_L1
                Call vaCodePath_L1.Append(sReadLine)
            Case ANDROID_VERSION_KK
                Call vaCodePath_KK.Append(sReadLine)
        End Select
    End If
End Sub

Sub addLiOfCodePath(vaObj, inputId, listId, ulId)
    If vaObj.Bound <> -1 Then
        Dim i
        For i = 0 To vaObj.Bound
            Call addAfterLi(vaObj.V(i), inputId, listId, ulId)
        Next
    End If
End Sub

Sub setListValue(inputId, listId, value)
    Call showAndHide(listId, "hide")

    If listId = ID_LIST_CODE_PATH_SELECT_VER Then
        Call showAndHide(Eval("ID_LIST_CODE_PATH_" & value), "show")
    Else
        Call setElementValue(inputId, value)
        Call Sleep(1)

        Select Case inputId
            Case ID_INPUT_CODE_PATH
                Call onChangeCodePath(CHECK_WHEN_INPUT)
            Case ID_INPUT_TARGET_PATH
                Call onChangeTargetPath(CHECK_WHEN_INPUT)
        End Select
    End If

End Sub
