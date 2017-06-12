Dim resultPrompt As String

'スタート行調整
Private Function setStartRow(inputSheeName As String) As Integer

        Dim x As Integer
        x = 1
        
        Select Case inputSheeName
                Case "stages"
                                x = 2
        End Select
        
        setStartRow = x
End Function

' 個別出力関数
Private Sub outputSheet(inputSheeName As String, exportPath As String, outputFileName As String, outputItems() As String, outputItemCount As Integer)
    Dim resultExported As Boolean
    Dim test As Long
    Dim inputData As Variant
    Dim inputDataRows, inputDataCols As Integer
    Dim targetRows() As Integer
    ReDim targetRows(outputItemCount)
    
    resultExported = False
    
    ' 入力ゲット
    inputData = Worksheets(inputSheeName).UsedRange
    inputDataRows = Worksheets(inputSheeName).UsedRange.Rows.Count
    inputDataCols = Worksheets(inputSheeName).UsedRange.Columns.Count
    
   'start行
    startRows = setStartRow(inputSheeName)

    ' 出力行探索
    For i = 0 To outputItemCount
        For j = 1 To inputDataCols
            If outputItems(i) = inputData(startRows, j) Then
                targetRows(i) = j
                Exit For
            End If
        Next j
    Next i

	' tmpの場合は出力範囲を+α
    If InStr(outputFileName, ".tmp") Then startRows = startRows + 1
    
    '出力先CSVファイルを開く
    lngFileNum = FreeFile()
    Open exportPath & outputFileName For Output As #lngFileNum

    For i = startRows To inputDataRows
        For j = 0 To outputItemCount
            cell = Trim(inputData(i, targetRows(j)))
            '先頭が空白の時だけ逃げる
            If cell = "" And j = startCols Then Exit For

			'ダブルクオート対応
			If InStr(cell, """") Then cell = Replace(cell, """","""""")	

            ' 文字列,改行,,区切り対応
            If InStr(cell, """") Or InStr(cell, "[") Or InStr(cell, vbCr) Or InStr(cell, ",") Then cell = """" & cell & """"

			'書き出し
            Print #lngFileNum, cell;

            If j = outputItemCount Then
                Print #1, ""
                        Else
                Print #1, ",";
            End If
        Next j
    Next i
    Close #lngFileNum

    resultExported = True
    If resultExported Then
        resultPrompt = resultPrompt & outputFileName & ": 成功" & vbCr
    End If
End Sub

' エクスポート改良のテスト関数
Sub ExportAllSheets()
    Dim Interval As Long
    Dim startTime, endTime As Date
    
    startTime = NOW()

    Dim inputInfo As Variant
    Dim inputInfoTotalCols, inputInfoTotalRows As Integer
    Dim inputSheetName() As String
    Dim outputFileCount As Integer
    Dim outputFileName() As String
    Dim outputItems() As String
    Dim outputItemCount As Integer
    Dim exportPath As String
    
    ' 保存するパスを準備
    exportPath = ActiveWorkbook.Path
    exportPath = Replace(exportPath, "master_excel", "")
    exportPath = exportPath & "master" & Application.PathSeparator
    
    ' 出力情報準備
    inputInfo = Worksheets("output").UsedRange
    outputFileCount = Worksheets("output").UsedRange.Rows.Count - 1
    inputInfoTotalCols = Worksheets("output").UsedRange.Columns.Count
    ReDim outputItems(inputInfoTotalCols)
    
    ReDim inputSheetName(outputFileCount - 1)
    ReDim outputFileName(outputFileCount - 1)
    For i = 0 To outputFileCount - 1
        inputSheetName(i) = inputInfo(i + 2, 5)
        outputFileName(i) = inputInfo(i + 2, 4)
    Next i
    
    ' 順番に出力
    For outSheetNum = 2 To outputFileCount + 1
        For i = 6 To inputInfoTotalCols
            outputItemCount = i - 6
            
            If inputInfo(outSheetNum, i) = "" Then
                outputItemCount = outputItemCount - 1
                Exit For
            End If
            
            outputItems(i - 6) = inputInfo(outSheetNum, i)
        Next i
            
        Call outputSheet(inputSheetName(outSheetNum - 2), exportPath, outputFileName(outSheetNum - 2), outputItems, outputItemCount)
    Next outSheetNum
    
        endTime = NOW()
        Interval = DateDiff("s", startTime, endTime)
    ' ダイアログに結果を表示
    MsgBox resultPrompt & vbCr & "処理時間：" & Interval & "sec", vbInformation, "エクスポート完了"

End Sub