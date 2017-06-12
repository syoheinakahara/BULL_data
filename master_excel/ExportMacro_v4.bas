'スタート行調整
Private Function setStartRow(inputSheeName As String) As Integer
	Dim x As Integer
	x = 1
    
	Select Case inputSheeName
		Case "stages"
			x = 2
		Case "mission_ACH"
			x = 2
		Case "mission_unlock_criteria"
			x = 2
		Case "weekly_missions"
			x = 2
		Case "weekly_mission_groups"
			x = 2
		Case "weekly_mission_schedules"
			x = 2
		Case "missions"
			x = 2
	End Select
    
	setStartRow = x
End Function

' 個別出力関数
Private Function outputSheet(inputSheetName As String, exportPath As String, outputFileName As String, outputItems() As String, outputItemCount As Integer) As Boolean
	On Error GoTo ErrorHandler

    Dim inputData As Variant
    Dim inputDataRows, inputDataCols As Integer
    Dim targetCols() As Integer
    ReDim targetCols(outputItemCount)
    Dim lngFileNum AS Integer
    IngFileNum = -1    
    
    ' 入力ゲット
    inputData = Worksheets(inputSheetName).UsedRange
    inputDataRows = Worksheets(inputSheetName).UsedRange.Rows.Count
    inputDataCols = Worksheets(inputSheetName).UsedRange.Columns.Count

    'start行一次設定
    outputStartRows = setStartRow(inputSheetName)

    ' 出力行の元シートでの位置探索
    For i = 0 To outputItemCount
        For j = 1 To inputDataCols
            If outputItems(i) = inputData(outputStartRows, j) Then
                targetCols(i) = j
                Exit For
            End If
        Next j
    Next i

    ' tmpの場合は出力範囲を一つ下げる
    If InStr(outputFileName, ".tmp") Then outputStartRows = outputStartRows + 1
        
    '出力先CSVファイルを開く
    lngFileNum = FreeFile()
    Open exportPath & outputFileName For Output As #lngFileNum

    For i = outputStartRows To inputDataRows
        For j = 0 To outputItemCount
            cell = Trim(inputData(i, targetCols(j)))
            '先頭が空白の時だけ逃げる
            If cell = "" And j = startCols Then Exit For

            'ダブルクオート対応
            If InStr(cell, """") Then cell = Replace(cell, """", """""")

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
	
	'正常終了系
	Close #lngFileNum
    outputSheet = True
	Exit Function

	'エラー発生系
ErrorHandler:
	outputSheet = False
	If lngFileNum <> -1 Then 	Close #lngFileNum
End Function

' エクスポート改良のテスト関数
Sub ExportAllSheets()
	On Error GoTo ErrorHandler

    Dim Interval As Long
    Dim startTime, endTime As Date
    
    startTime = Now()

    Dim inputInfo As Variant
    Dim inputInfoTotalCols, inputInfoTotalRows As Integer
    Dim inputSheetName() As String
    Dim outputFileCount As Integer
    Dim outputFileName() As String
    Dim outputItems() As String
    Dim outputItemCount As Integer
    Dim outputStartRow As Integer
    Dim exportPath_A As String
    Dim exportPath_B As String
    Dim result As String
    
    'inputInfoの読み込み
    Const inputStartRows = 2
    Const inputStartCols = 7
    
    '出力フォルダ
    Const directory_A = ""
    Const directory_B = "v1_2_0"

    ' 保存するパスを準備
    exportPath_A = ActiveWorkbook.Path
    exportPath_A = Replace(exportPath_A, "master_excel", "")
    exportPath_A = exportPath_A & "master" & Application.PathSeparator
    
    exportPath_B = ActiveWorkbook.Path
    exportPath_B = Replace(exportPath_B, "master_excel", "")
    exportPath_B = exportPath_B & "master" & Application.PathSeparator & directory_B & Application.PathSeparator
    
    ' 出力情報準備
    inputInfo = Worksheets("output").UsedRange
    outputFileCount = Worksheets("output").UsedRange.Rows.Count - 1
    inputInfoTotalCols = Worksheets("output").UsedRange.Columns.Count
    ReDim outputItems(inputInfoTotalCols)
    
    ReDim inputSheetName(outputFileCount - 1)
    ReDim outputFileName(outputFileCount - 1)
    For i = 0 To outputFileCount - 1
        inputSheetName(i) = inputInfo(i + inputStartRows, inputStartCols - 1)
        outputFileName(i) = inputInfo(i + inputStartRows, inputStartCols - 2)
    Next i
    
    ' 順番に出力
    result = "出力先：" & exportPath_A & vbCr
    For outSheetNum = inputStartRows To outputFileCount + 1

        '出力データ列数探索
        For i = inputStartCols To inputInfoTotalCols
            outputItemCount = i - inputStartCols
            
            If inputInfo(outSheetNum, i) = "" Then
                outputItemCount = outputItemCount - 1
                Exit For
            End If
            
            outputItems(i - inputStartCols) = inputInfo(outSheetNum, i)
        Next i
        
        'モードごとに出力関数に任せる
        sheetNum = outSheetNum - inputStartRows
        Select Case inputInfo(outSheetNum, 1)
            Case "old"
                tmpSuccessFlg = outputSheet(inputSheetName(sheetNum), exportPath_A, outputFileName(sheetNum), outputItems, outputItemCount)
                If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "：" & directory_A & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "：エラーが発生しました" & vbCr
	                Exit For
	            End if
            Case "new"
                tmpSuccessFlg = outputSheet(inputSheetName(sheetNum), exportPath_B, outputFileName(sheetNum), outputItems, outputItemCount)
                If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "：" & directory_B & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "：エラーが発生しました" & vbCr
	                Exit For
	            End if
            Case "skip"
                result = result & outputFileName(sheetNum) & "：skip" & vbCr
            Case Else
                tmpSuccessFlg = outputSheet(inputSheetName(sheetNum), exportPath_A, outputFileName(sheetNum), outputItems, outputItemCount)
                 If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "：" & directory_A & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "：エラーが発生しました" & vbCr
	                Exit For
	            End if
	            
                Call outputSheet(inputSheetName(sheetNum), exportPath_B, outputFileName(sheetNum), outputItems, outputItemCount)
                If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "：" & directory_B & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "：エラーが発生しました" & vbCr
	                Exit For
	            End if
        End Select

    Next outSheetNum
    
        endTime = Now()
        Interval = DateDiff("s", startTime, endTime)
    ' ダイアログに結果を表示
    MsgBox result & vbCr & "処理時間：" & Interval & "sec", vbInformation, "エクスポート終了"
	Exit Sub

ErrorHandler:
    '-- 例外処理
    MsgBox "エラーが発生したためcsv出力を停止します。" & vbCr & Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "エラー"
End Sub