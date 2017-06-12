Attribute VB_Name = "ExportMacro"
' Option Explicit

' グローバル変数
Dim exportPath As String
Dim outputSheetName As String
Dim sourceSheetName As String
Dim exportFileName As String
Dim exportFileExtention As String

Dim sourceRows, sourceCols, sourceItems As Long
Dim outputRows, outputCols, outputItems As Long

Dim resultExported As Boolean
Dim resultPrompt As String

' 各種シートやファイル情報を準備
Function PrepareSheetAndFileInfo(sheetName As String) As Boolean
    
    ' 保存するパスを準備
    exportPath = ActiveWorkbook.Path
    exportPath = Replace(exportPath, "master_excel", "")
    exportPath = exportPath & "master" & Application.PathSeparator
    
    ' 保存するファイル名と拡張子を準備
    If InStr(sheetName, ".") Then ' アクティブシートが出力シートの場合
        ' アクティブシート名を出力シート名として保存
        outputSheetName = sheetName
        ' 出力シート名から拡張子を取り出して保存
        exportFileExtention = Mid(outputSheetName, InStr(outputSheetName, "."))
        ' 元データシート名を保存
        If Worksheets(outputSheetName).Cells(1, 3).Value = Replace(outputSheetName, exportFileExtention, "") Then
            sourceSheetName = Replace(outputSheetName, exportFileExtention, "")
        Else ' 元データシートが直接指定されている場合
            sourceSheetName = Worksheets(outputSheetName).Cells(1, 3).Value
        End If
    Else ' アクティブシートが元データシートの場合
        ' 対応する出力シートを探す
        Dim sheet_id
        For sheet_id = 1 To Sheets.Count
            Dim targetsheetName
            targetsheetName = Sheets(sheet_id).Name
            If InStr(targetsheetName, sheetName & ".") Then
                ' アクティブシート名を元データシート名として保存
                sourceSheetName = sheetName
                ' 出力シート名を保存
                outputSheetName = targetsheetName
                ' 対応する出力シートの拡張子を取得
                exportFileExtention = Mid(outputSheetName, InStr(outputSheetName, "."))
            ElseIf sheet_id = Sheets.Count Then ' 見つからなかったら
                resultExported = False
                resultPrompt = resultPrompt & "No matching sheet found."
                PrepareSheetAndFileInfo = False
            End If
        Next sheet_id
    End If
    
    ' 出力シート名をエクスポートファイル名として保存
    exportFileName = outputSheetName
    
    PrepareSheetAndFileInfo = True

End Function

' 元データシートと出力シートを同期する
Sub SyncSourceAndOutputSheets()

    ' 元データシートにフィルターがかかっていたら解除
    With Worksheets(sourceSheetName)
        If .FilterMode = True Then
            .ShowAllData
        End If
    End With

    ' 出力シートを再計算
    RecalculateSheet (outputSheetName)
    
    sourceRows = Worksheets(sourceSheetName).UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    sourceCols = Worksheets(sourceSheetName).UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    sourceItems = sourceRows - Worksheets(outputSheetName).Cells(1, 1).Value - 2
    
    outputRows = Worksheets(outputSheetName).UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    outputCols = Worksheets(outputSheetName).UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    outputItems = outputRows - 2
    
    ' 元データシートと出力シートの要素数が違う場合
    If sourceItems <> outputItems Then
        ' 出力シートの4行目以降をクリア
        If outputItems <> 1 Then
'            Worksheets(outputSheetName).Range(Cells(4, 1), Cells(outputRows, outputCols)).Clear
            Worksheets(outputSheetName).Rows("4:" & outputRows).Clear
        End If
        
        ' 3行目の内容を元データシートの要素数分だけコピー
        Worksheets(outputSheetName).Rows(3).Copy Worksheets(outputSheetName).Rows("4:" & sourceItems + 2)
        
        ' これで出力シートの要素数と元データシートの要素数は同じ
        outputItems = sourceItems
        
        ' 出力シートを再計算
        RecalculateSheet (outputSheetName)
    End If
    
End Sub

Sub RecalculateSheet(targetSheet As String)

    Worksheets(targetSheet).EnableCalculation = True
    Worksheets(targetSheet).EnableCalculation = False

End Sub


' シートを書き出す
Sub ExportSheet(sheetName As String)
	' 新書き出し処理用
    ' Dim outputRows, outputCols As Long
    Dim startRows, startCols As Long
    ' Dim outputSheetName As String
    Dim outputData As Variant
    ' Dim exportPath As String
    Dim cell As String

    ' ダイアログを出さないようにする
    Application.DisplayAlerts = False
    
    ' 画面描画を停止する
    Application.ScreenUpdating = False
    
    ' エクスポート成否情報を失敗を前提に準備
    resultExported = True
    
    ' 各種シートやファイル情報を準備
    If PrepareSheetAndFileInfo(sheetName) = False Then GoTo Closing
            
    ' 元データシートと出力シートを同期する
    SyncSourceAndOutputSheets
    
    ' エラーチェック
    
    ' 無効なデータの検出
    If IsError(Worksheets(outputSheetName).Cells(outputRows, outputCols)) Then
        resultExported = False
        resultPrompt = resultPrompt & sheetName & ": 失敗@" & outputRows & "行" & vbCr
        GoTo Closing
    End If

    ' シート全体をコピー
    Worksheets(outputSheetName).Cells.Copy
    
    ' 新規ワークブックを開く
    Workbooks.Add
    
    ' コピーした内容の値を新規ワークブックとして保存
    With ActiveWorkbook
        .ActiveSheet.Cells.PasteSpecial Paste:=xlPasteValues ' 値を貼り付け

        ' 不要なデータを削除
        .ActiveSheet.Range("1:1").Delete ' 管理用の行を削除
        If InStr(.ActiveSheet.Cells(1, 1).Value, "temp") Then
            .ActiveSheet.Columns(1).Delete ' 一時使用の列を削除
        End If
        If InStr(exportFileName, "tmp") Then
            .ActiveSheet.Range("1:1").Delete ' 一時使用の行を削除
        End If
        
        .SaveAs Filename:=exportPath & exportFileName, FileFormat:=xlCSV ' 保存
        .Close ' ワークブックを閉じる

        ' 最終行に改行を入れる
        Dim fp
        fp = FreeFile
        Open exportPath & exportFileName For Append As #fp
        ' 最終行に改行を追加
        Write #fp, Lf
        Close #fp

    End With

Closing:
    
    ' 画面描画を再開する
    Application.ScreenUpdating = True

    ' ダイアログを出るようにする
    Application.DisplayAlerts = True
    
    ' エクスポートに成功した場合の表示内容を準備
    If resultExported Then
        resultPrompt = resultPrompt & sheetName & ": 成功" & vbCr
    End If

End Sub

Sub ExportAllSheets()
	Dim Interval As Long
	Dim startTime,endTime As Date
    
    startTime = NOW()
    
    ' 該当シートを順番にエクスポート
    For sheet_id = 1 To Sheets.Count
        If InStr(Sheets(sheet_id).Name, ".") > 0 Then
            ExportSheet (Sheets(sheet_id).Name)
        End If
    Next sheet_id

	endTime = NOW()
	Interval = DateDiff("s",startTime,endTime)
    ' ダイアログに結果を表示
    MsgBox resultPrompt & vbCr & "処理時間：" & Interval & "sec" , vbInformation, "エクスポート完了"

End Sub
