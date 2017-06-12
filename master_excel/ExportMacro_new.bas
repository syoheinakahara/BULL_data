Dim resultPrompt As String

'�X�^�[�g�s����
Private Function setStartRow(inputSheeName As String) As Integer

        Dim x As Integer
        x = 1
        
        Select Case inputSheeName
                Case "stages"
                                x = 2
        End Select
        
        setStartRow = x
End Function

' �ʏo�͊֐�
Private Sub outputSheet(inputSheeName As String, exportPath As String, outputFileName As String, outputItems() As String, outputItemCount As Integer)
    Dim resultExported As Boolean
    Dim test As Long
    Dim inputData As Variant
    Dim inputDataRows, inputDataCols As Integer
    Dim targetRows() As Integer
    ReDim targetRows(outputItemCount)
    
    resultExported = False
    
    ' ���̓Q�b�g
    inputData = Worksheets(inputSheeName).UsedRange
    inputDataRows = Worksheets(inputSheeName).UsedRange.Rows.Count
    inputDataCols = Worksheets(inputSheeName).UsedRange.Columns.Count
    
   'start�s
    startRows = setStartRow(inputSheeName)

    ' �o�͍s�T��
    For i = 0 To outputItemCount
        For j = 1 To inputDataCols
            If outputItems(i) = inputData(startRows, j) Then
                targetRows(i) = j
                Exit For
            End If
        Next j
    Next i

	' tmp�̏ꍇ�͏o�͔͈͂�+��
    If InStr(outputFileName, ".tmp") Then startRows = startRows + 1
    
    '�o�͐�CSV�t�@�C�����J��
    lngFileNum = FreeFile()
    Open exportPath & outputFileName For Output As #lngFileNum

    For i = startRows To inputDataRows
        For j = 0 To outputItemCount
            cell = Trim(inputData(i, targetRows(j)))
            '�擪���󔒂̎�����������
            If cell = "" And j = startCols Then Exit For

			'�_�u���N�I�[�g�Ή�
			If InStr(cell, """") Then cell = Replace(cell, """","""""")	

            ' ������,���s,,��؂�Ή�
            If InStr(cell, """") Or InStr(cell, "[") Or InStr(cell, vbCr) Or InStr(cell, ",") Then cell = """" & cell & """"

			'�����o��
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
        resultPrompt = resultPrompt & outputFileName & ": ����" & vbCr
    End If
End Sub

' �G�N�X�|�[�g���ǂ̃e�X�g�֐�
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
    
    ' �ۑ�����p�X������
    exportPath = ActiveWorkbook.Path
    exportPath = Replace(exportPath, "master_excel", "")
    exportPath = exportPath & "master" & Application.PathSeparator
    
    ' �o�͏�񏀔�
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
    
    ' ���Ԃɏo��
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
    ' �_�C�A���O�Ɍ��ʂ�\��
    MsgBox resultPrompt & vbCr & "�������ԁF" & Interval & "sec", vbInformation, "�G�N�X�|�[�g����"

End Sub