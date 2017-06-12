'�X�^�[�g�s����
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

' �ʏo�͊֐�
Private Function outputSheet(inputSheetName As String, exportPath As String, outputFileName As String, outputItems() As String, outputItemCount As Integer) As Boolean
	On Error GoTo ErrorHandler

    Dim inputData As Variant
    Dim inputDataRows, inputDataCols As Integer
    Dim targetCols() As Integer
    ReDim targetCols(outputItemCount)
    Dim lngFileNum AS Integer
    IngFileNum = -1    
    
    ' ���̓Q�b�g
    inputData = Worksheets(inputSheetName).UsedRange
    inputDataRows = Worksheets(inputSheetName).UsedRange.Rows.Count
    inputDataCols = Worksheets(inputSheetName).UsedRange.Columns.Count

    'start�s�ꎟ�ݒ�
    outputStartRows = setStartRow(inputSheetName)

    ' �o�͍s�̌��V�[�g�ł̈ʒu�T��
    For i = 0 To outputItemCount
        For j = 1 To inputDataCols
            If outputItems(i) = inputData(outputStartRows, j) Then
                targetCols(i) = j
                Exit For
            End If
        Next j
    Next i

    ' tmp�̏ꍇ�͏o�͔͈͂��������
    If InStr(outputFileName, ".tmp") Then outputStartRows = outputStartRows + 1
        
    '�o�͐�CSV�t�@�C�����J��
    lngFileNum = FreeFile()
    Open exportPath & outputFileName For Output As #lngFileNum

    For i = outputStartRows To inputDataRows
        For j = 0 To outputItemCount
            cell = Trim(inputData(i, targetCols(j)))
            '�擪���󔒂̎�����������
            If cell = "" And j = startCols Then Exit For

            '�_�u���N�I�[�g�Ή�
            If InStr(cell, """") Then cell = Replace(cell, """", """""")

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
	
	'����I���n
	Close #lngFileNum
    outputSheet = True
	Exit Function

	'�G���[�����n
ErrorHandler:
	outputSheet = False
	If lngFileNum <> -1 Then 	Close #lngFileNum
End Function

' �G�N�X�|�[�g���ǂ̃e�X�g�֐�
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
    
    'inputInfo�̓ǂݍ���
    Const inputStartRows = 2
    Const inputStartCols = 7
    
    '�o�̓t�H���_
    Const directory_A = ""
    Const directory_B = "v1_2_0"

    ' �ۑ�����p�X������
    exportPath_A = ActiveWorkbook.Path
    exportPath_A = Replace(exportPath_A, "master_excel", "")
    exportPath_A = exportPath_A & "master" & Application.PathSeparator
    
    exportPath_B = ActiveWorkbook.Path
    exportPath_B = Replace(exportPath_B, "master_excel", "")
    exportPath_B = exportPath_B & "master" & Application.PathSeparator & directory_B & Application.PathSeparator
    
    ' �o�͏�񏀔�
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
    
    ' ���Ԃɏo��
    result = "�o�͐�F" & exportPath_A & vbCr
    For outSheetNum = inputStartRows To outputFileCount + 1

        '�o�̓f�[�^�񐔒T��
        For i = inputStartCols To inputInfoTotalCols
            outputItemCount = i - inputStartCols
            
            If inputInfo(outSheetNum, i) = "" Then
                outputItemCount = outputItemCount - 1
                Exit For
            End If
            
            outputItems(i - inputStartCols) = inputInfo(outSheetNum, i)
        Next i
        
        '���[�h���Ƃɏo�͊֐��ɔC����
        sheetNum = outSheetNum - inputStartRows
        Select Case inputInfo(outSheetNum, 1)
            Case "old"
                tmpSuccessFlg = outputSheet(inputSheetName(sheetNum), exportPath_A, outputFileName(sheetNum), outputItems, outputItemCount)
                If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "�F" & directory_A & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "�F�G���[���������܂���" & vbCr
	                Exit For
	            End if
            Case "new"
                tmpSuccessFlg = outputSheet(inputSheetName(sheetNum), exportPath_B, outputFileName(sheetNum), outputItems, outputItemCount)
                If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "�F" & directory_B & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "�F�G���[���������܂���" & vbCr
	                Exit For
	            End if
            Case "skip"
                result = result & outputFileName(sheetNum) & "�Fskip" & vbCr
            Case Else
                tmpSuccessFlg = outputSheet(inputSheetName(sheetNum), exportPath_A, outputFileName(sheetNum), outputItems, outputItemCount)
                 If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "�F" & directory_A & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "�F�G���[���������܂���" & vbCr
	                Exit For
	            End if
	            
                Call outputSheet(inputSheetName(sheetNum), exportPath_B, outputFileName(sheetNum), outputItems, outputItemCount)
                If tmpSuccessFlg Then
	                result = result & outputFileName(sheetNum) & "�F" & directory_B & vbCr
	            Else
	                result = result & outputFileName(sheetNum) & "�F�G���[���������܂���" & vbCr
	                Exit For
	            End if
        End Select

    Next outSheetNum
    
        endTime = Now()
        Interval = DateDiff("s", startTime, endTime)
    ' �_�C�A���O�Ɍ��ʂ�\��
    MsgBox result & vbCr & "�������ԁF" & Interval & "sec", vbInformation, "�G�N�X�|�[�g�I��"
	Exit Sub

ErrorHandler:
    '-- ��O����
    MsgBox "�G���[��������������csv�o�͂��~���܂��B" & vbCr & Err.Number & ":" & Err.Description, vbCritical & vbOKOnly, "�G���["
End Sub