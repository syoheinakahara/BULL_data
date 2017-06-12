Attribute VB_Name = "ExportMacro"
' Option Explicit

' �O���[�o���ϐ�
Dim exportPath As String
Dim outputSheetName As String
Dim sourceSheetName As String
Dim exportFileName As String
Dim exportFileExtention As String

Dim sourceRows, sourceCols, sourceItems As Long
Dim outputRows, outputCols, outputItems As Long

Dim resultExported As Boolean
Dim resultPrompt As String

' �e��V�[�g��t�@�C����������
Function PrepareSheetAndFileInfo(sheetName As String) As Boolean
    
    ' �ۑ�����p�X������
    exportPath = ActiveWorkbook.Path
    exportPath = Replace(exportPath, "master_excel", "")
    exportPath = exportPath & "master" & Application.PathSeparator
    
    ' �ۑ�����t�@�C�����Ɗg���q������
    If InStr(sheetName, ".") Then ' �A�N�e�B�u�V�[�g���o�̓V�[�g�̏ꍇ
        ' �A�N�e�B�u�V�[�g�����o�̓V�[�g���Ƃ��ĕۑ�
        outputSheetName = sheetName
        ' �o�̓V�[�g������g���q�����o���ĕۑ�
        exportFileExtention = Mid(outputSheetName, InStr(outputSheetName, "."))
        ' ���f�[�^�V�[�g����ۑ�
        If Worksheets(outputSheetName).Cells(1, 3).Value = Replace(outputSheetName, exportFileExtention, "") Then
            sourceSheetName = Replace(outputSheetName, exportFileExtention, "")
        Else ' ���f�[�^�V�[�g�����ڎw�肳��Ă���ꍇ
            sourceSheetName = Worksheets(outputSheetName).Cells(1, 3).Value
        End If
    Else ' �A�N�e�B�u�V�[�g�����f�[�^�V�[�g�̏ꍇ
        ' �Ή�����o�̓V�[�g��T��
        Dim sheet_id
        For sheet_id = 1 To Sheets.Count
            Dim targetsheetName
            targetsheetName = Sheets(sheet_id).Name
            If InStr(targetsheetName, sheetName & ".") Then
                ' �A�N�e�B�u�V�[�g�������f�[�^�V�[�g���Ƃ��ĕۑ�
                sourceSheetName = sheetName
                ' �o�̓V�[�g����ۑ�
                outputSheetName = targetsheetName
                ' �Ή�����o�̓V�[�g�̊g���q���擾
                exportFileExtention = Mid(outputSheetName, InStr(outputSheetName, "."))
            ElseIf sheet_id = Sheets.Count Then ' ������Ȃ�������
                resultExported = False
                resultPrompt = resultPrompt & "No matching sheet found."
                PrepareSheetAndFileInfo = False
            End If
        Next sheet_id
    End If
    
    ' �o�̓V�[�g�����G�N�X�|�[�g�t�@�C�����Ƃ��ĕۑ�
    exportFileName = outputSheetName
    
    PrepareSheetAndFileInfo = True

End Function

' ���f�[�^�V�[�g�Əo�̓V�[�g�𓯊�����
Sub SyncSourceAndOutputSheets()

    ' ���f�[�^�V�[�g�Ƀt�B���^�[���������Ă��������
    With Worksheets(sourceSheetName)
        If .FilterMode = True Then
            .ShowAllData
        End If
    End With

    ' �o�̓V�[�g���Čv�Z
    RecalculateSheet (outputSheetName)
    
    sourceRows = Worksheets(sourceSheetName).UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    sourceCols = Worksheets(sourceSheetName).UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    sourceItems = sourceRows - Worksheets(outputSheetName).Cells(1, 1).Value - 2
    
    outputRows = Worksheets(outputSheetName).UsedRange.Find("*", , xlFormulas, , xlByRows, xlPrevious).Row
    outputCols = Worksheets(outputSheetName).UsedRange.Find("*", , xlFormulas, , xlByColumns, xlPrevious).Column
    outputItems = outputRows - 2
    
    ' ���f�[�^�V�[�g�Əo�̓V�[�g�̗v�f�����Ⴄ�ꍇ
    If sourceItems <> outputItems Then
        ' �o�̓V�[�g��4�s�ڈȍ~���N���A
        If outputItems <> 1 Then
'            Worksheets(outputSheetName).Range(Cells(4, 1), Cells(outputRows, outputCols)).Clear
            Worksheets(outputSheetName).Rows("4:" & outputRows).Clear
        End If
        
        ' 3�s�ڂ̓��e�����f�[�^�V�[�g�̗v�f���������R�s�[
        Worksheets(outputSheetName).Rows(3).Copy Worksheets(outputSheetName).Rows("4:" & sourceItems + 2)
        
        ' ����ŏo�̓V�[�g�̗v�f���ƌ��f�[�^�V�[�g�̗v�f���͓���
        outputItems = sourceItems
        
        ' �o�̓V�[�g���Čv�Z
        RecalculateSheet (outputSheetName)
    End If
    
End Sub

Sub RecalculateSheet(targetSheet As String)

    Worksheets(targetSheet).EnableCalculation = True
    Worksheets(targetSheet).EnableCalculation = False

End Sub


' �V�[�g�������o��
Sub ExportSheet(sheetName As String)
	' �V�����o�������p
    ' Dim outputRows, outputCols As Long
    Dim startRows, startCols As Long
    ' Dim outputSheetName As String
    Dim outputData As Variant
    ' Dim exportPath As String
    Dim cell As String

    ' �_�C�A���O���o���Ȃ��悤�ɂ���
    Application.DisplayAlerts = False
    
    ' ��ʕ`����~����
    Application.ScreenUpdating = False
    
    ' �G�N�X�|�[�g���ۏ������s��O��ɏ���
    resultExported = True
    
    ' �e��V�[�g��t�@�C����������
    If PrepareSheetAndFileInfo(sheetName) = False Then GoTo Closing
            
    ' ���f�[�^�V�[�g�Əo�̓V�[�g�𓯊�����
    SyncSourceAndOutputSheets
    
    ' �G���[�`�F�b�N
    
    ' �����ȃf�[�^�̌��o
    If IsError(Worksheets(outputSheetName).Cells(outputRows, outputCols)) Then
        resultExported = False
        resultPrompt = resultPrompt & sheetName & ": ���s@" & outputRows & "�s" & vbCr
        GoTo Closing
    End If

    ' �V�[�g�S�̂��R�s�[
    Worksheets(outputSheetName).Cells.Copy
    
    ' �V�K���[�N�u�b�N���J��
    Workbooks.Add
    
    ' �R�s�[�������e�̒l��V�K���[�N�u�b�N�Ƃ��ĕۑ�
    With ActiveWorkbook
        .ActiveSheet.Cells.PasteSpecial Paste:=xlPasteValues ' �l��\��t��

        ' �s�v�ȃf�[�^���폜
        .ActiveSheet.Range("1:1").Delete ' �Ǘ��p�̍s���폜
        If InStr(.ActiveSheet.Cells(1, 1).Value, "temp") Then
            .ActiveSheet.Columns(1).Delete ' �ꎞ�g�p�̗���폜
        End If
        If InStr(exportFileName, "tmp") Then
            .ActiveSheet.Range("1:1").Delete ' �ꎞ�g�p�̍s���폜
        End If
        
        .SaveAs Filename:=exportPath & exportFileName, FileFormat:=xlCSV ' �ۑ�
        .Close ' ���[�N�u�b�N�����

        ' �ŏI�s�ɉ��s������
        Dim fp
        fp = FreeFile
        Open exportPath & exportFileName For Append As #fp
        ' �ŏI�s�ɉ��s��ǉ�
        Write #fp, Lf
        Close #fp

    End With

Closing:
    
    ' ��ʕ`����ĊJ����
    Application.ScreenUpdating = True

    ' �_�C�A���O���o��悤�ɂ���
    Application.DisplayAlerts = True
    
    ' �G�N�X�|�[�g�ɐ��������ꍇ�̕\�����e������
    If resultExported Then
        resultPrompt = resultPrompt & sheetName & ": ����" & vbCr
    End If

End Sub

Sub ExportAllSheets()
	Dim Interval As Long
	Dim startTime,endTime As Date
    
    startTime = NOW()
    
    ' �Y���V�[�g�����ԂɃG�N�X�|�[�g
    For sheet_id = 1 To Sheets.Count
        If InStr(Sheets(sheet_id).Name, ".") > 0 Then
            ExportSheet (Sheets(sheet_id).Name)
        End If
    Next sheet_id

	endTime = NOW()
	Interval = DateDiff("s",startTime,endTime)
    ' �_�C�A���O�Ɍ��ʂ�\��
    MsgBox resultPrompt & vbCr & "�������ԁF" & Interval & "sec" , vbInformation, "�G�N�X�|�[�g����"

End Sub
