Attribute VB_Name = "CommonUtility"
Option Explicit

Const DATE_FORMAT_GAPPI As String = "����"
Const SHEET_HOLIDAY As String = "�j���ݒ�"
Const HOLIDAY_RANGE As String = "A2:A31" 'as an example


Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' �t�@�C�����J���֐�
' ����
'  1. TargetPath   �J���t�@�C���̃p�X
'  2. check  �I�[�v�����`�F�b�N������t���O
' �߂�l
'  �J�����u�b�N�̃I�u�W�F�N�g
'  �`�F�b�N��NG�ɂȂ����ꍇ�͖߂�l(Workbook)��Nothing��Ԃ�
 Public Function FileOpen(ByVal targetPath As String, ByVal check As Boolean) As Workbook
'Function FileOpen(ByVal TargetPath As String, ByVal check As Boolean)
    
    Dim buf As String, wb As Workbook
        
    ' �t�@�C�����݃`�F�b�N
    buf = dir(targetPath)
    If buf = "" Then
        MsgBox targetPath & vbCrLf & "�͑��݂��܂���B�����𒆎~���܂�", vbExclamation
        Exit Function
    End If
    
    ' �I�[�v�����̃u�b�N�̃`�F�b�N
    If check = True Then
        For Each wb In Workbooks
            If wb.name = buf Then
                MsgBox buf & vbCrLf & "�͊��ɊJ���Ă��܂��B�����𒆎~���܂�", vbExclamation
                Exit Function
            End If
        Next wb
    End If
    Workbooks.Open (targetPath)
    Set FileOpen = Workbooks(buf)
    
End Function


' Active��Book���������Ńt�@�C���Ƃ��ĕۑ�����֐�
' ����
'  1. TargetPath   �ۑ�����t�@�C���p�X
' �߂�l �Ȃ�
Public Function SaveFile(targetPath As String)
    Dim buf As String           'Dir�`�F�b�N�p�o�b�t�@
    Dim wb As Workbook          '�I�[�v�����`�F�b�N�Ώۂ�WorkBook
    Dim re As Long              '���݃`�F�b�N����vb�I�v�V�����̖߂�
    
            
    ' �����t�@�C���̃`�F�b�N
    buf = dir(targetPath)
   
    ' �����u�b�N���J���Ă��邩�`�F�b�N�i�����𒆎~����j
    For Each wb In Workbooks
        If wb.name = buf Then
            Application.DisplayAlerts = False
            MsgBox buf & vbCrLf & "�͊��ɊJ���Ă��܂��B�t�@�C���쐬�𒆎~���܂�", _
                vbExclamation
                ActiveWorkbook.Close
            Exit Function
            Application.DisplayAlerts = True
        End If
    Next wb
        
    ' �����u�b�N�����݂��邩�`�F�b�N�i�㏑�����邩�m�F����j
    If buf <> "" Then
        re = MsgBox(buf & vbCrLf & "�͊��ɑ��݂��܂��B�u�������܂����H", _
            vbInformation + vbYesNoCancel + vbDefaultButton2)

        Application.DisplayAlerts = False
        If re = vbYes Then
            SaveAndClose (targetPath)
        Else
            ActiveWorkbook.Close
        End If
        Application.DisplayAlerts = True
    
    Else
        SaveAndClose (targetPath)
    End If

End Function

' Active��Workbook���������ŕۑ��A�N���[�Y����֐�
Public Function SaveAndClose(filePath As String)

    ' ����ɕۑ�����ƃG���[���邱�Ƃ����邽��2�b�҂�
    Call Sleep(500)

    On Error GoTo SaveError
'         re = MsgBox("check", vbInformation + vbYesNoCancel + vbDefalutButton2)
        ActiveWorkbook.SaveAs fileName:=filePath
        MsgBox (filePath & "��ۑ����܂���")
        ActiveWorkbook.Close
        Exit Function
SaveError:
    MsgBox ("�G���[�������������ߏ����𒆎~���܂�")
        
End Function

' Active��Workbook���A���[�g�Ȃ��ň������ŕۑ��A�N���[�Y����֐�
Public Function SilentSaveAndClose(filePath As String)
    
    ' ����ɕۑ�����ƃG���[���邱�Ƃ����邽�ߑ҂�
    Call Sleep(500)
    
    Application.DisplayAlerts = False
    On Error GoTo SaveError
'         re = MsgBox("check", vbInformation + vbYesNoCancel + vbDefalutButton2)
        ActiveWorkbook.SaveAs fileName:=filePath
        ActiveWorkbook.Close
        Exit Function
    Application.DisplayAlerts = True

SaveError:
    MsgBox ("�G���[�������������ߏ����𒆎~���܂�")
        
End Function

' Active��Workbook���A���[�g�Ȃ��ŕۑ������ɃN���[�Y����֐�
Public Function SilentClose()
    
    ' ����ɕۑ�����ƃG���[���邱�Ƃ����邽�ߑ҂�
    Call Sleep(500)
    
    Application.DisplayAlerts = False
    On Error GoTo SaveError
'         re = MsgBox("check", vbInformation + vbYesNoCancel + vbDefalutButton2)
        ActiveWorkbook.Close
        Exit Function
    Application.DisplayAlerts = True

SaveError:
    MsgBox ("�G���[�������������ߏ����𒆎~���܂�")
        
End Function



' �����ɓn�����f�B���N�g�������`�F�b�N�A��΃p�X�ɒ����ĕԂ��֐�
Public Function getAbsoluteDirPath(filedir As String) As String
    If dir(filedir, vbDirectory) = "" Then
        
        ' ���΃p�X�̏ꍇ�̑Ή�
        If filedir Like "\*" Then
            filedir = ThisWorkbook.path & filedir
        Else
            filedir = ThisWorkbook.path & "\" & filedir
        End If
        ' ���΃p�X�ɂ��Ȃ���ΏI��
        If dir(filedir, vbDirectory) = "" Then
            MsgBox ("�o�͐�t�H���_������܂���B�t�H���_���쐬���Ă�������")
            Exit Function
        End If
    End If
    getAbsoluteDirPath = filedir
End Function

' �����ɓn�����Z���ɐݒ肳�ꂽ���t�𔻒肵�A�y���j���ł���Δw�i�F��ς���֐�
Public Function setHolidayColor(dateCell As Range)
    Dim foundCell As Range
    
    If Not IsDate(dateCell.Value) Then
        Exit Function
    End If
    
    ' �j���V�[�g���`�F�b�N
    Set foundCell = ThisWorkbook.Worksheets(SHEET_HOLIDAY).Range(HOLIDAY_RANGE).Find(dateCell.Value)
    
    If Weekday(dateCell.Value) = vbSunday Or _
       Weekday(dateCell.Value) = vbSaturday Or _
       Not foundCell Is Nothing _
    Then
       dateCell.Interior.ColorIndex = 22
    Else
       dateCell.Interior.ColorIndex = 2
    End If
End Function

Function getLastLow() As Long
    getLastLow = Rows.Count
End Function

Public Function GetFileName() As String
    Dim fileName As String
    fileName = Application.GetOpenFilename("Microsoft Excel�u�b�N, *xls ; xlsx")
    If fileName <> "False" Then
        GetFileName = fileName
    Else
        Exit Function
    End If
End Function


Public Function CnvZenKanaToHanEx(a_sZen As String) As String

    Dim sZenList                    '// �S�p������
    Dim sHanList                    '// ���p������
    Dim sZenAr()                    '// �S�p�����z��
    Dim sHanAr()                    '// ���p�����z��
    Dim sZen                        '// �S�p����
    Dim sHan                        '// ���p����
    Dim i                           '// ���[�v�J�E���^
    Dim iLen                        '// ������
    Dim a_sHan As String
    
    
    '// �S�p�����Ɣ��p������񋓁i���я��͑S�p�Ɣ��p�œ��������ɂ���j
    sZenList = "�K�M�O�Q�S�U�W�Y�[�]�_�a�d�f�h�o�r�u�x�{�p�s�v�y�|�B�u�v�A�E���@�B�D�F�H�������b�[�A�C�E�G�I�J�L�N�P�R�T�V�X�Z�\�^�`�c�e�g�i�j�k�l�m�n�q�t�w�z�}�~��������������������������"
    sHanList = "�޷޸޹޺޻޼޽޾޿������������������������������ߡ������������������������������������������������������������"
    
    '// ���������擾
    iLen = Len(sZenList)
    
    '// �z���������
    ReDim sZenAr(iLen)
    ReDim sHanAr(iLen)
    
    '// ���������[�v���A�S�p�����Ɣ��p���������ꂼ��z��
    For i = 0 To iLen - 1
        '// ���p�̃K����|�܂ł͋L�����܂߂ĂQ�������擾
        If (i < 25) Then
            sHanAr(i) = Mid(sHanList, (i * 2) + 1, 2)
        Else
            sHanAr(i) = Mid(sHanList, i + 26, 1)
        End If
        
        '// �S�p�������擾
        sZenAr(i) = Mid(sZenList, i + 1, 1)
    Next
        
    '// �����l�Ƃ��đ�������ݒ�
    a_sHan = a_sZen
    
    '// ���p�J�^�J�i�̎�ސ����[�v
    For i = 0 To iLen - 1
        '// �J�^�J�i�̕�����S�p���甼�p�ɒu��
        a_sHan = Replace(a_sHan, sZenAr(i), sHanAr(i))
    Next
    
    CnvZenKanaToHanEx = a_sHan
End Function

Public Function myXLookup(key As Variant, srcRng As Range, tgtRng As Range, defaultRet As String) As String

    Dim startTime As Single
    Dim endTime As Single
    startTime = Timer

    Dim srcArray As Variant
    Dim tgtArray As Variant
    
    Dim ret As String
    Dim srcSize As Long
    Dim tgtSize As Long
    
    Dim sr As Long
    
    srcArray = srcRng.Value
    tgtArray = tgtRng.Value
    
    srcSize = UBound(srcArray)
    tgtSize = UBound(tgtArray)
    
    On Error GoTo myXlookupError

        If Not srcSize = tgtSize Then
            MsgBox "�����͈�(" & srcSize & ")�Ɩ߂�͈�(" & tgtSize & ")�̑傫�����قȂ�܂�"
            Exit Function
        End If
    
        ret = defaultRet
        For sr = LBound(srcArray) To UBound(srcArray)
            If srcArray(sr, 1) = key Then
                ret = tgtArray(sr, 1)
                Exit For
            End If
        Next
        
        endTime = Timer
        endTime = endTime - startTime
        Debug.Print "myXlookup: " & endTime
        
        myXLookup = ret
        Exit Function

myXlookupError:
    Debug.Print sr
    Debug.Print srcSize
    Debug.Print tgtSize
    MsgBox "myXlookup�ŃG���[�������������ߏ����𒆎~���܂�"

End Function

Public Function myXLookup2(key As Variant, srcRng As Range, tgtRng As Range, defaultRet As String) As String

    Dim startTime As Single
    Dim endTime As Single
    startTime = Timer
    
    Dim ret As String
    Dim idIndex As Long
    
    ret = defaultRet

    On Error GoTo myXlookup2Error
    
        If Not srcRng.Rows.Count = tgtRng.Rows.Count Then
            MsgBox "�����͈�(" & srcRng.Rows.Count & ")�Ɩ߂�͈�(" & tgtRng.Rows.Count & ")�̑傫�����قȂ�܂�"
            Exit Function
        End If

        idIndex = -1
        On Error Resume Next
        ' spIndex = WorksheetFunction.Match(cid, cidCidArr, 0)
        idIndex = WorksheetFunction.Match(key, srcRng, 0)
        On Error GoTo -1
        
        If Not idIndex = -1 Then
            ret = tgtRng.Resize(1, 1).Offset(idIndex - 1, 0).Value
        End If
        
        endTime = Timer
        endTime = endTime - startTime
        ' Debug.Print "myXlookup2: " & endTime
        myXLookup2 = ret
        Exit Function
        
myXlookup2Error:
    Debug.Print idIndex
    Debug.Print srcRng.Rows.Count
    Debug.Print tgtRng.Rows.Count
    MsgBox "myXlookup2�ŃG���[�������������ߏ����𒆎~���܂�"

End Function


Public Function myXLookup2byDouble(key As Double, srcRng As Range, tgtRng As Range, defaultRet As String) As String

    Dim startTime As Single
    Dim endTime As Single
    startTime = Timer
    
    Dim ret As String
    Dim idIndex As Long
    
    ret = defaultRet

    On Error GoTo myXLookup2byDoubleError
    
        If Not srcRng.Rows.Count = tgtRng.Rows.Count Then
            MsgBox "�����͈�(" & srcRng.Rows.Count & ")�Ɩ߂�͈�(" & tgtRng.Rows.Count & ")�̑傫�����قȂ�܂�"
            Exit Function
        End If

        idIndex = -1
        ' Debug.Print "key: " & key & " type: " & VarType(key)
        On Error Resume Next
        ' spIndex = WorksheetFunction.Match(cid, cidCidArr, 0)
        idIndex = WorksheetFunction.Match(key, srcRng, 0)
        ' Debug.Print "idIndex" & idIndex
        On Error GoTo -1
        
        If Not idIndex = -1 Then
            ret = tgtRng.Resize(1, 1).Offset(idIndex - 1, 0).Value
        End If
        
        endTime = Timer
        endTime = endTime - startTime
        ' Debug.Print "myXLookup2byDouble: " & endTime
        myXLookup2byDouble = ret
        Exit Function
        
myXLookup2byDoubleError:
    Debug.Print idIndex
    Debug.Print srcRng.Rows.Count
    Debug.Print tgtRng.Rows.Count
    MsgBox "myXlookup2�ŃG���[�������������ߏ����𒆎~���܂�"

End Function



'yyyy/mm/dd, yyyymmdd, mm/dd, mmdd �`���̕�����݂̂�Date�^�ɒ����邩�`�F�b�N����֐�
Function IsDateEx(s)
    Dim i
    Dim sDate
    Dim sTemp
    
    sDate = ""
    
    '// �����݂̂𒊏o
    For i = 0 To Len(s)
        sTemp = Mid(s, i + 1, 1)
        
        '// �����̏ꍇ
        If IsNumeric(sTemp) = True Then
            sDate = sDate & sTemp
        End If
    Next
    
    '// ����8�����łȂ��ꍇ�͕s���Ƃ݂Ȃ�
    If Len(sDate) = 8 Then
        '// ���t�`���ɕϊ�
        sDate = format(sDate, "####/##/##")
        
    ElseIf Len(sDate) = 4 Then
        sDate = format(sDate, "##/##")
    Else
        IsDateEx = False
        Exit Function
    End If
        
    '// ���t�`�F�b�N
    IsDateEx = IsDate(sDate)
End Function


Public Function convertMD(input_date As Range, Optional date_format As String = DATE_FORMAT_GAPPI) As String

    Dim typeFileDate As Integer
    Dim fileNameDateStr As String
    Dim fileNameDateStrLen As Integer
    Dim fileNameDateDt As Date
    Dim isFileDateConvert As Boolean
    Dim convertedDate As String

    isFileDateConvert = False
        
    typeFileDate = VarType(input_date.Value)
    
    Debug.Print typeFileDate
    ' date type
    If typeFileDate = 7 Then
        fileNameDateDt = input_date.Value
        isFileDateConvert = True
        ' Debug.Print fileNameDateDt
    ' type is string
    ElseIf typeFileDate = 8 Or typeFileDate = 1 Or typeFileDate = 2 Then
        If IsDateEx(input_date.Value) Then
            ' Debug.Print .Range(input_date.Value).Value
            fileNameDateStr = input_date.Value
            
            Dim i
            Dim sDate
            Dim sTemp

            sDate = ""

            '// �����݂̂𒊏o
            For i = 0 To Len(fileNameDateStr)
                sTemp = Mid(fileNameDateStr, i + 1, 1)
    
                '// �����̏ꍇ
                If IsNumeric(sTemp) = True Then
                    sDate = sDate & sTemp
                End If
            Next
            
            
            fileNameDateStrLen = Len(sDate)
            ' Debug.Print fileNameDateStrLen
            
            If fileNameDateStrLen = 8 Then
                fileNameDateDt = CDate(format(input_date.Value, "####/##/##"))
                isFileDateConvert = True
            ElseIf fileNameDateStrLen = 4 Then
                fileNameDateDt = CDate(format(input_date.Value, "####"))
                isFileDateConvert = True
            Else
                isFileDateConvert = False
            End If
            
        Else
            isFileDateConvert = False
        End If
        ' Debug.Print .Range(input_date.Value).Value
    Else
        isFileDateConvert = False
    End If
    
    If isFileDateConvert Then
        ' Debug.Print fileNameDateDt
        If date_format = DATE_FORMAT_GAPPI Then
            convertedDate = format(fileNameDateDt, "mm��dd��")
        ElseIf date_format = DATE_FORMAT_MMDD Then
            convertedDate = format(fileNameDateDt, "mmdd")
        End If
        Debug.Print convertedDate
    Else
        ' MsgBox "���t�`�� �܂��� yyyy/mm/dd�̃e�L�X�g�`���Ńt�@�C�����ɕt�^������t���w�肵�Ă��������B�������I�����܂�"
        convertedDate = ERROR_STR
    End If
    
    convertMD = convertedDate
    

End Function

Public Function collectionToArray(ByVal targetCollection As Collection) As Variant
 
    Dim retArray() As Variant
    Dim arraySize
    Dim index
    Dim val
  
    arraySize = targetCollection.Count
    If arraySize <> 0 Then
        arraySize = arraySize - 1
        ReDim retArray(arraySize)
    End If
     
    ' index�̏����l��0�Ƃ���
    index = 0
    For Each val In targetCollection
        retArray(index) = val
        index = index + 1
    Next
  
    collectionToArray = retArray
End Function

' Quick sort for one dimensional array
' Usage: Call quickSortArray(myArray, LBound(myArray), UBound(myArray))
Public Sub quickSortArray(ByRef argAry() As Variant, _
                   ByVal lngMin As Long, _
                   ByVal lngMax As Long)
    Dim i As Long
    Dim j As Long
    Dim vBase As Variant
    Dim vSwap As Variant
    vBase = argAry(Int((lngMin + lngMax) / 2))
    i = lngMin
    j = lngMax
    Do
        Do While argAry(i) < vBase
            i = i + 1
        Loop
        Do While argAry(j) > vBase
            j = j - 1
        Loop
        If i >= j Then Exit Do
        vSwap = argAry(i)
        argAry(i) = argAry(j)
        argAry(j) = vSwap
        i = i + 1
        j = j - 1
    Loop
    If (lngMin < i - 1) Then
        Call quickSortArray(argAry, lngMin, i - 1)
    End If
    If (lngMax > j + 1) Then
        Call quickSortArray(argAry, j + 1, lngMax)
    End If
End Sub

Public Function getRangeTo1DArray(rng As Range, Optional includeBlank As Boolean = False) As Variant
    Dim tmp() As Variant
    Dim i As Long: i = 1
    Dim r As Range
     
    For Each r In rng
        '�󔒈ȊO����荞��(��P�ʂŏ�����A�s�P�ʂŏ���)
        If r.Value <> "" Or includeBlank Then
            ReDim Preserve tmp(1 To i)
            tmp(i) = r.Value
            i = i + 1
        End If
    Next
    getRangeTo1DArray = tmp
End Function
