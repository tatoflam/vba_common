Attribute VB_Name = "CommonUtility"
Option Explicit

Const DATE_FORMAT_GAPPI As String = "月日"
Const SHEET_HOLIDAY As String = "祝日設定"
Const HOLIDAY_RANGE As String = "A2:A31" 'as an example


Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ファイルを開く関数
' 引数
'  1. TargetPath   開くファイルのパス
'  2. check  オープン中チェックをするフラグ
' 戻り値
'  開いたブックのオブジェクト
'  チェックでNGになった場合は戻り値(Workbook)にNothingを返す
 Public Function FileOpen(ByVal targetPath As String, ByVal check As Boolean) As Workbook
'Function FileOpen(ByVal TargetPath As String, ByVal check As Boolean)
    
    Dim buf As String, wb As Workbook
        
    ' ファイル存在チェック
    buf = dir(targetPath)
    If buf = "" Then
        MsgBox targetPath & vbCrLf & "は存在しません。処理を中止します", vbExclamation
        Exit Function
    End If
    
    ' オープン中のブックのチェック
    If check = True Then
        For Each wb In Workbooks
            If wb.name = buf Then
                MsgBox buf & vbCrLf & "は既に開いています。処理を中止します", vbExclamation
                Exit Function
            End If
        Next wb
    End If
    Workbooks.Open (targetPath)
    Set FileOpen = Workbooks(buf)
    
End Function


' ActiveなBookを引数名でファイルとして保存する関数
' 引数
'  1. TargetPath   保存するファイルパス
' 戻り値 なし
Public Function SaveFile(targetPath As String)
    Dim buf As String           'Dirチェック用バッファ
    Dim wb As Workbook          'オープン中チェック対象のWorkBook
    Dim re As Long              '存在チェック時のvbオプションの戻り
    
            
    ' 既存ファイルのチェック
    buf = dir(targetPath)
   
    ' 同名ブックが開いているかチェック（処理を中止する）
    For Each wb In Workbooks
        If wb.name = buf Then
            Application.DisplayAlerts = False
            MsgBox buf & vbCrLf & "は既に開いています。ファイル作成を中止します", _
                vbExclamation
                ActiveWorkbook.Close
            Exit Function
            Application.DisplayAlerts = True
        End If
    Next wb
        
    ' 同名ブックが存在するかチェック（上書きするか確認する）
    If buf <> "" Then
        re = MsgBox(buf & vbCrLf & "は既に存在します。置き換えますか？", _
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

' ActiveなWorkbookを引数名で保存、クローズする関数
Public Function SaveAndClose(filePath As String)

    ' 直後に保存するとエラーすることがあるため2秒待つ
    Call Sleep(500)

    On Error GoTo SaveError
'         re = MsgBox("check", vbInformation + vbYesNoCancel + vbDefalutButton2)
        ActiveWorkbook.SaveAs fileName:=filePath
        MsgBox (filePath & "を保存しました")
        ActiveWorkbook.Close
        Exit Function
SaveError:
    MsgBox ("エラーが発生したため処理を中止します")
        
End Function

' ActiveなWorkbookをアラートなしで引数名で保存、クローズする関数
Public Function SilentSaveAndClose(filePath As String)
    
    ' 直後に保存するとエラーすることがあるため待つ
    Call Sleep(500)
    
    Application.DisplayAlerts = False
    On Error GoTo SaveError
'         re = MsgBox("check", vbInformation + vbYesNoCancel + vbDefalutButton2)
        ActiveWorkbook.SaveAs fileName:=filePath
        ActiveWorkbook.Close
        Exit Function
    Application.DisplayAlerts = True

SaveError:
    MsgBox ("エラーが発生したため処理を中止します")
        
End Function

' ActiveなWorkbookをアラートなしで保存せずにクローズする関数
Public Function SilentClose()
    
    ' 直後に保存するとエラーすることがあるため待つ
    Call Sleep(500)
    
    Application.DisplayAlerts = False
    On Error GoTo SaveError
'         re = MsgBox("check", vbInformation + vbYesNoCancel + vbDefalutButton2)
        ActiveWorkbook.Close
        Exit Function
    Application.DisplayAlerts = True

SaveError:
    MsgBox ("エラーが発生したため処理を中止します")
        
End Function



' 引数に渡したディレクトリ名をチェック、絶対パスに直して返す関数
Public Function getAbsoluteDirPath(filedir As String) As String
    If dir(filedir, vbDirectory) = "" Then
        
        ' 相対パスの場合の対応
        If filedir Like "\*" Then
            filedir = ThisWorkbook.path & filedir
        Else
            filedir = ThisWorkbook.path & "\" & filedir
        End If
        ' 相対パスにもなければ終了
        If dir(filedir, vbDirectory) = "" Then
            MsgBox ("出力先フォルダがありません。フォルダを作成してください")
            Exit Function
        End If
    End If
    getAbsoluteDirPath = filedir
End Function

' 引数に渡したセルに設定された日付を判定し、土日祝日であれば背景色を変える関数
Public Function setHolidayColor(dateCell As Range)
    Dim foundCell As Range
    
    If Not IsDate(dateCell.Value) Then
        Exit Function
    End If
    
    ' 祝日シートをチェック
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
    fileName = Application.GetOpenFilename("Microsoft Excelブック, *xls ; xlsx")
    If fileName <> "False" Then
        GetFileName = fileName
    Else
        Exit Function
    End If
End Function


Public Function CnvZenKanaToHanEx(a_sZen As String) As String

    Dim sZenList                    '// 全角文字列挙
    Dim sHanList                    '// 半角文字列挙
    Dim sZenAr()                    '// 全角文字配列
    Dim sHanAr()                    '// 半角文字配列
    Dim sZen                        '// 全角文字
    Dim sHan                        '// 半角文字
    Dim i                           '// ループカウンタ
    Dim iLen                        '// 文字数
    Dim a_sHan As String
    
    
    '// 全角文字と半角文字を列挙（並び順は全角と半角で同じ文字にする）
    sZenList = "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポ。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン"
    sHanList = "ｶﾞｷﾞｸﾞｹﾞｺﾞｻﾞｼﾞｽﾞｾﾞｿﾞﾀﾞﾁﾞﾂﾞﾃﾞﾄﾞﾊﾞﾋﾞﾌﾞﾍﾞﾎﾞﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ"
    
    '// 文字数を取得
    iLen = Len(sZenList)
    
    '// 配列を初期化
    ReDim sZenAr(iLen)
    ReDim sHanAr(iLen)
    
    '// 文字数ループし、全角文字と半角文字をそれぞれ配列化
    For i = 0 To iLen - 1
        '// 半角のガからポまでは記号を含めて２文字ずつ取得
        If (i < 25) Then
            sHanAr(i) = Mid(sHanList, (i * 2) + 1, 2)
        Else
            sHanAr(i) = Mid(sHanList, i + 26, 1)
        End If
        
        '// 全角文字を取得
        sZenAr(i) = Mid(sZenList, i + 1, 1)
    Next
        
    '// 初期値として第一引数を設定
    a_sHan = a_sZen
    
    '// 半角カタカナの種類数ループ
    For i = 0 To iLen - 1
        '// カタカナの文字を全角から半角に置換
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
            MsgBox "検索範囲(" & srcSize & ")と戻り範囲(" & tgtSize & ")の大きさが異なります"
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
    MsgBox "myXlookupでエラーが発生したため処理を中止します"

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
            MsgBox "検索範囲(" & srcRng.Rows.Count & ")と戻り範囲(" & tgtRng.Rows.Count & ")の大きさが異なります"
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
    MsgBox "myXlookup2でエラーが発生したため処理を中止します"

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
            MsgBox "検索範囲(" & srcRng.Rows.Count & ")と戻り範囲(" & tgtRng.Rows.Count & ")の大きさが異なります"
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
    MsgBox "myXlookup2でエラーが発生したため処理を中止します"

End Function



'yyyy/mm/dd, yyyymmdd, mm/dd, mmdd 形式の文字列のみをDate型に直せるかチェックする関数
Function IsDateEx(s)
    Dim i
    Dim sDate
    Dim sTemp
    
    sDate = ""
    
    '// 数字のみを抽出
    For i = 0 To Len(s)
        sTemp = Mid(s, i + 1, 1)
        
        '// 数字の場合
        If IsNumeric(sTemp) = True Then
            sDate = sDate & sTemp
        End If
    Next
    
    '// 数字8文字でない場合は不正とみなす
    If Len(sDate) = 8 Then
        '// 日付形式に変換
        sDate = format(sDate, "####/##/##")
        
    ElseIf Len(sDate) = 4 Then
        sDate = format(sDate, "##/##")
    Else
        IsDateEx = False
        Exit Function
    End If
        
    '// 日付チェック
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

            '// 数字のみを抽出
            For i = 0 To Len(fileNameDateStr)
                sTemp = Mid(fileNameDateStr, i + 1, 1)
    
                '// 数字の場合
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
            convertedDate = format(fileNameDateDt, "mm月dd日")
        ElseIf date_format = DATE_FORMAT_MMDD Then
            convertedDate = format(fileNameDateDt, "mmdd")
        End If
        Debug.Print convertedDate
    Else
        ' MsgBox "日付形式 または yyyy/mm/ddのテキスト形式でファイル名に付与する日付を指定してください。処理を終了します"
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
     
    ' indexの初期値を0とする
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
        '空白以外を取り込む(列単位で処理後、行単位で処理)
        If r.Value <> "" Or includeBlank Then
            ReDim Preserve tmp(1 To i)
            tmp(i) = r.Value
            i = i + 1
        End If
    Next
    getRangeTo1DArray = tmp
End Function
