Option Explicit
Sub main()
    'メイン処理を行う
    
    Dim newBook, appereBook As Workbook
    Dim newBookName, resultSheet, appereBookName

    newBookName = "AppearanceChecker_" & Format(Now, "yyyy_mmdd_hhmm_ss") & ".xlsx"
    MsgBox newBookName
    resultSheet = "Result"

    Set newBook = Workbooks.Add '新ブック作成
    newBook.SaveAs newBookName '新ブック名前変更
    
    Worksheets().Add After:=Worksheets(Worksheets.Count)  ' 新ブック末尾にシート作成
    ActiveSheet.Name = resultSheet 'シート名をResultに指定
    newBook.Worksheets(resultSheet).Cells(2, 1) = "シート名" 'A2セルに「シート名」と記載
    

    Worksheets().Add After:=newBook.Worksheets(newBook.Worksheets.Count)   ' 末尾に追加
    ActiveSheet.Name = "Result-Columns" 'シート名はResult-Columunsに指定
    newBook.Worksheets("Result-columns").Select '今いじるのは出力先なので、指定
    Range("A1").Value = "Sheets"
    Range("B1").Value = "Cells"
    Range("C1").Value = "Font_Color"
    Range("D1").Value = "Font_type"
    Range("E1").Value = "Font_Size"
    Range("F1").Value = "Background_Color"
    
    appereBookName = Application.GetOpenFilename()  '対象のファイル名を選択
    Set appereBook = Workbooks.Open(appereBookName) '対象ファイルをオブジェクトに格納
    
    'ここまでで、出力先ファイルと対象ファイルの指定、取り込みが完了

    
    'ブック情報の取得
    '(1)で選択したファイルの各ブック情報を取得し、新ブックに書き込む
    '新ブック名と、対象ファイル名を渡す
    Dim i
    For i = 1 To appereBook.Worksheets.Count
        '出力ファイル.ResultSheet.2行目←インプットファイル.全シート.シート名の転記を実施
        newBook.Worksheets(resultSheet).Cells(2, i + 1) = appereBook.Worksheets(i).Name
    Next i
    
    '各シートのセル情報を取得
    'ブック情報を渡す
    '(1)で取得したブックの全シートの全セル情報を取得して、新ブック書き込む
    
    Dim j
    Dim k
    Dim h
    
    Dim OutputColumns As Integer
    OutputColumns = 2
    
    Dim MaxRow As Integer '最終セルの行番号
    Dim MaxColumn As Integer '最終セルの列番号
 
    Dim hoge
    
   
    For j = 1 To appereBook.Sheets.Count 'シート数がん回し
    
        'シートの最終行、最終列の取得
        MaxRow = 0
        MaxColumn = 0
        MaxRow = appereBook.Sheets(Sheets(j).Name).Range("A1").SpecialCells(xlLastCell).Row
        MaxColumn = appereBook.Sheets(Sheets(j).Name).Range("A1").SpecialCells(xlLastCell).Column
        
        k = 1
        h = 1
    
        For k = 1 To MaxColumn '最終行まで分回し
            For h = 1 To MaxRow '最終列までがん回し
            
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 1) = appereBook.Sheets(Sheets(j).Name).Name
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 2) = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Address
                
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 3) = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Font.ColorIndex
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 3).Font.ColorIndex = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Font.ColorIndex

                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 4) = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Font.Name
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 4).Font.Name = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Font.Name

                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 5) = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Font.Size
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 5).Font.Size = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Font.Size

                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 6) = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Interior.ColorIndex
                Workbooks(newBookName).Worksheets("Result-columns").Cells(OutputColumns, 6).Interior.ColorIndex = appereBook.Sheets(Sheets(j).Name).Cells(h, k).Interior.ColorIndex
                OutputColumns = OutputColumns + 1
                
            Next h
        Next k
    Next j
    
     Workbooks(newBookName).Close
     appereBook.Close

End Sub






Sub getcellsettings()
    ' 出力ブックに新しいシートを作成して、各シートの各セルのフォント情報を格納するモジュール
    
    Dim j
    Dim k
    Dim h
    
    Dim OutputColumns As Integer
    OutputColumns = 2
    
    Dim MaxRow As Long '最終セルの行番号
    Dim MaxColumn As Integer '最終セルの列番号
 
    
    For j = 1 To Sheets.Count 'シート数がん回し
    
        'シートの最終行、最終列の取得
        MaxRow = 0
        MaxColumn = 0
        MaxRow = Worksheets(Sheets(j).Name).Range("A1").SpecialCells(xlLastCell).Row
        MaxColumn = Worksheets(Sheets(j).Name).Range("A1").SpecialCells(xlLastCell).Column
        
        k = 1
        h = 1
    
        For k = 1 To MaxColumn '最終行まで分回し
            For h = 1 To MaxRow '最終列までがん回し
            
                Worksheets("Result-Columns").Cells(OutputColumns, 1) = Sheets(Sheets(j).Name).Name
                Worksheets("Result-Columns").Cells(OutputColumns, 2) = Sheets(Sheets(j).Name).Cells(h, k).Address
                
                Worksheets("Result-Columns").Cells(OutputColumns, 3) = Sheets(Sheets(j).Name).Cells(h, k).Font.ColorIndex
                Worksheets("Result-Columns").Cells(OutputColumns, 3).Font.ColorIndex = Sheets(Sheets(j).Name).Cells(h, k).Font.ColorIndex

                Worksheets("Result-Columns").Cells(OutputColumns, 4) = Sheets(Sheets(j).Name).Cells(h, k).Font.Name
                Worksheets("Result-Columns").Cells(OutputColumns, 4).Font.Name = Sheets(Sheets(j).Name).Cells(h, k).Font.Name


                Worksheets("Result-Columns").Cells(OutputColumns, 5) = Sheets(Sheets(j).Name).Cells(h, k).Font.Size
                Worksheets("Result-Columns").Cells(OutputColumns, 5).Font.Size = Sheets(Sheets(j).Name).Cells(h, k).Font.Size

                Worksheets("Result-Columns").Cells(OutputColumns, 6) = Sheets(Sheets(j).Name).Cells(h, k).Interior.ColorIndex
                Worksheets("Result-Columns").Cells(OutputColumns, 6).Interior.ColorIndex = Sheets(Sheets(j).Name).Cells(h, k).Interior.ColorIndex
                OutputColumns = OutputColumns + 1
                
            Next h
        Next k
    Next j
        
End Sub
