/**
 * 🚀 高度なExcel AI アシスタント
 * VBA生成・関数作成・テンプレート生成機能
 */

// 包括的なExcel関数データベース
const excelFunctions = {
  // 数学・統計関数
  MATH: {
    SUM: {
      name: 'SUM',
      syntax: 'SUM(number1, [number2], ...)',
      description: '範囲内の数値を合計します',
      examples: [
        { formula: '=SUM(A1:A10)', description: 'A1からA10までの合計' },
        { formula: '=SUM(A1:A5, C1:C5)', description: '複数の範囲の合計' },
        { formula: '=SUM(A1:A10, 100)', description: '範囲の合計に定数を追加' }
      ]
    },
    SUMIF: {
      name: 'SUMIF',
      syntax: 'SUMIF(range, criteria, [sum_range])',
      description: '条件に一致するセルの値を合計します',
      examples: [
        { formula: '=SUMIF(A1:A10, ">5", B1:B10)', description: 'A列が5より大きい場合のB列の合計' },
        { formula: '=SUMIF(A1:A10, "りんご", B1:B10)', description: 'A列が"りんご"の場合のB列の合計' },
        { formula: '=SUMIF(A1:A10, ">=100")', description: 'A列が100以上の値の合計' }
      ]
    },
    SUMIFS: {
      name: 'SUMIFS',
      syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
      description: '複数の条件に一致するセルの値を合計します',
      examples: [
        { formula: '=SUMIFS(C1:C10, A1:A10, "東京", B1:B10, ">100")', description: '東京かつ100より大きい値の合計' },
        { formula: '=SUMIFS(D1:D10, A1:A10, ">=2023/1/1", A1:A10, "<=2023/12/31")', description: '2023年のデータの合計' }
      ]
    },
    AVERAGE: {
      name: 'AVERAGE',
      syntax: 'AVERAGE(number1, [number2], ...)',
      description: '数値の平均値を計算します',
      examples: [
        { formula: '=AVERAGE(A1:A10)', description: 'A1からA10までの平均' },
        { formula: '=AVERAGE(A1:A5, C1:C5)', description: '複数範囲の平均' }
      ]
    },
    COUNT: {
      name: 'COUNT',
      syntax: 'COUNT(value1, [value2], ...)',
      description: '数値が入っているセルの個数を数えます',
      examples: [
        { formula: '=COUNT(A1:A10)', description: 'A1からA10で数値が入っているセル数' }
      ]
    },
    COUNTIF: {
      name: 'COUNTIF',
      syntax: 'COUNTIF(range, criteria)',
      description: '条件に一致するセルの個数を数えます',
      examples: [
        { formula: '=COUNTIF(A1:A10, ">5")', description: '5より大きい値の個数' },
        { formula: '=COUNTIF(A1:A10, "完了")', description: '"完了"と一致するセル数' }
      ]
    }
  },

  // 検索・参照関数
  LOOKUP: {
    VLOOKUP: {
      name: 'VLOOKUP',
      syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
      description: '縦方向の表から値を検索して返します',
      examples: [
        { formula: '=VLOOKUP(A2, D:G, 2, FALSE)', description: 'A2の値でD:G表の2列目から検索' },
        { formula: '=VLOOKUP("田中", A:C, 3, 0)', description: '"田中"を検索してC列の値を取得' }
      ]
    },
    HLOOKUP: {
      name: 'HLOOKUP',
      syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
      description: '横方向の表から値を検索して返します',
      examples: [
        { formula: '=HLOOKUP(A1, 1:5, 3, FALSE)', description: 'A1の値で1-5行目の3行目から検索' }
      ]
    },
    INDEX: {
      name: 'INDEX',
      syntax: 'INDEX(array, row_num, [column_num])',
      description: '配列から指定した位置の値を返します',
      examples: [
        { formula: '=INDEX(A1:C10, 5, 2)', description: '5行目の2列目の値を取得' },
        { formula: '=INDEX(A:A, MATCH("検索値", B:B, 0))', description: 'MATCHと組み合わせた検索' }
      ]
    },
    MATCH: {
      name: 'MATCH',
      syntax: 'MATCH(lookup_value, lookup_array, [match_type])',
      description: '配列内での値の相対的な位置を返します',
      examples: [
        { formula: '=MATCH("りんご", A1:A10, 0)', description: '"りんご"がA1:A10の何番目にあるか' },
        { formula: '=MATCH(100, B1:B10, 1)', description: '100以下の最大値の位置' }
      ]
    }
  },

  // 論理関数
  LOGICAL: {
    IF: {
      name: 'IF',
      syntax: 'IF(logical_test, value_if_true, [value_if_false])',
      description: '条件に基づいて値を返します',
      examples: [
        { formula: '=IF(A1>10, "合格", "不合格")', description: 'A1が10より大きければ合格' },
        { formula: '=IF(B2="", "未入力", B2)', description: 'B2が空なら未入力と表示' }
      ]
    },
    IFS: {
      name: 'IFS',
      syntax: 'IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)',
      description: '複数の条件をチェックして最初に真となる条件の値を返します',
      examples: [
        { formula: '=IFS(A1>=90, "A", A1>=80, "B", A1>=70, "C", TRUE, "D")', description: '成績判定' }
      ]
    },
    AND: {
      name: 'AND',
      syntax: 'AND(logical1, [logical2], ...)',
      description: '全ての条件が真の場合にTRUEを返します',
      examples: [
        { formula: '=AND(A1>0, B1<100)', description: 'A1が0より大きくB1が100未満' }
      ]
    },
    OR: {
      name: 'OR',
      syntax: 'OR(logical1, [logical2], ...)',
      description: 'いずれかの条件が真の場合にTRUEを返します',
      examples: [
        { formula: '=OR(A1="A", A1="B")', description: 'A1がAまたはB' }
      ]
    }
  },

  // テキスト関数
  TEXT: {
    CONCATENATE: {
      name: 'CONCATENATE',
      syntax: 'CONCATENATE(text1, [text2], ...)',
      description: 'テキストを連結します',
      examples: [
        { formula: '=CONCATENATE(A1, " ", B1)', description: 'A1とB1をスペースで連結' },
        { formula: '=A1&" "&B1', description: '&演算子を使った連結' }
      ]
    },
    LEFT: {
      name: 'LEFT',
      syntax: 'LEFT(text, [num_chars])',
      description: 'テキストの左から指定した文字数を取得します',
      examples: [
        { formula: '=LEFT(A1, 3)', description: 'A1の左から3文字' }
      ]
    },
    RIGHT: {
      name: 'RIGHT',
      syntax: 'RIGHT(text, [num_chars])',
      description: 'テキストの右から指定した文字数を取得します',
      examples: [
        { formula: '=RIGHT(A1, 4)', description: 'A1の右から4文字' }
      ]
    },
    MID: {
      name: 'MID',
      syntax: 'MID(text, start_num, num_chars)',
      description: 'テキストの途中から指定した文字数を取得します',
      examples: [
        { formula: '=MID(A1, 3, 5)', description: 'A1の3文字目から5文字取得' }
      ]
    }
  },

  // 日付・時刻関数
  DATE: {
    TODAY: {
      name: 'TODAY',
      syntax: 'TODAY()',
      description: '今日の日付を返します',
      examples: [
        { formula: '=TODAY()', description: '今日の日付' },
        { formula: '=TODAY()+7', description: '1週間後の日付' }
      ]
    },
    DATE: {
      name: 'DATE',
      syntax: 'DATE(year, month, day)',
      description: '指定した年月日の日付を作成します',
      examples: [
        { formula: '=DATE(2024, 12, 25)', description: '2024年12月25日' }
      ]
    },
    DATEDIF: {
      name: 'DATEDIF',
      syntax: 'DATEDIF(start_date, end_date, unit)',
      description: '2つの日付の差を計算します',
      examples: [
        { formula: '=DATEDIF(A1, B1, "Y")', description: '年数の差' },
        { formula: '=DATEDIF(A1, B1, "M")', description: '月数の差' },
        { formula: '=DATEDIF(A1, B1, "D")', description: '日数の差' }
      ]
    }
  }
};

// VBAテンプレート
const vbaTemplates = {
  // 基本操作
  basic: {
    copyData: {
      name: 'データコピー',
      description: '範囲のデータを別の場所にコピー',
      code: `Sub CopyData()
    ' ソース範囲を指定
    Range("A1:D10").Copy
    
    ' コピー先を指定
    Range("F1").PasteSpecial Paste:=xlPasteValues
    
    ' クリップボードをクリア
    Application.CutCopyMode = False
End Sub`
    },
    
    formatCells: {
      name: 'セル書式設定',
      description: 'セルの書式を一括設定',
      code: `Sub FormatCells()
    With Range("A1:D10")
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 255)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
End Sub`
    },
    
    autoFilter: {
      name: 'オートフィルター設定',
      description: 'データにオートフィルターを適用',
      code: `Sub SetAutoFilter()
    ' データ範囲を選択
    Range("A1").CurrentRegion.Select
    
    ' オートフィルターを適用
    Selection.AutoFilter
    
    ' 特定の条件でフィルター（例：売上 > 100000）
    ActiveSheet.Range("$A$1:$D$100").AutoFilter Field:=3, Criteria1:=">100000"
End Sub`
    }
  },

  // データ処理
  dataProcessing: {
    removeBlankRows: {
      name: '空白行削除',
      description: '空白行を一括削除',
      code: `Sub RemoveBlankRows()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If WorksheetFunction.CountA(Rows(i)) = 0 Then
            Rows(i).Delete
        End If
    Next i
End Sub`
    },
    
    sortData: {
      name: 'データソート',
      description: '複数条件でデータを並び替え',
      code: `Sub SortData()
    Range("A1").CurrentRegion.Sort _
        Key1:=Range("A1"), Order1:=xlAscending, _
        Key2:=Range("B1"), Order2:=xlDescending, _
        Header:=xlYes
End Sub`
    },
    
    createPivotTable: {
      name: 'ピボットテーブル作成',
      description: 'データからピボットテーブルを自動作成',
      code: `Sub CreatePivotTable()
    Dim sourceRange As Range
    Dim pivotTableSheet As Worksheet
    
    ' ソースデータ範囲
    Set sourceRange = Range("A1").CurrentRegion
    
    ' 新しいシートを作成
    Set pivotTableSheet = Sheets.Add
    pivotTableSheet.Name = "ピボットテーブル"
    
    ' ピボットテーブル作成
    ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sourceRange).CreatePivotTable _
        TableDestination:=pivotTableSheet.Range("A1")
End Sub`
    }
  },

  // ファイル操作
  fileOperations: {
    saveAsCSV: {
      name: 'CSV保存',
      description: '現在のシートをCSVで保存',
      code: `Sub SaveAsCSV()
    Dim filePath As String
    
    ' 保存先パスを指定
    filePath = ThisWorkbook.Path & "\\" & "export_" & Format(Now, "yyyymmdd") & ".csv"
    
    ' CSVで保存
    ActiveSheet.SaveAs Filename:=filePath, FileFormat:=xlCSV
    
    MsgBox "CSVファイルを保存しました: " & filePath
End Sub`
    },
    
    importCSV: {
      name: 'CSV読み込み',
      description: 'CSVファイルをインポート',
      code: `Sub ImportCSV()
    Dim filePath As String
    
    ' ファイル選択ダイアログ
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    
    If filePath <> "False" Then
        ' CSVをインポート
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=Range("A1"))
            .TextFileCommaDelimiter = True
            .Refresh BackgroundQuery:=False
        End With
    End If
End Sub`
    }
  },

  // 高度な機能
  advanced: {
    emailSender: {
      name: 'メール送信',
      description: 'Outlookでメールを自動送信',
      code: `Sub SendEmail()
    Dim outlookApp As Object
    Dim mailItem As Object
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)
    
    With mailItem
        .To = "recipient@example.com"
        .Subject = "Excel自動送信メール"
        .Body = "このメールはExcel VBAから送信されました。" & vbNewLine & _
               "添付ファイル: " & ThisWorkbook.Name
        .Attachments.Add ThisWorkbook.FullName
        .Send
    End With
    
    MsgBox "メールを送信しました"
End Sub`
    },
    
    webScraping: {
      name: 'Webスクレイピング',
      description: 'WebページからデータをExcelに取得',
      code: `Sub WebScraping()
    Dim ie As Object
    Dim doc As Object
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    With ie
        .Visible = False
        .Navigate "https://example.com"
        Do While .Busy Or .ReadyState <> 4
            DoEvents
        Loop
        
        Set doc = .Document
        
        ' データを取得してセルに入力
        Range("A1").Value = doc.getElementById("data").innerText
        
        .Quit
    End With
End Sub`
    }
  }
};

// Excel問題解決パターン
const problemSolvers = {
  '重複データを削除したい': {
    solution: 'データ > 重複の削除を使用',
    formula: '=COUNTIF($A$1:A1,A1)=1',
    vba: 'Range("A1:A100").RemoveDuplicates Columns:=1, Header:=xlYes'
  },
  
  '条件に合うデータをカウントしたい': {
    solution: 'COUNTIF関数またはCOUNTIFS関数を使用',
    formula: '=COUNTIFS(A:A,"条件1",B:B,">100")',
    vba: 'WorksheetFunction.CountIfs(Range("A:A"), "条件1", Range("B:B"), ">100")'
  },
  
  '売上の合計を計算したい': {
    solution: 'SUM関数またはSUMIF関数を使用',
    formula: '=SUMIF(A:A,"商品名",B:B)',
    vba: 'WorksheetFunction.SumIf(Range("A:A"), "商品名", Range("B:B"))'
  },
  
  '日付の差を計算したい': {
    solution: 'DATEDIF関数を使用',
    formula: '=DATEDIF(A1,B1,"D")',
    vba: 'DateDiff("d", Range("A1").Value, Range("B1").Value)'
  }
};

class ExcelAI {
  
  // 関数を検索して詳細を返す
  static searchFunction(query) {
    const results = [];
    const lowerQuery = query.toLowerCase();
    
    for (const [category, functions] of Object.entries(excelFunctions)) {
      for (const [key, func] of Object.entries(functions)) {
        if (func.name.toLowerCase().includes(lowerQuery) ||
            func.description.toLowerCase().includes(lowerQuery)) {
          results.push({
            ...func,
            category: category
          });
        }
      }
    }
    
    return results;
  }
  
  // VBAコードを生成
  static generateVBA(request) {
    const lowerRequest = request.toLowerCase();
    
    // キーワードベースでVBAテンプレートを検索
    for (const [category, templates] of Object.entries(vbaTemplates)) {
      for (const [key, template] of Object.entries(templates)) {
        if (lowerRequest.includes(key) || 
            template.name.toLowerCase().includes(lowerRequest) ||
            template.description.toLowerCase().includes(lowerRequest)) {
          return {
            category: category,
            template: template,
            code: template.code
          };
        }
      }
    }
    
    // 見つからない場合は基本テンプレートを返す
    return this.createCustomVBA(request);
  }
  
  // カスタムVBA生成
  static createCustomVBA(request) {
    const templates = {
      'ループ処理': `Sub LoopExample()
    Dim i As Long
    
    For i = 1 To 100
        ' 処理内容をここに記述
        Cells(i, 1).Value = i
    Next i
End Sub`,
      
      'セル操作': `Sub CellOperation()
    ' セルの値を設定
    Range("A1").Value = "Hello World"
    
    ' セルの書式設定
    Range("A1").Font.Bold = True
    Range("A1").Interior.Color = RGB(255, 255, 0)
End Sub`,
      
      'ファイル操作': `Sub FileOperation()
    ' ファイルを開く
    Workbooks.Open "C:\\path\\to\\file.xlsx"
    
    ' ファイルを保存
    ActiveWorkbook.Save
End Sub`
    };
    
    // 最も適切なテンプレートを選択
    for (const [key, code] of Object.entries(templates)) {
      if (request.toLowerCase().includes(key.toLowerCase())) {
        return {
          category: 'custom',
          template: {
            name: key,
            description: `${request}の処理`,
            code: code
          },
          code: code
        };
      }
    }
    
    // デフォルトテンプレート
    return {
      category: 'basic',
      template: {
        name: 'カスタムマクロ',
        description: request,
        code: `Sub CustomMacro()
    ' ${request}の処理をここに記述してください
    
    ' 例: セルに値を設定
    Range("A1").Value = "処理結果"
    
    MsgBox "処理が完了しました"
End Sub`
      },
      code: `Sub CustomMacro()
    ' ${request}の処理をここに記述してください
    
    ' 例: セルに値を設定
    Range("A1").Value = "処理結果"
    
    MsgBox "処理が完了しました"
End Sub`
    };
  }
  
  // 問題解決の提案
  static solveProblem(problem) {
    const lowerProblem = problem.toLowerCase();
    
    for (const [key, solution] of Object.entries(problemSolvers)) {
      if (lowerProblem.includes(key.toLowerCase()) ||
          key.toLowerCase().includes(lowerProblem)) {
        return solution;
      }
    }
    
    return null;
  }
  
  // 総合的な回答生成
  static generateResponse(query) {
    const response = {
      type: 'comprehensive',
      query: query,
      functions: [],
      vba: null,
      solution: null,
      suggestions: []
    };
    
    // 関数検索
    response.functions = this.searchFunction(query);
    
    // VBA生成
    if (query.toLowerCase().includes('vba') || 
        query.toLowerCase().includes('マクロ') || 
        query.toLowerCase().includes('自動')) {
      response.vba = this.generateVBA(query);
    }
    
    // 問題解決
    response.solution = this.solveProblem(query);
    
    // 提案生成
    if (response.functions.length === 0 && !response.vba && !response.solution) {
      response.suggestions = [
        '具体的な関数名を教えてください（例：VLOOKUP、SUM）',
        'やりたいことを詳しく説明してください',
        'VBAマクロが必要でしたら「VBA」と含めてください'
      ];
    }
    
    return response;
  }
}

module.exports = { ExcelAI, excelFunctions, vbaTemplates, problemSolvers };