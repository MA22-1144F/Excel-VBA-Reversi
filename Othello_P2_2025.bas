Attribute VB_Name = "Othello_P2_2025"
' ==========================================================
' VBA オセロゲーム - 2人対戦版
' ==========================================================

Option Explicit

' ゲーム設定
Public Const BOARD_SIZE As Integer = 8    ' ボードサイズ（偶数推奨）
Public Const CELL_EMPTY As Integer = 0    ' 空のセル
Public Const BLACK As Integer = 1         ' プレイヤー1（黒）
Public Const WHITE As Integer = 2         ' プレイヤー2（白）

' プレイヤー情報
Public blackPlayerName As String          ' 黒プレイヤー名
Public whitePlayerName As String          ' 白プレイヤー名

' 持ち時間関連
Public blackTotalTime As Double           ' 黒の総考慮時間（秒）
Public whiteTotalTime As Double           ' 白の総考慮時間（秒）
Public currentTurnStartTime As Double     ' 現在のターン開始時刻

' ヒント機能
Public showHints As Boolean

' 棋譜表示設定
Public RECORD_START_COL As Integer        ' 棋譜表示開始列
Public Const RECORD_WIDTH As Integer = 6  ' 棋譜表示幅（列数）

' ゲーム状態保存用
Type GameState
    board(1 To 20, 1 To 20) As Integer
    Player As Integer
    moveRow As Integer
    moveCol As Integer
    movePlayer As Integer
    timestamp As String
    thinkingTime As Double
End Type

' 有効手格納用
Type ValidMove
    row As Integer
    col As Integer
End Type

' ゲーム変数
Public gameBoard(1 To 20, 1 To 20) As Integer
Public CurrentPlayer As Integer
Public gameOver As Boolean
Public targetSheet As Worksheet
Public targetWorkbook As Workbook
Public gameHistory() As GameState
Public historyCount As Integer
Public moveHistory() As String
Public moveCount As Integer

' リボンから呼び出すメイン関数
Sub StartReversiGame2P()
    On Error GoTo ErrorHandler
    Call CreateNewGameWorkbook
    If targetWorkbook Is Nothing Or targetSheet Is Nothing Then
        MsgBox "ワークブックまたはワークシートの作成に失敗しました。再度実行してください。", vbCritical
        Exit Sub
    End If
    
    Call InitializeGame2P
    Exit Sub
    
ErrorHandler:
    MsgBox "ゲーム開始時にエラーが発生しました: " & Err.description, vbCritical
End Sub

' 新しいワークブックを作成
Sub CreateNewGameWorkbook()
    Dim newWb As Workbook
    Dim ws As Worksheet
    Set newWb = Workbooks.Add
    Set ws = newWb.ActiveSheet
    If newWb Is Nothing Or ws Is Nothing Then
        MsgBox "エラー: ワークブックまたはワークシートの作成に失敗しました。", vbCritical
        Exit Sub
    End If
    ws.Name = "ゲーム盤"
    Set targetWorkbook = newWb
    Set targetSheet = ws
    Call SetupWorksheetEvents
End Sub

' ワークシートイベントの動的設定
Sub SetupWorksheetEvents()
    On Error GoTo UseSimpleMethod
    If Not CheckVBAProjectAccess() Then
        GoTo UseSimpleMethod
    End If
    Dim vbComp As Object
    Dim CodeModule As Object
    Dim eventCode As String
    
    Set vbComp = targetWorkbook.VBProject.VBComponents(targetSheet.CodeName)
    Set CodeModule = vbComp.CodeModule

    If CodeModule.CountOfLines > 0 Then
        CodeModule.DeleteLines 1, CodeModule.CountOfLines
    End If
    
    eventCode = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
               "    On Error Resume Next" & vbCrLf & _
               "    If Target.Cells.Count = 1 And Not Application.ScreenUpdating = False Then" & vbCrLf & _
               "        Application.Run """ & ThisWorkbook.Name & "!ProcessCellClick2P"", Target" & vbCrLf & _
               "    End If" & vbCrLf & _
               "End Sub"
    
    CodeModule.AddFromString eventCode

    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange2P"
    
    Exit Sub
    
UseSimpleMethod:
    Call SetupSimpleEvents
End Sub

' VBAプロジェクトアクセス可能性チェック
Function CheckVBAProjectAccess() As Boolean
    On Error GoTo AccessDenied
    Dim testName As String
    testName = targetWorkbook.VBProject.Name
    
    CheckVBAProjectAccess = True
    Exit Function
    
AccessDenied:
    CheckVBAProjectAccess = False
End Function

' 簡易版
Sub SetupSimpleEvents()
    On Error Resume Next
    Application.OnSheetSelectionChange = ""
    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange2P"
    
    On Error GoTo 0
End Sub

' アプリケーションレベルのSelectionChangeイベント
Public Sub OnSheetSelectionChange2P(Sh As Object, Target As Range)
    On Error Resume Next
    If Sh Is Nothing Or Target Is Nothing Then Exit Sub
    If Target.Cells.count <> 1 Then Exit Sub
    If targetSheet Is Nothing Then Exit Sub
    If gameOver Then Exit Sub
    If Sh.Name <> targetSheet.Name Then Exit Sub
    If Sh.Parent.Name <> targetSheet.Parent.Name Then Exit Sub
    
    Call ProcessCellClick2P(Target)
End Sub

' セルクリック処理の統合関数
Public Sub ProcessCellClick2P(Target As Range)
    Dim row As Integer, col As Integer
    
    On Error Resume Next
    If Not Target.Worksheet Is targetSheet Then Exit Sub
    
    row = Target.row
    col = Target.Column
    If row < 1 Or row > BOARD_SIZE Or col < 1 Or col > BOARD_SIZE Then Exit Sub
    If IsValidMove(row, col, CurrentPlayer) Then
        Call ExecutePlayerMove(row, col)
    Else
        Call ShowMessage("そこには置けません。別の場所を選んでください。")
    End If
End Sub

' プレイヤーの手を実行
Sub ExecutePlayerMove(row As Integer, col As Integer)
    Dim thinkingTime As Double
    thinkingTime = Timer - currentTurnStartTime
    If CurrentPlayer = BLACK Then
        blackTotalTime = blackTotalTime + thinkingTime
    Else
        whiteTotalTime = whiteTotalTime + thinkingTime
    End If
    Call SaveGameState(row, col, CurrentPlayer, thinkingTime)
    Call MakeMove(row, col, CurrentPlayer)
    Call UpdateDisplay
    Call RecordMove(row, col, CurrentPlayer, thinkingTime)
    Call UpdateGameRecordDisplay
    
    If CheckGameEnd() Then
        Call ShowResult
        Exit Sub
    End If

    Call SwitchToNextPlayer
End Sub

' 次のプレイヤーに交代
Sub SwitchToNextPlayer()
    Dim nextPlayer As Integer
    nextPlayer = 3 - CurrentPlayer
    If HasValidMoves(nextPlayer) Then
        CurrentPlayer = nextPlayer
        currentTurnStartTime = Timer
        Call UpdateTurnDisplay
        If showHints Then
            Call UpdateDisplay
        End If
    Else
        Dim skipMessage As String
        If nextPlayer = BLACK Then
            skipMessage = blackPlayerName & "（黒）に有効な手がありません。" & vbCrLf & _
            GetCurrentPlayerName() & "が続行します。"
        Else
            skipMessage = whitePlayerName & "（白）に有効な手がありません。" & vbCrLf & _
            GetCurrentPlayerName() & "が続行します。"
        End If
        
        Call ShowMessage(skipMessage)
        Call RecordMove(0, 0, nextPlayer, 0)
        Call UpdateGameRecordDisplay
        MsgBox skipMessage, vbInformation, "ターンスキップ"
        If HasValidMoves(CurrentPlayer) Then
            currentTurnStartTime = Timer
            Call UpdateTurnDisplay
            If showHints Then
                Call UpdateDisplay
            End If
        Else
            Call ShowResult
        End If
    End If
End Sub

' 現在のプレイヤー名を取得
Function GetCurrentPlayerName() As String
    If CurrentPlayer = BLACK Then
        GetCurrentPlayerName = blackPlayerName & "（黒）"
    Else
        GetCurrentPlayerName = whitePlayerName & "（白）"
    End If
End Function

' ターン表示を更新
Sub UpdateTurnDisplay()
    Dim message As String
    message = GetCurrentPlayerName() & "のターンです。"
    Call ShowMessage(message)
End Sub

' ゲーム初期化
Sub InitializeGame2P()
    On Error GoTo InitError
    
    Call SetupPlayerNames
    Call SetupGameFeatures
    Call InitializeGameBoard
    Call InitializeHistory
    Call SetupUI2P
    Call InitializeGameRecordDisplay
    blackTotalTime = 0
    whiteTotalTime = 0
    
    Call UpdateDisplay
    Call UpdateGameRecordDisplay
    Call UpdateTurnDisplay

    If Not HasValidMoves(BLACK) Then
        Call SwitchToNextPlayer
    End If

    Call SetupWorksheetEvents
    
    MsgBox "オセロ2人対戦ゲームを開始しました。" & vbCrLf & _
           "ボードサイズ: " & BOARD_SIZE & "x" & BOARD_SIZE & vbCrLf & _
           "黒: " & blackPlayerName & vbCrLf & _
           "白: " & whitePlayerName & vbCrLf & vbCrLf & _
           "セルクリックまたは手動入力でプレイしてください。", vbInformation

    currentTurnStartTime = Timer
    
    Exit Sub
    
InitError:
    MsgBox "ゲーム初期化中にエラーが発生しました: " & Err.description & vbCrLf & vbCrLf & _
           "「手動入力」ボタンでプレイを続行できます。", vbExclamation
    Resume Next
End Sub

' プレイヤー名の設定
Sub SetupPlayerNames()
    blackPlayerName = InputBox("黒（先手）のプレイヤー名を入力してください:", "プレイヤー名設定", "プレイヤー1")
    If blackPlayerName = "" Then blackPlayerName = "プレイヤー1"
    
    whitePlayerName = InputBox("白（後手）のプレイヤー名を入力してください:", "プレイヤー名設定", "プレイヤー2")
    If whitePlayerName = "" Then whitePlayerName = "プレイヤー2"
End Sub

' ゲーム機能の設定
Sub SetupGameFeatures()
    Dim response As String

    response = MsgBox("有効手のヒントを表示しますか？", vbYesNo + vbQuestion, "機能設定")
    showHints = (response = vbYes)
End Sub

' ゲームボード初期化
Sub InitializeGameBoard()
    Dim i As Integer, j As Integer, center As Integer

    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            gameBoard(i, j) = CELL_EMPTY
        Next j
    Next i
    
    CurrentPlayer = BLACK
    gameOver = False
    
    ' 初期石配置（中央4マス）
    center = BOARD_SIZE / 2
    gameBoard(center, center) = WHITE
    gameBoard(center + 1, center + 1) = WHITE
    gameBoard(center, center + 1) = BLACK
    gameBoard(center + 1, center) = BLACK
End Sub

' 履歴初期化
Sub InitializeHistory()
    ReDim gameHistory(0 To 400)
    ReDim moveHistory(1 To 400)
    historyCount = 0
    moveCount = 0
End Sub

' UI設定
Sub SetupUI2P()
    Dim i As Integer, j As Integer

    If targetSheet Is Nothing Then
        MsgBox "エラー: ワークシートオブジェクトが設定されていません。", vbCritical
        Exit Sub
    End If
    
    targetSheet.Cells.Clear
    
    With targetSheet
        For i = 1 To BOARD_SIZE
            For j = 1 To BOARD_SIZE
                With .Cells(i, j)
                    .ColumnWidth = 3
                    .RowHeight = 24
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Font.Size = 14
                    .Font.Bold = True
                    .Interior.color = RGB(0, 100, 0)
                    .Borders.LineStyle = xlContinuous
                    .Borders.color = RGB(0, 0, 0)
                End With
            Next j
        Next i

        .Cells(BOARD_SIZE + 2, 1).Font.Size = 12
        .Cells(BOARD_SIZE + 2, 1).Font.Bold = False
        .Cells(BOARD_SIZE + 3, 1).Font.Size = 10
        .Cells(BOARD_SIZE + 4, 1).Font.Size = 10
        .Cells(BOARD_SIZE + 5, 1).Font.Size = 10
        .Cells(BOARD_SIZE + 6, 1).Font.Size = 10
        
        Call CreateGameButtons
        
        Call SetupGameRecordArea
    End With
End Sub

' ゲーム用ボタン作成
Sub CreateGameButtons()
    Call DeleteAllButtons
    
    Call CreateResetButton2P        ' ゲームリセットボタンを配置
    Call CreateManualInputButton2P  ' 手動入力ボタンを配置
    Call CreateHintToggleButton     ' ヒント表示ボタンを配置
    Call CreateHelpButton           ' ヘルプボタンを配置
End Sub

' 棋譜表示エリアの設定
Sub SetupGameRecordArea()
    RECORD_START_COL = BOARD_SIZE + 2
    
    With targetSheet
        .Cells(1, RECORD_START_COL).value = "ターン"
        .Cells(1, RECORD_START_COL + 1).value = "プレイヤー"
        .Cells(1, RECORD_START_COL + 2).value = "座標"
        .Cells(1, RECORD_START_COL + 3).value = "時刻"
        .Cells(1, RECORD_START_COL + 4).value = "考慮時間"
        .Cells(1, RECORD_START_COL + 5).value = "累計時間"

        With .Range(.Cells(1, RECORD_START_COL), .Cells(1, RECORD_START_COL + 5))
            .Font.Bold = False
            .Font.Size = 10
            .Interior.color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns(RECORD_START_COL).ColumnWidth = 6      ' ターン
        .Columns(RECORD_START_COL + 1).ColumnWidth = 10 ' プレイヤー
        .Columns(RECORD_START_COL + 2).ColumnWidth = 6  ' 座標
        .Columns(RECORD_START_COL + 3).ColumnWidth = 8  ' 時刻
        .Columns(RECORD_START_COL + 4).ColumnWidth = 8  ' 考慮時間
        .Columns(RECORD_START_COL + 5).ColumnWidth = 8  ' 累計時間
    End With
End Sub

' リアルタイム棋譜表示初期化
Sub InitializeGameRecordDisplay()
    Dim i As Integer

    With targetSheet
        For i = 2 To BOARD_SIZE + 20
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 5)).ClearContents
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 5)).Interior.color = xlNone
        Next i
    End With
End Sub

' リアルタイム棋譜表示更新
Sub UpdateGameRecordDisplay()
    Dim i As Integer, displayRow As Integer
    Dim moveInfo As String, playerName As String, movePos As String
    Dim timeInfo As String, thinkTime As String, totalTime As String
    Dim moveNum As Integer
    Dim maxDisplayRows As Integer
    
    Call InitializeGameRecordDisplay
    
    maxDisplayRows = BOARD_SIZE + 0
    
    For i = moveCount To 1 Step -1
        displayRow = (moveCount - i) + 2
        
        If displayRow > maxDisplayRows Then Exit For
        
        moveInfo = moveHistory(i)

        moveNum = Val(Left(moveInfo, 3))

        If InStr(moveInfo, "黒") > 0 Then
            playerName = blackPlayerName & "（黒）"
        Else
            playerName = whitePlayerName & "（白）"
        End If

        If InStr(moveInfo, "スキップ") > 0 Then
            movePos = "スキップ"
            thinkTime = "-"
            totalTime = "-"
        Else
            Dim startPos As Integer, endPos As Integer
            startPos = InStr(moveInfo, "）") + 2
            endPos = InStr(startPos, moveInfo, " (") - 1
            movePos = Trim(Mid(moveInfo, startPos, endPos - startPos + 1))
            
            If i <= historyCount And i >= 1 Then
                thinkTime = FormatTime(gameHistory(i).thinkingTime)
                Dim cumTime As Double
                If gameHistory(i).movePlayer = BLACK Then
                    cumTime = CalculateCumulativeTime(i, BLACK)
                Else
                    cumTime = CalculateCumulativeTime(i, WHITE)
                End If
                totalTime = FormatTime(cumTime)
            Else
                thinkTime = "0s"
                totalTime = "0s"
            End If
        End If

        startPos = InStr(moveInfo, "(") + 1
        endPos = InStr(moveInfo, ")") - 1
        timeInfo = Mid(moveInfo, startPos, endPos - startPos + 1)

        With targetSheet
            .Cells(displayRow, RECORD_START_COL).value = moveNum
            .Cells(displayRow, RECORD_START_COL + 1).value = playerName
            .Cells(displayRow, RECORD_START_COL + 2).value = movePos
            .Cells(displayRow, RECORD_START_COL + 3).value = timeInfo
            .Cells(displayRow, RECORD_START_COL + 4).value = thinkTime
            .Cells(displayRow, RECORD_START_COL + 5).value = totalTime
            
            .Range(Cells(2, RECORD_START_COL + 1), Cells(3, RECORD_START_COL + 1)).Columns.AutoFit

            With .Range(.Cells(displayRow, RECORD_START_COL), .Cells(displayRow, RECORD_START_COL + 5))
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin

                If InStr(playerName, "黒") > 0 Then
                    .Interior.color = RGB(240, 240, 240)
                Else
                    .Interior.color = RGB(255, 255, 255)
                End If
            End With
        End With
    Next i
    
    If moveCount > 0 Then
        displayRow = 2
        With targetSheet.Range(targetSheet.Cells(displayRow, RECORD_START_COL), _
                              targetSheet.Cells(displayRow, RECORD_START_COL + 5))
            .Interior.color = RGB(255, 255, 150)
            .Font.Bold = True
        End With
    End If
End Sub

' 時間のフォーマット
Function FormatTime(seconds As Double) As String
    If seconds < 60 Then
        FormatTime = Format(seconds, "0.0") & "s"
    Else
        Dim minutes As Integer
        minutes = Int(seconds / 60)
        seconds = seconds - minutes * 60
        FormatTime = minutes & "m" & Format(seconds, "0") & "s"
    End If
End Function

' 累計時間を計算
Function CalculateCumulativeTime(moveIndex As Integer, Player As Integer) As Double
    Dim cumTime As Double
    Dim i As Integer
    
    For i = 1 To moveIndex
        If i <= historyCount And gameHistory(i).movePlayer = Player Then
            cumTime = cumTime + gameHistory(i).thinkingTime
        End If
    Next i
    
    CalculateCumulativeTime = cumTime
End Function

' リセットボタンを作成
Sub CreateResetButton2P()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 9, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!RestartGame2P"
        .caption = "リセット"
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' ゲーム再開始
Sub RestartGame2P()
    Call CleanupEvents
    Call InitializeGame2P
End Sub

' 手動入力ボタンを作成
Sub CreateManualInputButton2P()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 10, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 10, BOARD_SIZE))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ManualInput2P"
        .caption = "手動入力"
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' 手動入力機能
Sub ManualInput2P()
    Dim userInput As String
    Dim col As Integer, row As Integer
    Dim colChar As String
    
    If gameOver Then
        MsgBox "ゲームは終了しています。", vbInformation
        Exit Sub
    End If
    
    userInput = InputBox("座標を入力してください (例: A4, B3):" & vbCrLf & _
                        "現在のターン: " & GetCurrentPlayerName(), "手動入力", "")
    If userInput = "" Then Exit Sub

    userInput = UCase(Trim(userInput))
    If Len(userInput) < 2 Then
        MsgBox "正しい形式で入力してください (例: A4, B3)", vbExclamation
        Exit Sub
    End If
    
    colChar = Left(userInput, 1)
    row = Val(Mid(userInput, 2))
    col = Asc(colChar) - 64

    If col < 1 Or col > BOARD_SIZE Or row < 1 Or row > BOARD_SIZE Then
        MsgBox "座標が範囲外です。A1から" & Chr(64 + BOARD_SIZE) & BOARD_SIZE & "の範囲で入力してください。", vbExclamation
        Exit Sub
    End If

    If IsValidMove(row, col, CurrentPlayer) Then
        Call ExecutePlayerMove(row, col)
    Else
        MsgBox "そこには置けません。別の場所を選んでください。", vbExclamation
    End If
End Sub

' ヒント表示切り替えボタンを作成
Sub CreateHintToggleButton()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 11, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 11, BOARD_SIZE / 2))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ToggleHints"
        .caption = IIf(showHints, "ヒント: ON", "ヒント: OFF")
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' ヒント表示モード切り替え
Sub ToggleHints()
    showHints = Not showHints

    Call UpdateHintButtonCaption

    Call UpdateDisplay
End Sub

' ヒント表示ボタンのキャプション更新
Sub UpdateHintButtonCaption()
    Dim btn As Button
    Dim i As Integer
    
    On Error Resume Next
    For i = 1 To targetSheet.Buttons.count
        Set btn = targetSheet.Buttons(i)
        If InStr(btn.caption, "ヒント") > 0 Then
            If showHints Then
                btn.caption = "ヒント: ON"
            Else
                btn.caption = "ヒント: OFF"
            End If
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' ヘルプボタンを作成
Sub CreateHelpButton()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 11, BOARD_SIZE / 2 + 1), _
                                    targetSheet.Cells(BOARD_SIZE + 11, BOARD_SIZE))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ShowHelp"
        .caption = "ヘルプ"
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' ヘルプ表示
Sub ShowHelp()
    Dim helpText As String
    helpText = "【オセロ2人対戦 操作方法】" & vbCrLf & vbCrLf & _
               "◆ 石の置き方:" & vbCrLf & _
               "・盤面の空いているマスをクリック" & vbCrLf & _
               "・「手動入力」ボタンで座標入力（例：A4）" & vbCrLf & vbCrLf & _
               "◆ 機能ボタン:" & vbCrLf & _
               "・新しいゲーム：ゲームをリセット" & vbCrLf & _
               "・ヒント：有効手の位置を表示" & vbCrLf & vbCrLf & _
               "◆ 棋譜:" & vbCrLf & _
               "・右側に最近の手が記録されます。" & vbCrLf & _
               "・考慮時間と累計時間も記録されます。" & vbCrLf & _
               "・ゲーム終了時に自動保存されます。" & vbCrLf & vbCrLf & _
               "◆ ルール:" & vbCrLf & _
               "・相手の石を挟んで自分の石にする" & vbCrLf & _
               "・置ける場所がない場合はスキップ" & vbCrLf & _
               "・盤面が埋まるか両者とも置けなくなったら終了"
    
    MsgBox helpText, vbInformation, "ヘルプ"
End Sub

' イベントハンドラをクリーンアップ
Sub CleanupEvents()
    On Error Resume Next
    Application.OnSheetSelectionChange = ""
    On Error GoTo 0
End Sub

' 全ボタンを削除
Sub DeleteAllButtons()
    Dim i As Integer
    
    On Error Resume Next
    
    If Not targetSheet Is Nothing Then
        For i = targetSheet.Buttons.count To 1 Step -1
            targetSheet.Buttons(i).Delete
        Next i
    End If
    
    On Error GoTo 0
End Sub

' 画面表示更新
Sub UpdateDisplay()
    Dim i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            With targetSheet.Cells(i, j)
                Select Case gameBoard(i, j)
                    Case CELL_EMPTY
                        .value = ""
                        .Font.color = RGB(0, 0, 0)
                        .Font.Size = 14

                        If showHints And Not gameOver Then
                            If IsValidMove(i, j, CurrentPlayer) Then
                                .value = "●"
                                .Font.Size = 10
                                .Font.color = RGB(100, 100, 100)
                                .Font.Bold = False
                            End If
                        End If
                        
                    Case BLACK
                        .value = "●"
                        .Font.color = RGB(0, 0, 0)
                        .Font.Size = 14
                        .Font.Bold = True
                    Case WHITE
                        .value = "●"
                        .Font.color = RGB(255, 255, 255)
                        .Font.Size = 14
                        .Font.Bold = True
                End Select
            End With
        Next j
    Next i
    
    Call ShowScore
End Sub

' メッセージ表示
Sub ShowMessage(msg As String)
    targetSheet.Cells(BOARD_SIZE + 2, 1).value = msg
End Sub

' スコア表示
Sub ShowScore()
    Dim blackCount As Integer, whiteCount As Integer
    Dim i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = BLACK Then blackCount = blackCount + 1
            If gameBoard(i, j) = WHITE Then whiteCount = whiteCount + 1
        Next j
    Next i
    
    targetSheet.Cells(BOARD_SIZE + 3, 1).value = _
        "スコア - " & blackPlayerName & "（黒）: " & blackCount & "  " & whitePlayerName & "（白）: " & whiteCount
    
    targetSheet.Cells(BOARD_SIZE + 4, 1).value = "ターン: " & moveCount

    Dim statusText As String
    statusText = "機能: "
    If showHints Then statusText = statusText & "ヒントON"
    targetSheet.Cells(BOARD_SIZE + 5, 1).value = statusText
End Sub

' ゲーム状態保存
Sub SaveGameState(row As Integer, col As Integer, Player As Integer, thinkingTime As Double)
    historyCount = historyCount + 1

    If historyCount > UBound(gameHistory) Then
        ReDim Preserve gameHistory(0 To historyCount + 50)
    End If
    
    With gameHistory(historyCount)
        Dim i As Integer, j As Integer
        For i = 1 To BOARD_SIZE
            For j = 1 To BOARD_SIZE
                .board(i, j) = gameBoard(i, j)
            Next j
        Next i
        
        .Player = CurrentPlayer
        .moveRow = row
        .moveCol = col
        .movePlayer = Player
        .timestamp = Format(Now, "hh:mm:ss")
        .thinkingTime = thinkingTime
    End With
End Sub

' 手の記録
Sub RecordMove(row As Integer, col As Integer, Player As Integer, thinkingTime As Double)
    moveCount = moveCount + 1
    
    If moveCount > UBound(moveHistory) Then
        ReDim Preserve moveHistory(1 To moveCount + 50)
    End If
    
    Dim playerName As String
    If Player = BLACK Then
        playerName = blackPlayerName & "（黒）"
    Else
        playerName = whitePlayerName & "（白）"
    End If
    
    Dim moveDescription As String
    If row = 0 And col = 0 Then
        moveDescription = "スキップ"
    Else
        Dim colLetter As String
        colLetter = Chr(64 + col)
        moveDescription = colLetter & row
    End If
    
    moveHistory(moveCount) = Format(moveCount, "000") & ": " & playerName & " " & _
                            moveDescription & " (" & Format(Now, "hh:mm:ss") & ")"
End Sub

' 棋譜保存
Sub SaveGameRecord2P()
    Dim recordSheet As Worksheet
    Dim lastRow As Integer, i As Integer
    Dim gameStartTime As String
    gameStartTime = Format(Now, "yyyy/mm/dd hh:mm:ss")

    On Error Resume Next
    Set recordSheet = targetWorkbook.Worksheets("棋譜")
    On Error GoTo 0
    
    If recordSheet Is Nothing Then
        Set recordSheet = targetWorkbook.Worksheets.Add
        recordSheet.Name = "棋譜"

        recordSheet.Cells(1, 1).value = "ゲームID"
        recordSheet.Cells(1, 2).value = "ターン"
        recordSheet.Cells(1, 3).value = "プレイヤー"
        recordSheet.Cells(1, 4).value = "座標"
        recordSheet.Cells(1, 5).value = "時刻"
        recordSheet.Cells(1, 6).value = "考慮時間"
        recordSheet.Cells(1, 7).value = "累計時間"
        recordSheet.Cells(1, 8).value = "ボードサイズ"
        recordSheet.Cells(1, 9).value = "スコア"
        recordSheet.Cells(1, 10).value = "勝者"

        With recordSheet.Range("A1:J1")
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With

        recordSheet.Columns("A").ColumnWidth = 18  ' ゲームID
        recordSheet.Columns("B").ColumnWidth = 6   ' ターン
        recordSheet.Columns("C").ColumnWidth = 12  ' プレイヤー
        recordSheet.Columns("D").ColumnWidth = 8   ' 座標
        recordSheet.Columns("E").ColumnWidth = 10  ' 時刻
        recordSheet.Columns("F").ColumnWidth = 10  ' 考慮時間
        recordSheet.Columns("G").ColumnWidth = 10  ' 累計時間
        recordSheet.Columns("H").ColumnWidth = 10  ' ボードサイズ
        recordSheet.Columns("I").ColumnWidth = 15  ' スコア
        recordSheet.Columns("J").ColumnWidth = 15  ' 勝者
    End If

    Dim finalBlackCount As Integer, finalWhiteCount As Integer
    Dim winner As String
    
    For i = 1 To BOARD_SIZE
        Dim j As Integer
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = BLACK Then finalBlackCount = finalBlackCount + 1
            If gameBoard(i, j) = WHITE Then finalWhiteCount = finalWhiteCount + 1
        Next j
    Next i
    
    If finalBlackCount > finalWhiteCount Then
        winner = blackPlayerName & "（黒）"
    ElseIf finalWhiteCount > finalBlackCount Then
        winner = whitePlayerName & "（白）"
    Else
        winner = "引き分け"
    End If

    lastRow = recordSheet.Cells(recordSheet.Rows.count, 1).End(xlUp).row
    
    For i = 1 To moveCount
        lastRow = lastRow + 1

        Dim moveInfo As String, playerName As String, movePos As String
        Dim timeInfo As String, moveNum As Integer
        Dim thinkTime As String, totalTime As String
        Dim currentScore As String

        moveInfo = moveHistory(i)

        moveNum = Val(Left(moveInfo, 3))

        If InStr(moveInfo, "黒") > 0 Then
            playerName = blackPlayerName & "（黒）"
        Else
            playerName = whitePlayerName & "（白）"
        End If

        If InStr(moveInfo, "スキップ") > 0 Then
            movePos = "スキップ"
            thinkTime = "-"
            totalTime = "-"
        Else

            Dim startPos As Integer, endPos As Integer
            startPos = InStr(moveInfo, "）") + 2
            endPos = InStr(startPos, moveInfo, " (") - 1
            movePos = Trim(Mid(moveInfo, startPos, endPos - startPos + 1))

            If i <= historyCount Then
                thinkTime = FormatTime(gameHistory(i).thinkingTime)

                Dim cumTime As Double
                If gameHistory(i).movePlayer = BLACK Then
                    cumTime = CalculateCumulativeTime(i, BLACK)
                Else
                    cumTime = CalculateCumulativeTime(i, WHITE)
                End If
                totalTime = FormatTime(cumTime)
            Else
                thinkTime = "0s"
                totalTime = "0s"
            End If
        End If

        startPos = InStr(moveInfo, "(") + 1
        endPos = InStr(moveInfo, ")") - 1
        timeInfo = Mid(moveInfo, startPos, endPos - startPos + 1)

        If i <= historyCount Then
            currentScore = CalculateScoreAtMove2P(i)
        Else
            currentScore = "0-0"
        End If

        With recordSheet
            .Cells(lastRow, 1).value = gameStartTime           ' ゲームID
            .Cells(lastRow, 2).value = moveNum                 ' ターン
            .Cells(lastRow, 3).value = playerName              ' プレイヤー
            .Cells(lastRow, 4).value = movePos                 ' 座標
            .Cells(lastRow, 5).value = timeInfo                ' 時刻
            .Cells(lastRow, 6).value = thinkTime               ' 考慮時間
            .Cells(lastRow, 7).value = totalTime               ' 累計時間
            .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE  ' ボードサイズ
            .Cells(lastRow, 9).value = currentScore            ' その時点でのスコア
            .Cells(lastRow, 10).value = IIf(i = moveCount, winner, "-")  ' 最終手のみ勝者を表示
        End With
    Next i

    lastRow = lastRow + 1
    With recordSheet
        .Cells(lastRow, 1).value = gameStartTime
        .Cells(lastRow, 2).value = "---"
        .Cells(lastRow, 3).value = "ゲーム終了"
        .Cells(lastRow, 4).value = "---"
        .Cells(lastRow, 5).value = Format(Now, "hh:mm:ss")
        .Cells(lastRow, 6).value = ""
        .Cells(lastRow, 7).value = ""
        .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE
        .Cells(lastRow, 9).value = blackPlayerName & ":" & finalBlackCount & "-" & finalWhiteCount & ":" & whitePlayerName
        .Cells(lastRow, 10).value = winner

        .Range(.Cells(lastRow, 1), .Cells(lastRow, 10)).Interior.color = RGB(240, 240, 240)
        .Range(.Cells(lastRow, 1), .Cells(lastRow, 10)).Font.Bold = True
    End With

    lastRow = lastRow + 1
    
    MsgBox "棋譜を保存しました。" & vbCrLf & _
           "シート「棋譜」で確認できます。" & vbCrLf, vbInformation, "棋譜保存完了"
End Sub

' 指定した手番時点でのスコアを計算する関数
Function CalculateScoreAtMove2P(moveIndex As Integer) As String
    Dim blackCount As Integer, whiteCount As Integer
    Dim i As Integer, j As Integer

    If moveIndex < 1 Or moveIndex > historyCount Then
        CalculateScoreAtMove2P = "エラー"
        Exit Function
    End If

    With gameHistory(moveIndex)
        For i = 1 To BOARD_SIZE
            For j = 1 To BOARD_SIZE
                If .board(i, j) = BLACK Then
                    blackCount = blackCount + 1
                ElseIf .board(i, j) = WHITE Then
                    whiteCount = whiteCount + 1
                End If
            Next j
        Next i
    End With

    CalculateScoreAtMove2P = blackPlayerName & ":" & blackCount & "-" & whiteCount & ":" & whitePlayerName
End Function

' 指定プレイヤーに有効な手があるかチェック
Function HasValidMoves(Player As Integer) As Boolean
    Dim i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If IsValidMove(i, j, Player) Then
                HasValidMoves = True
                Exit Function
            End If
        Next j
    Next i
    
    HasValidMoves = False
End Function

' 有効手判定
Function IsValidMove(row As Integer, col As Integer, Player As Integer) As Boolean
    If row < 1 Or row > BOARD_SIZE Or col < 1 Or col > BOARD_SIZE Then Exit Function
    If gameBoard(row, col) <> CELL_EMPTY Then Exit Function
    
    Dim directions As Variant
    directions = Array(Array(-1, -1), Array(-1, 0), Array(-1, 1), _
                      Array(0, -1), Array(0, 1), _
                      Array(1, -1), Array(1, 0), Array(1, 1))
    
    Dim i As Integer, dr As Integer, dc As Integer
    Dim r As Integer, c As Integer
    Dim opponent As Integer
    
    opponent = 3 - Player
    
    For i = 0 To 7
        dr = directions(i)(0)
        dc = directions(i)(1)
        r = row + dr
        c = col + dc
        
        If r >= 1 And r <= BOARD_SIZE And c >= 1 And c <= BOARD_SIZE Then
            If gameBoard(r, c) = opponent Then
                Do
                    r = r + dr
                    c = c + dc
                    If r < 1 Or r > BOARD_SIZE Or c < 1 Or c > BOARD_SIZE Then Exit Do
                    If gameBoard(r, c) = CELL_EMPTY Then Exit Do
                    If gameBoard(r, c) = Player Then
                        IsValidMove = True
                        Exit Function
                    End If
                Loop
            End If
        End If
    Next i
End Function

' 手を実行
Sub MakeMove(row As Integer, col As Integer, Player As Integer)
    gameBoard(row, col) = Player
    
    Dim directions As Variant
    directions = Array(Array(-1, -1), Array(-1, 0), Array(-1, 1), _
                      Array(0, -1), Array(0, 1), _
                      Array(1, -1), Array(1, 0), Array(1, 1))
    
    Dim i As Integer, dr As Integer, dc As Integer
    Dim r As Integer, c As Integer
    Dim opponent As Integer
    
    opponent = 3 - Player
    
    For i = 0 To 7
        dr = directions(i)(0)
        dc = directions(i)(1)
        r = row + dr
        c = col + dc
        
        If r >= 1 And r <= BOARD_SIZE And c >= 1 And c <= BOARD_SIZE Then
            If gameBoard(r, c) = opponent Then
                Dim flipPositions() As Integer
                Dim flipCount As Integer
                ReDim flipPositions(1 To BOARD_SIZE * 2, 1 To 2)
                flipCount = 0
                
                Do
                    flipCount = flipCount + 1
                    flipPositions(flipCount, 1) = r
                    flipPositions(flipCount, 2) = c
                    r = r + dr
                    c = c + dc
                    
                    If r < 1 Or r > BOARD_SIZE Or c < 1 Or c > BOARD_SIZE Then Exit Do
                    If gameBoard(r, c) = CELL_EMPTY Then Exit Do
                    If gameBoard(r, c) = Player Then
                        Dim j As Integer
                        For j = 1 To flipCount
                            gameBoard(flipPositions(j, 1), flipPositions(j, 2)) = Player
                        Next j
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next i
End Sub

' ゲーム終了判定
Function CheckGameEnd() As Boolean
    Dim hasBlackMove As Boolean, hasWhiteMove As Boolean
    
    hasBlackMove = HasValidMoves(BLACK)
    hasWhiteMove = HasValidMoves(WHITE)
    
    CheckGameEnd = Not (hasBlackMove Or hasWhiteMove)
    gameOver = CheckGameEnd
End Function

' 結果表示
Sub ShowResult()
    Dim blackCount As Integer, whiteCount As Integer
    Dim i As Integer, j As Integer, result As String

    Call CleanupEvents
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = BLACK Then blackCount = blackCount + 1
            If gameBoard(i, j) = WHITE Then whiteCount = whiteCount + 1
        Next j
    Next i
    
    If blackCount > whiteCount Then
        result = blackPlayerName & "（黒）の勝ちです！"
    ElseIf whiteCount > blackCount Then
        result = whitePlayerName & "（白）の勝ちです！"
    Else
        result = "引き分けです！"
    End If
    
    Call ShowMessage(result)
    targetSheet.Cells(BOARD_SIZE + 6, 1).value = _
        "最終スコア - " & blackPlayerName & "（黒）: " & blackCount & "  " & whitePlayerName & "（白）: " & whiteCount

    Call UpdateGameRecordDisplay
    Call SaveGameRecord2P
    
    Dim finalMessage As String
    finalMessage = result & vbCrLf & vbCrLf & _
           "最終スコア" & vbCrLf & _
           blackPlayerName & "（黒）: " & blackCount & vbCrLf & _
           whitePlayerName & "（白）: " & whiteCount & vbCrLf & vbCrLf & _
           "総考慮時間" & vbCrLf & _
           blackPlayerName & ": " & FormatTime(blackTotalTime) & vbCrLf & _
           whitePlayerName & ": " & FormatTime(whiteTotalTime) & vbCrLf & vbCrLf & _
           "棋譜を自動保存しました。" & vbCrLf & _
           "新しいゲームを始めるには「新しいゲーム」ボタンをクリックしてください。"
    
    MsgBox finalMessage, vbInformation, "ゲーム終了"
End Sub

