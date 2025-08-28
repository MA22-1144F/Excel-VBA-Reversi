Attribute VB_Name = "Othello_2025"
' ==========================================================
' VBA オセロゲーム
' ==========================================================

Option Explicit

' ゲーム設定
Public Const BOARD_SIZE As Integer = 8    ' ボードサイズ（偶数推奨）
Public Const CELL_EMPTY As Integer = 0    ' 空のセル
Public Const BLACK As Integer = 1         ' プレイヤー（黒）
Public Const WHITE As Integer = 2         ' CPU（白）
Public showEvaluationMode As Boolean  ' 評価表示モードのオン/オフ

' 棋譜表示設定
Public RECORD_START_COL As Integer             ' 棋譜表示開始列
Public Const RECORD_WIDTH As Integer = 5       ' 棋譜表示幅（列数）

' 難易度設定
Public Enum DifficultyLevel
    EASY = 1        ' 簡単（評価関数のみ）
    MEDIUM = 2      ' 普通（α-β探索 深度3）
    HARD = 3        ' 難しい（α-β探索 深度5）
End Enum

' ゲーム状態保存用
Type GameState
    board(1 To 20, 1 To 20) As Integer
    Player As Integer
    moveRow As Integer
    moveCol As Integer
    movePlayer As Integer
    timestamp As String
    evaluation As Integer
End Type

' 有効手格納用
Type ValidMove
    row As Integer
    col As Integer
    Score As Integer
End Type

' ゲーム段階の定義
Public Const PHASE_OPENING As Integer = 1     ' 序盤（～20手）
Public Const PHASE_MIDGAME As Integer = 2     ' 中盤（21～45手）
Public Const PHASE_ENDGAME As Integer = 3     ' 終盤（46手～）

' ゲーム変数
Public gameBoard(1 To 20, 1 To 20) As Integer
Public CurrentPlayer As Integer
Public gameOver As Boolean
Public targetSheet As Worksheet
Public targetWorkbook As Workbook
Public difficulty As DifficultyLevel
Public gameHistory() As GameState
Public historyCount As Integer
Public moveHistory() As String
Public moveCount As Integer

' リボンから呼び出すメイン関数
Sub StartReversiGame()
    On Error GoTo ErrorHandler
    
    ' 新しいワークブックを作成してゲームを開始
    Call CreateNewGameWorkbook
    
    If targetWorkbook Is Nothing Or targetSheet Is Nothing Then
        MsgBox "ワークブックまたはワークシートの作成に失敗しました。再度実行してください。", vbCritical
        Exit Sub
    End If
    
    Call InitializeGame
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
    
    ' ワークシート名を設定
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
               "        Application.Run """ & ThisWorkbook.Name & "!ProcessCellClick"", Target" & vbCrLf & _
               "    End If" & vbCrLf & _
               "End Sub"
    
    CodeModule.AddFromString eventCode

    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange"
    
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

' 簡易版：Application.OnSheetSelectionChangeのみを使用
Sub SetupSimpleEvents()
    On Error Resume Next

    Application.OnSheetSelectionChange = ""
    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange"
    
    On Error GoTo 0
End Sub

' アプリケーションレベルのSelectionChangeイベント
Public Sub OnSheetSelectionChange(Sh As Object, Target As Range)
    On Error Resume Next

    If Sh Is Nothing Or Target Is Nothing Then Exit Sub
    If Target.Cells.count <> 1 Then Exit Sub

    If targetSheet Is Nothing Then Exit Sub
    If gameOver Then Exit Sub
    If CurrentPlayer <> BLACK Then Exit Sub

    If Sh.Name <> targetSheet.Name Then Exit Sub

    If Sh.Parent.Name <> targetSheet.Parent.Name Then Exit Sub
    
    Call ProcessCellClick(Target)
End Sub

' セルクリック処理の関数
Public Sub ProcessCellClick(Target As Range)
    Dim row As Integer, col As Integer
    
    On Error Resume Next

    If Not Target.Worksheet Is targetSheet Then Exit Sub
    
    row = Target.row
    col = Target.Column

    If row < 1 Or row > BOARD_SIZE Or col < 1 Or col > BOARD_SIZE Then Exit Sub

    If IsValidMove(row, col, BLACK) Then
        Dim playerEvaluation As Integer
        playerEvaluation = GetUnifiedEvaluation(row, col, BLACK)

        Call SaveGameState(row, col, BLACK, playerEvaluation)
        Call MakeMove(row, col, BLACK)
        Call UpdateDisplay

        Call RecordMove(row, col, BLACK)
        Call UpdateGameRecordDisplay
        
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If

        Call ExecuteCPUTurn
        
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If

        Call SwitchToPlayerTurn
    Else
        Call ShowMessage("そこには置けません。別の場所を選んでください。")
    End If
End Sub

' ワークシートイベントから呼び出される関数
Public Sub HandleCellClickFromEvent(Target As Range)
    Call ProcessCellClick(Target)
End Sub

' ゲーム初期化
Sub InitializeGame()
    Dim response As String
    Dim diffLevel As Integer
    
    On Error GoTo InitError
    
    ' 評価表示モードを初期化（デフォルトはオフ）
    showEvaluationMode = False
    
    ' 難易度選択
    response = InputBox("難易度を選択してください:" & vbCrLf & _
                       "1: 初級（評価関数のみ）" & vbCrLf & _
                       "2: 中級（α-β探索 深度3）" & vbCrLf & _
                       "3: 上級（α-β探索 深度5）", _
                       "難易度選択", "2")
    
    If response = "" Then
        ' キャンセルされた場合はワークブックを閉じる
        If Not targetWorkbook Is Nothing Then
            targetWorkbook.Close False
        End If
        Exit Sub
    End If
    
    diffLevel = Val(response)
    If diffLevel < 1 Or diffLevel > 3 Then diffLevel = 2
    difficulty = diffLevel

    Call InitializeGameBoard
    Call InitializeHistory
    Call SetupUI
    Call InitializeGameRecordDisplay
    Call UpdateDisplay
    Call UpdateGameRecordDisplay

    Call CheckAndHandleSkip
    If Not gameOver Then
        Call ShowMessage("黒のターンです。置きたい場所をクリックしてください。")
    End If

    Call SetupWorksheetEvents
    
    Dim diffName As String
    Select Case difficulty
        Case EASY: diffName = "初級"
        Case MEDIUM: diffName = "中級"
        Case HARD: diffName = "上級"
    End Select
    
    MsgBox "オセロゲームを開始しました。" & vbCrLf & _
           "ボードサイズ: " & BOARD_SIZE & "x" & BOARD_SIZE & vbCrLf & _
           "難易度: " & diffName & vbCrLf & _
           "黒のターンです。" & vbCrLf & vbCrLf & _
           "セルクリックまたは手動入力でプレイしてください。", vbInformation
    
    Exit Sub
    
InitError:
    MsgBox "ゲーム初期化中にエラーが発生しました: " & Err.description & vbCrLf & vbCrLf & _
           "「手動入力」ボタンでプレイを続行できます。", vbExclamation
    Resume Next
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
    ReDim gameHistory(0 To 400)  ' 最大400手（20x20対応）
    ReDim moveHistory(1 To 400)
    historyCount = 0
    moveCount = 0
    
    ' 初期状態を保存
    Call SaveGameState(0, 0, 0, 0)
End Sub

' UI設定
Sub SetupUI()
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
        
        ' ボタン類を配置
        Call CreateResetButton      ' ゲームリセットボタンを配置
        Call CreateManualInputButton        ' 手動入力ボタンを配置
        Call CreateEvaluationToggleButton       ' 評価値表示ボタンを配置
        
        ' 棋譜表示エリアの設定
        Call SetupGameRecordArea
    End With
End Sub

' 棋譜表示エリアの設定
Sub SetupGameRecordArea()
    RECORD_START_COL = BOARD_SIZE + 2
    
    With targetSheet
        ' 棋譜ヘッダー行（1行目に配置）
        .Cells(1, RECORD_START_COL).value = "ターン"
        .Cells(1, RECORD_START_COL + 1).value = "プレイヤー"
        .Cells(1, RECORD_START_COL + 2).value = "座標"
        .Cells(1, RECORD_START_COL + 3).value = "時刻"
        .Cells(1, RECORD_START_COL + 4).value = "評価"
        
        ' ヘッダー行の書式設定
        With .Range(.Cells(1, RECORD_START_COL), .Cells(1, RECORD_START_COL + 4))
            .Font.Bold = False
            .Font.Size = 10
            .Interior.color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' 列幅設定
        .Columns(RECORD_START_COL).ColumnWidth = 6      ' ターン
        .Columns(RECORD_START_COL + 1).ColumnWidth = 8  ' プレイヤー
        .Columns(RECORD_START_COL + 2).ColumnWidth = 6  ' 座標
        .Columns(RECORD_START_COL + 3).ColumnWidth = 8  ' 時刻
        .Columns(RECORD_START_COL + 4).ColumnWidth = 6  ' 評価
    End With
End Sub

' リアルタイム棋譜表示初期化
Sub InitializeGameRecordDisplay()
    Dim i As Integer
    
    ' 既存の棋譜データをクリア（ヘッダーは保持）
    With targetSheet
        For i = 2 To BOARD_SIZE + 20
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 4)).ClearContents
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 4)).Interior.color = xlNone
        Next i
    End With
End Sub

' リアルタイム棋譜表示更新
Sub UpdateGameRecordDisplay()
    Dim i As Integer, displayRow As Integer
    Dim moveInfo As String, playerName As String, movePos As String
    Dim timeInfo As String, moveNum As Integer, evalValue As Integer
    Dim startPos As Integer, endPos As Integer
    Dim maxDisplayRows As Integer
    
    ' 既存表示をクリア
    Call InitializeGameRecordDisplay
    
    ' 表示可能行数を計算（ボードサイズに応じて）
    maxDisplayRows = BOARD_SIZE + 15
    
    ' 各手を逆順で表示（最新手が上、最初の手が下）
    For i = moveCount To 1 Step -1
        displayRow = (moveCount - i) + 2  ' ヘッダー分をオフセット
        
        If displayRow > maxDisplayRows Then Exit For  ' 表示範囲を超えた場合
        
        moveInfo = moveHistory(i)
        moveNum = Val(Left(moveInfo, 3))
        If InStr(moveInfo, "黒") > 0 Then
            playerName = "黒"
        Else
            playerName = "白"
        End If
        If InStr(moveInfo, "スキップ") > 0 Then
            movePos = "スキップ"
            evalValue = 0
        Else

            startPos = InStr(moveInfo, playerName) + Len(playerName) + 1
            endPos = InStr(startPos, moveInfo, " (") - 1
            movePos = Trim(Mid(moveInfo, startPos, endPos - startPos + 1))
            
            If i <= historyCount Then
                evalValue = gameHistory(i).evaluation
            Else
                evalValue = 0
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
            .Cells(displayRow, RECORD_START_COL + 4).value = IIf(evalValue = 0 And movePos = "スキップ", "-", evalValue)

            With .Range(.Cells(displayRow, RECORD_START_COL), .Cells(displayRow, RECORD_START_COL + 4))
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin

                If playerName = "黒" Then
                    .Interior.color = RGB(240, 240, 240)
                Else
                    .Interior.color = RGB(255, 255, 255)
                End If
            End With
        End With
    Next i
    
    ' 最新手をハイライト
    If moveCount > 0 Then
        displayRow = 2  ' 最新手は常に2行目
        With targetSheet.Range(targetSheet.Cells(displayRow, RECORD_START_COL), _
                              targetSheet.Cells(displayRow, RECORD_START_COL + 4))
            .Interior.color = RGB(255, 255, 150)  ' 黄色ハイライト
            .Font.Bold = True
        End With
    End If
End Sub

' 現在のゲーム段階を判定
Function GetGamePhase() As Integer
    Dim totalMoves As Integer
    totalMoves = CountTotalStones() - 4  ' 初期4石を除く
    
    If totalMoves <= 20 Then
        GetGamePhase = PHASE_OPENING
    ElseIf totalMoves <= 45 Then
        GetGamePhase = PHASE_MIDGAME
    Else
        GetGamePhase = PHASE_ENDGAME
    End If
End Function

' 盤上の石の総数をカウント
Function CountTotalStones() As Integer
    Dim i As Integer, j As Integer, count As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) <> CELL_EMPTY Then
                count = count + 1
            End If
        Next j
    Next i
    
    CountTotalStones = count
End Function

' 評価関数
Function GetUnifiedEvaluation(row As Integer, col As Integer, Player As Integer) As Integer
    Dim evaluation As Integer
    Dim tempBoard(1 To 20, 1 To 20) As Integer
    Dim phase As Integer
    
    ' 有効手でない場合は0を返す
    If Not IsValidMove(row, col, Player) Then
        GetUnifiedEvaluation = 0
        Exit Function
    End If

    Call CopyBoard(gameBoard, tempBoard)

    Call MakeMove(row, col, Player)

    phase = GetGamePhase()

    Select Case phase
        Case PHASE_OPENING
            evaluation = EvaluateOpening(row, col, Player)
        Case PHASE_MIDGAME
            evaluation = EvaluateMidgame(row, col, Player)
        Case PHASE_ENDGAME
            evaluation = EvaluateEndgame(row, col, Player)
    End Select

    Call CopyBoard(tempBoard, gameBoard)

    GetUnifiedEvaluation = NormalizeEvaluationImproved(evaluation, phase)
End Function

' 正規化処理
Function NormalizeEvaluationSmooth(rawEval As Integer, phase As Integer) As Integer
    Dim normalizedValue As Single
    
    Select Case phase
        Case PHASE_OPENING
            ' 序盤：機動力重視、角の価値を強調
            normalizedValue = SigmoidNormalization(rawEval, 800, 1.2)
            
        Case PHASE_MIDGAME
            ' 中盤：バランス重視
            normalizedValue = SigmoidNormalization(rawEval, 1000, 1)
            
        Case PHASE_ENDGAME
            ' 終盤：細かい差を重視、より敏感な反応
            normalizedValue = TanhNormalization(rawEval, 600, 0.8)
    End Select
    
    ' -100～+100の範囲に制限
    If normalizedValue > 100 Then
        NormalizeEvaluationSmooth = 100
    ElseIf normalizedValue < -100 Then
        NormalizeEvaluationSmooth = -100
    Else
        NormalizeEvaluationSmooth = Round(normalizedValue, 0)
    End If
End Function

' シグモイド関数による正規化
Function SigmoidNormalization(value As Integer, scaleParam As Single, sensitivity As Single) As Single

    Dim x As Single
    Dim result As Single
    
    ' ゼロ除算を防ぐ
    If scaleParam = 0 Then scaleParam = 1
    
    x = value / scaleParam
    
    ' Exp関数の引数が大きすぎる場合の対策
    If x > 50 Then
        result = 1
    ElseIf x < -50 Then
        result = -1
    Else
        result = 2 / (1 + Exp(-x)) - 1
    End If
    
    ' 100倍して -100～+100 の範囲にし、感度を調整
    SigmoidNormalization = result * 100 * sensitivity
End Function

' ハイパーボリックタンジェント関数による正規化
Function TanhNormalization(value As Integer, scaleParam As Single, sensitivity As Single) As Single
    Dim x As Single
    Dim result As Single
    Dim expX As Single, expNegX As Single
    
    ' ゼロ除算を防ぐ
    If scaleParam = 0 Then scaleParam = 1
    
    x = value / scaleParam
    
    ' Exp関数の引数が大きすぎる場合の対策
    If x > 50 Then
        result = 1
    ElseIf x < -50 Then
        result = -1
    Else
        expX = Exp(x)
        expNegX = Exp(-x)
        result = (expX - expNegX) / (expX + expNegX)
    End If
    
    ' 100倍して -100～+100 の範囲にし、感度を調整
    TanhNormalization = result * 100 * sensitivity
End Function

' より高度な正規化（複数の関数を組み合わせ）
Function AdvancedNormalization(value As Integer, phase As Integer) As Integer
    Dim result As Double
    Dim absValue As Double
    absValue = Abs(value)
    
    Select Case phase
        Case PHASE_OPENING
            ' 序盤：角の価値を強調する非線形関数
            If absValue <= 100 Then
                ' 小さい値は線形
                result = value * 0.3
            ElseIf absValue <= 1000 Then
                ' 中程度の値は対数的圧縮
                result = Sgn(value) * (30 + 40 * Log(absValue / 100) / Log(10))
            Else
                ' 大きい値（角など）は指数的に強調
                result = Sgn(value) * (70 + 30 * (1 - Exp(-(absValue - 1000) / 2000)))
            End If
            
        Case PHASE_MIDGAME
            ' 中盤：バランスの取れた評価
            If absValue <= 200 Then
                result = value * 0.2
            Else
                result = Sgn(value) * (40 + 60 * (1 - Exp(-absValue / 1500)))
            End If
            
        Case PHASE_ENDGAME
            ' 終盤：石数差を重視した線形性の強い関数
            result = 100 * (1 - Exp(-Abs(value) / 1200)) * Sgn(value)
    End Select
    
    ' -100～+100の範囲に制限
    If result > 100 Then
        AdvancedNormalization = 100
    ElseIf result < -100 Then
        AdvancedNormalization = -100
    Else
        AdvancedNormalization = Round(result, 0)
    End If
End Function

' パラメータ調整可能な正規化関数
Function ParametricNormalization(value As Integer, phase As Integer) As Integer
    Dim scaleParam As Double, steepness As Double, threshold As Double
    Dim result As Double
    
    ' 段階別パラメータ設定
    Select Case phase
        Case PHASE_OPENING
            scaleParam = 1200      ' より大きな値まで考慮
            steepness = 0.8   ' やや緩やかな変化
            threshold = 100   ' 閾値
            
        Case PHASE_MIDGAME
            scaleParam = 1000      ' 標準的なスケール
            steepness = 1#    ' 標準的な急峻さ
            threshold = 150   ' 中程度の閾値
            
        Case PHASE_ENDGAME
            scaleParam = 800       ' より敏感に反応
            steepness = 1.2   ' より急峻な変化
            threshold = 200   ' より高い閾値
    End Select
    
    ' 調整可能なシグモイド関数
    Dim adjustedValue As Double
    adjustedValue = (value - threshold) * steepness / scaleParam
    
    ' tanh関数で正規化
    Dim expPos As Double, expNeg As Double
    expPos = Exp(adjustedValue)
    expNeg = Exp(-adjustedValue)
    result = 100 * (expPos - expNeg) / (expPos + expNeg)
    
    ' 範囲制限
    If result > 100 Then
        ParametricNormalization = 100
    ElseIf result < -100 Then
        ParametricNormalization = -100
    Else
        ParametricNormalization = Round(result, 0)
    End If
End Function

' 正規化方式を選択する関数
Function NormalizeEvaluationImproved(rawEval As Integer, phase As Integer) As Integer
    ' 3つの正規化方式から選択可能
    ' 1. 滑らかなシグモイド/tanh関数
    ' NormalizeEvaluationImproved = NormalizeEvaluationSmooth(rawEval, phase)
    ' 2. より高度な複合関数
    NormalizeEvaluationImproved = AdvancedNormalization(rawEval, phase)
    ' 3. 調整可能なパラメータ付き関数
    ' NormalizeEvaluationImproved = ParametricNormalization(rawEval, phase)
End Function


' 序盤評価（機動力重視、石数は控えめに）
Function EvaluateOpening(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' 1. 角の価値
    If IsCorner(row, col) Then
        Score = Score + 5000
    End If
    
    ' 2. 相手に角を与える手への厳罰
    Dim cornerGift As Integer
    cornerGift = CheckCornerGift(row, col, Player)
    If cornerGift > 0 Then
        Score = Score - (cornerGift * 3000)  ' 1つの角で-3000点
    End If
    
    ' 3. 危険な角周辺への厳罰
    Dim dangerLevel As Integer
    dangerLevel = GetCornerDangerLevel(row, col)
    Score = Score - (dangerLevel * 200)
    
    ' 4. 機動力の評価
    Dim myMobility As Integer, oppMobility As Integer
    myMobility = CountValidMoves(Player)
    oppMobility = CountValidMoves(opponent)
    Score = Score + (myMobility - oppMobility) * 100
    
    ' 5. 中央制御の価値
    If IsCenterRegion(row, col) Then
        Score = Score + 80
    End If
    
    ' 6. 石数は控えめに
    Dim myStones As Integer, oppStones As Integer
    Call CountStones(myStones, oppStones, Player)
    Score = Score - (myStones - oppStones) * 30
    
    ' 7. ひっくり返す石数
    Dim flippedCount As Integer
    flippedCount = CountFlips(row, col, Player)
    Score = Score + flippedCount * 15
    
    EvaluateOpening = Score
End Function

' 中盤評価（バランス重視）
Function EvaluateMidgame(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' 1. 角の価値
    If IsCorner(row, col) Then
        Score = Score + 4000
    End If
    
    ' 2. 相手に角を与える手への厳罰
    Dim cornerGift As Integer
    cornerGift = CheckCornerGift(row, col, Player)
    If cornerGift > 0 Then
        Score = Score - (cornerGift * 2500)
    End If
    
    ' 3. 危険な角周辺への罰則
    Dim dangerLevel As Integer
    dangerLevel = GetCornerDangerLevel(row, col)
    Score = Score - (dangerLevel * 150)
    
    ' 4. 安全な辺の戦略的価値
    If IsEdge(row, col) And dangerLevel = 0 Then
        Score = Score + GetEdgeValue(row, col)
    End If
    
    ' 5. 機動力（重要）
    Dim myMobility As Integer, oppMobility As Integer
    myMobility = CountValidMoves(Player)
    oppMobility = CountValidMoves(opponent)
    Score = Score + (myMobility - oppMobility) * 80
    
    ' 6. 確定石の評価
    Score = Score + CountStableStones(Player) * 50
    
    ' 7. 石数は中立的に評価
    Dim myStones As Integer, oppStones As Integer
    Call CountStones(myStones, oppStones, Player)
    Score = Score + (myStones - oppStones) * 10
    
    EvaluateMidgame = Score
End Function

' 終盤評価（石数重視）
Function EvaluateEndgame(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' 1. 角の価値
    If IsCorner(row, col) Then
        Score = Score + 3000
    End If
    
    ' 2. 相手に角を与える手への厳罰
    Dim cornerGift As Integer
    cornerGift = CheckCornerGift(row, col, Player)
    If cornerGift > 0 Then
        Score = Score - (cornerGift * 2000)
    End If
    
    ' 3. 石数が最重要
    Dim myStones As Integer, oppStones As Integer
    Call CountStones(myStones, oppStones, Player)
    Dim stoneDiff As Integer
    stoneDiff = myStones - oppStones
    Score = Score + stoneDiff * 150
    
    ' 4. 確定石の価値
    Dim myStable As Integer, oppStable As Integer
    myStable = CountStableStones(Player)
    oppStable = CountStableStones(opponent)
    Score = Score + (myStable - oppStable) * 100
    
    ' 5. 機動力
    Dim myMobility As Integer, oppMobility As Integer
    myMobility = CountValidMoves(Player)
    oppMobility = CountValidMoves(opponent)
    Score = Score + (myMobility - oppMobility) * 60
    
    ' 6. パリティ（残り手数の偶奇）
    Dim emptySquares As Integer
    emptySquares = BOARD_SIZE * BOARD_SIZE - CountTotalStones()
    If emptySquares <= 12 Then
        Dim parityBonus As Integer
        parityBonus = CalculateParityBonus(emptySquares, myMobility, oppMobility)
        Score = Score + parityBonus
    End If
    
    ' 7. ひっくり返す石数
    Dim flippedCount As Integer
    flippedCount = CountFlips(row, col, Player)
    Score = Score + flippedCount * 25
    
    EvaluateEndgame = Score
End Function

' 角かどうか判定
Function IsCorner(row As Integer, col As Integer) As Boolean
    IsCorner = (row = 1 Or row = BOARD_SIZE) And (col = 1 Or col = BOARD_SIZE)
End Function

' 辺かどうか判定
Function IsEdge(row As Integer, col As Integer) As Boolean
    IsEdge = (row = 1 Or row = BOARD_SIZE Or col = 1 Or col = BOARD_SIZE) And Not IsCorner(row, col)
End Function

' 中央領域かどうか判定
Function IsCenterRegion(row As Integer, col As Integer) As Boolean
    Dim center As Integer
    center = BOARD_SIZE / 2
    IsCenterRegion = (row >= center - 1 And row <= center + 2) And (col >= center - 1 And col <= center + 2)
End Function

' 辺の戦略的価値を計算
Function GetEdgeValue(row As Integer, col As Integer) As Integer
    Dim value As Integer
    value = 200  ' 基本的な辺の価値
    
    ' 角に近い辺ほど価値が高い
    If IsCornerAdjacent(row, col) Then
        value = value + 100
    End If
    
    GetEdgeValue = value
End Function

' 角に隣接する辺かどうか
Function IsCornerAdjacent(row As Integer, col As Integer) As Boolean
    ' 角の直線上にある辺
    IsCornerAdjacent = (row = 1 And (col = 1 Or col = BOARD_SIZE)) Or _
                      (row = BOARD_SIZE And (col = 1 Or col = BOARD_SIZE)) Or _
                      (col = 1 And (row = 1 Or row = BOARD_SIZE)) Or _
                      (col = BOARD_SIZE And (row = 1 Or row = BOARD_SIZE))
End Function

' 確定石（絶対にひっくり返されない石）をカウント
Function CountStableStones(Player As Integer) As Integer
    Dim stableCount As Integer
    Dim i As Integer, j As Integer
    
    ' 簡易版：角とその隣接する確定石のみカウント
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = Player Then
                If IsStableStone(i, j, Player) Then
                    stableCount = stableCount + 1
                End If
            End If
        Next j
    Next i
    
    CountStableStones = stableCount
End Function

' 指定位置の石が確定石かどうか判定
Function IsStableStone(row As Integer, col As Integer, Player As Integer) As Boolean
    ' 角は常に確定石
    If IsCorner(row, col) Then
        IsStableStone = True
        Exit Function
    End If
    
    ' 角から連続する辺の石で、間に空きがない場合
    If IsEdge(row, col) Then
        IsStableStone = IsStableEdge(row, col, Player)
    Else
        IsStableStone = False
    End If
End Function

' 辺の石が確定石かどうか判定
Function IsStableEdge(row As Integer, col As Integer, Player As Integer) As Boolean
    ' 簡易判定：角に隣接し、角が同じプレイヤーの石の場合
    If row = 1 Then  ' 上辺
        If col > 1 And gameBoard(1, 1) = Player Then
            IsStableEdge = True
        ElseIf col < BOARD_SIZE And gameBoard(1, BOARD_SIZE) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    ElseIf row = BOARD_SIZE Then  ' 下辺
        If col > 1 And gameBoard(BOARD_SIZE, 1) = Player Then
            IsStableEdge = True
        ElseIf col < BOARD_SIZE And gameBoard(BOARD_SIZE, BOARD_SIZE) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    ElseIf col = 1 Then  ' 左辺
        If row > 1 And gameBoard(1, 1) = Player Then
            IsStableEdge = True
        ElseIf row < BOARD_SIZE And gameBoard(BOARD_SIZE, 1) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    ElseIf col = BOARD_SIZE Then  ' 右辺
        If row > 1 And gameBoard(1, BOARD_SIZE) = Player Then
            IsStableEdge = True
        ElseIf row < BOARD_SIZE And gameBoard(BOARD_SIZE, BOARD_SIZE) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    Else
        IsStableEdge = False
    End If
End Function

' 相手に角を与える手かどうかをチェック
Function CheckCornerGift(row As Integer, col As Integer, Player As Integer) As Integer
    Dim giftCount As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' この手を打った後、相手が角を取れるようになるかチェック
    Dim corners As Variant
    corners = Array(Array(1, 1), Array(1, BOARD_SIZE), Array(BOARD_SIZE, 1), Array(BOARD_SIZE, BOARD_SIZE))
    
    Dim i As Integer
    For i = 0 To 3
        Dim cornerRow As Integer, cornerCol As Integer
        cornerRow = corners(i)(0)
        cornerCol = corners(i)(1)
        
        ' この角が空いていて、相手が取れるようになるかチェック
        If gameBoard(cornerRow, cornerCol) = CELL_EMPTY Then
            If IsValidMove(cornerRow, cornerCol, opponent) Then
                giftCount = giftCount + 1
            End If
        End If
    Next i
    
    CheckCornerGift = giftCount
End Function

' 角周辺の危険度レベルを詳細に判定
Function GetCornerDangerLevel(row As Integer, col As Integer) As Integer
    Dim dangerLevel As Integer
    
    ' 各角について危険度をチェック
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, 1, 1)                    ' 左上
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, 1, BOARD_SIZE)          ' 右上
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, BOARD_SIZE, 1)          ' 左下
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, BOARD_SIZE, BOARD_SIZE) ' 右下
    
    GetCornerDangerLevel = dangerLevel
End Function

' 特定の角に対する危険度をチェック
Function CheckSingleCornerDanger(row As Integer, col As Integer, cornerRow As Integer, cornerCol As Integer) As Integer
    ' その角が既に埋まっている場合は危険なし
    If gameBoard(cornerRow, cornerCol) <> CELL_EMPTY Then
        CheckSingleCornerDanger = 0
        Exit Function
    End If
    
    ' X-square（対角線隣接）: 最も危険
    If row = cornerRow + IIf(cornerRow = 1, 1, -1) And col = cornerCol + IIf(cornerCol = 1, 1, -1) Then
        CheckSingleCornerDanger = 5
    ' C-square（直線隣接）: 非常に危険
    ElseIf (row = cornerRow And Abs(col - cornerCol) = 1) Or (col = cornerCol And Abs(row - cornerRow) = 1) Then
        CheckSingleCornerDanger = 4
    ' A-square（角から2つ目の辺）: やや危険
    ElseIf (row = cornerRow And Abs(col - cornerCol) = 2) Or (col = cornerCol And Abs(row - cornerRow) = 2) Then
        CheckSingleCornerDanger = 2
    ' B-square（X-squareの隣）: 少し危険
    ElseIf Abs(row - cornerRow) = 2 And Abs(col - cornerCol) = 1 Then
        CheckSingleCornerDanger = 1
    ElseIf Abs(row - cornerRow) = 1 And Abs(col - cornerCol) = 2 Then
        CheckSingleCornerDanger = 1
    Else
        CheckSingleCornerDanger = 0
    End If
End Function

' パリティボーナスの詳細計算
Function CalculateParityBonus(emptySquares As Integer, myMobility As Integer, oppMobility As Integer) As Integer
    Dim bonus As Integer
    
    ' 基本的なパリティ
    If emptySquares Mod 2 = 0 Then
        bonus = 80  ' 偶数手残りは有利
    Else
        bonus = -40 ' 奇数手残りは不利
    End If
    
    ' 機動力との組み合わせ
    If myMobility > oppMobility Then
        bonus = bonus + 30  ' 選択肢が多い方が有利
    ElseIf myMobility < oppMobility Then
        bonus = bonus - 30
    End If
    
    ' 残り手数が少ない場合はより重要
    If emptySquares <= 6 Then
        bonus = bonus * 2
    ElseIf emptySquares <= 3 Then
        bonus = bonus * 3
    End If
    
    CalculateParityBonus = bonus
End Function

' 石数をカウント
Sub CountStones(ByRef myStones As Integer, ByRef oppStones As Integer, Player As Integer)
    Dim i As Integer, j As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    myStones = 0
    oppStones = 0
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = Player Then
                myStones = myStones + 1
            ElseIf gameBoard(i, j) = opponent Then
                oppStones = oppStones + 1
            End If
        Next j
    Next i
End Sub

' ゲーム段階名を取得
Function GetPhaseName(phase As Integer) As String
    Select Case phase
        Case PHASE_OPENING: GetPhaseName = "序盤"
        Case PHASE_MIDGAME: GetPhaseName = "中盤"
        Case PHASE_ENDGAME: GetPhaseName = "終盤"
        Case Else: GetPhaseName = "不明"
    End Select
End Function


' リセットボタンを作成
Sub CreateResetButton()
    Dim btn As Button
    Dim btnRange As Range
    
    ' 対象ワークシートが設定されているかチェック
    If targetSheet Is Nothing Then Exit Sub
    
    ' 全ボタンを削除してから作成
    Call DeleteAllButtons
    
    ' ボタンの配置場所を設定（棋譜エリアを避ける）
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 7, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 6, BOARD_SIZE))
    
    ' ボタンを作成
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!RestartGame"  ' 動的にワークブック名を指定
        .caption = "リセット"     ' ボタンの表示テキスト
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' ゲーム再開始（リセットボタン用）
Sub RestartGame()
    ' 既存のイベントハンドラをクリア
    Call CleanupEvents
    ' 評価表示モードをリセット
    showEvaluationMode = False
    Call InitializeGame
End Sub

' イベントハンドラをクリーンアップ
Sub CleanupEvents()
    On Error Resume Next
    Application.OnSheetSelectionChange = ""
    On Error GoTo 0
End Sub

' 手動入力ボタンを作成
Sub CreateManualInputButton()
    Dim btn As Button
    Dim btnRange As Range
    
    ' 対象ワークシートが設定されているかチェック
    If targetSheet Is Nothing Then Exit Sub
    
    ' ボタンの配置場所を設定
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 8, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE / 2))
    
    ' ボタンを作成
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ManualInput"  ' 動的にワークブック名を指定
        .caption = "手動入力"  ' ボタンの表示テキスト
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' 手動入力機能
Sub ManualInput()
    Dim userInput As String
    Dim col As Integer, row As Integer
    Dim colChar As String
    
    If gameOver Or CurrentPlayer <> BLACK Then
        MsgBox "現在は手動入力できません。", vbInformation
        Exit Sub
    End If
    
    userInput = InputBox("座標を入力してください (例: A4, B3):", "手動入力", "")
    If userInput = "" Then Exit Sub
    
    ' 入力を解析
    userInput = UCase(Trim(userInput))
    If Len(userInput) < 2 Then
        MsgBox "正しい形式で入力してください (例: A4, B3)", vbExclamation
        Exit Sub
    End If
    
    colChar = Left(userInput, 1)
    row = Val(Mid(userInput, 2))
    col = Asc(colChar) - 64  ' A=1, B=2, etc.
    
    ' 範囲チェック
    If col < 1 Or col > BOARD_SIZE Or row < 1 Or row > BOARD_SIZE Then
        MsgBox "座標が範囲外です。A1から" & Chr(64 + BOARD_SIZE) & BOARD_SIZE & "の範囲で入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 有効手チェックと実行
    If IsValidMove(row, col, BLACK) Then
        Dim playerEvaluation As Integer
        playerEvaluation = GetUnifiedEvaluation(row, col, BLACK)

        Call SaveGameState(row, col, BLACK, playerEvaluation)
        Call MakeMove(row, col, BLACK)
        Call UpdateDisplay
        
        ' 手の記録
        Call RecordMove(row, col, BLACK)
        Call UpdateGameRecordDisplay
        
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If
        
        ' CPUターン
        Call ExecuteCPUTurn
        
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If
        
        ' プレイヤーのターンに戻る
        Call SwitchToPlayerTurn
    Else
        MsgBox "そこには置けません。別の場所を選んでください。", vbExclamation
    End If
End Sub

' 評価表示切り替えボタンを作成
Sub CreateEvaluationToggleButton()
    Dim btn As Button
    Dim btnRange As Range
    
    ' 対象ワークシートが設定されているかチェック
    If targetSheet Is Nothing Then Exit Sub
    
    ' ボタンの配置場所を設定
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE / 2 + 1), _
                                    targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE))
    
    ' ボタンを作成
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ToggleEvaluationDisplay"  ' 動的にワークブック名を指定
        .caption = "評価表示: OFF"     ' ボタンの表示テキスト
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' 評価表示モード切り替え
Sub ToggleEvaluationDisplay()
    showEvaluationMode = Not showEvaluationMode
    
    ' ボタンのキャプションを更新
    Call UpdateEvaluationButtonCaption
    
    ' 画面表示を更新
    Call UpdateDisplay
End Sub
' 評価表示ボタンのキャプション更新
Sub UpdateEvaluationButtonCaption()
    Dim btn As Button
    Dim i As Integer
    
    On Error Resume Next
    ' 評価表示ボタンを探して更新
    For i = 1 To targetSheet.Buttons.count
        Set btn = targetSheet.Buttons(i)
        If InStr(btn.caption, "評価表示") > 0 Then
            If showEvaluationMode Then
                btn.caption = "評価表示: ON"
            Else
                btn.caption = "評価表示: OFF"
            End If
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' 全ボタンを削除
Sub DeleteAllButtons()
    Dim i As Integer
    
    On Error Resume Next  ' エラーが発生しても続行
    
    ' 対象ワークシートが設定されている場合のみ実行
    If Not targetSheet Is Nothing Then
        ' 逆順で全ボタンを削除
        For i = targetSheet.Buttons.count To 1 Step -1
            targetSheet.Buttons(i).Delete
        Next i
    End If
    
    On Error GoTo 0  ' エラーハンドリングを元に戻す
End Sub

' 画面表示更新
Sub UpdateDisplay()
    Dim i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            With targetSheet.Cells(i, j)
                Select Case gameBoard(i, j)
                    Case CELL_EMPTY
                        ' 空のセルの処理
                        .value = ""
                        .Font.color = RGB(0, 0, 0)
                        .Font.Size = 14

                        If showEvaluationMode And CurrentPlayer = BLACK And Not gameOver Then
                            If IsValidMove(i, j, BLACK) Then
                                Dim evalValue As Integer
                                evalValue = GetUnifiedEvaluation(i, j, BLACK)
                                .value = evalValue
                                .Font.Size = 8

                                .Font.color = GetEvaluationColor(evalValue)
                                .Font.Bold = True
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

' 評価値に応じた色分け
Function GetEvaluationColor(evalValue As Integer) As Long
    If evalValue >= 90 Then
        GetEvaluationColor = RGB(255, 0, 0)        ' 赤：最良手
    ElseIf evalValue >= 70 Then
        GetEvaluationColor = RGB(255, 100, 0)      ' 赤橙：非常に良い手
    ElseIf evalValue >= 50 Then
        GetEvaluationColor = RGB(255, 165, 0)      ' オレンジ：良い手
    ElseIf evalValue >= 30 Then
        GetEvaluationColor = RGB(255, 200, 0)      ' 黄橙：やや良い手
    ElseIf evalValue >= 10 Then
        GetEvaluationColor = RGB(255, 255, 0)      ' 黄：普通の手
    ElseIf evalValue >= -10 Then
        GetEvaluationColor = RGB(255, 255, 255)    ' 白：互角
    ElseIf evalValue >= -30 Then
        GetEvaluationColor = RGB(200, 200, 200)    ' 薄灰：やや悪い
    ElseIf evalValue >= -50 Then
        GetEvaluationColor = RGB(128, 128, 128)    ' 灰：悪い手
    ElseIf evalValue >= -70 Then
        GetEvaluationColor = RGB(100, 100, 100)    ' 濃灰：非常に悪い
    Else
        GetEvaluationColor = RGB(0, 0, 0)          ' 黒：危険な手
    End If
End Function

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
        "スコア - 黒: " & blackCount & "  白: " & whiteCount & "  ターン: " & moveCount
    
    Dim diffName As String
    Select Case difficulty
        Case EASY: diffName = "初級"
        Case MEDIUM: diffName = "中級"
        Case HARD: diffName = "上級"
    End Select
    
    Dim modeStatus As String
    If showEvaluationMode Then
        modeStatus = " | 評価表示: ON"
    Else
        modeStatus = " | 評価表示: OFF"
    End If
    
    targetSheet.Cells(BOARD_SIZE + 4, 1).value = "難易度: " & diffName & modeStatus
End Sub

' CPUターンを実行
Sub ExecuteCPUTurn()
    CurrentPlayer = WHITE
    
    ' CPUがスキップする必要があるかチェック
    If Not HasValidMoves(WHITE) Then
        Call ShowMessage("白に有効な手がありません。スキップします。")
        Call RecordMove(0, 0, WHITE)  ' スキップを記録
        Call UpdateGameRecordDisplay  ' 棋譜表示更新
        MsgBox "白に有効な手がありません。黒のターンです。", vbInformation, "ターンスキップ"
        Exit Sub
    End If
    
    ' CPUが手を実行
    Call ShowMessage("白のターンです。")
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    Call CPUTurn
    Call UpdateDisplay
    Call UpdateGameRecordDisplay  ' 棋譜表示更新
End Sub

' プレイヤーターンに切り替え
Sub SwitchToPlayerTurn()
    CurrentPlayer = BLACK
    
    ' プレイヤーがスキップする必要があるかチェック
    If Not HasValidMoves(BLACK) Then
        Call ShowMessage("黒に有効な手がありません。スキップします。")
        Call RecordMove(0, 0, BLACK)  ' スキップを記録
        Call UpdateGameRecordDisplay  ' 棋譜表示更新
        MsgBox "黒に有効な手がありません。白のターンです。", vbInformation, "ターンスキップ"

        ' 再度CPUターンを実行
        Call ExecuteCPUTurn
        Call UpdateDisplay
        Call UpdateGameRecordDisplay  ' 棋譜表示更新
        
        ' ゲーム終了チェック
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If
        
        ' もう一度プレイヤーターンをチェック
        Call SwitchToPlayerTurn
    Else
        ' プレイヤーに有効な手がある場合
        Call ShowMessage("黒色のターンです。")
        ' 評価表示モードの場合は画面を更新して評価値を表示
        If showEvaluationMode Then
            Call UpdateDisplay
        End If
    End If
End Sub

' セルクリック処理
Public Sub HandleCellClick(Target As Range)
    Call ProcessCellClick(Target)
End Sub

' 有効手リストを取得
Function GetValidMoves(Player As Integer) As ValidMove()
    Dim validMoves() As ValidMove
    Dim moveCount As Integer
    Dim i As Integer, j As Integer
    
    ReDim validMoves(1 To BOARD_SIZE * BOARD_SIZE)
    moveCount = 0
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If IsValidMove(i, j, Player) Then
                moveCount = moveCount + 1
                validMoves(moveCount).row = i
                validMoves(moveCount).col = j
                validMoves(moveCount).Score = GetUnifiedEvaluation(i, j, Player)
            End If
        Next j
    Next i
    
    ' 配列サイズを調整
    If moveCount > 0 Then
        ReDim Preserve validMoves(1 To moveCount)
    Else
        ReDim validMoves(1 To 1)  ' 空配列
    End If
    
    GetValidMoves = validMoves
End Function

' CPUの手（難易度に応じて処理を分岐）
Sub CPUTurn()
    Dim bestRow As Integer, bestCol As Integer
    Dim evaluation As Integer
    
    ' CPUに有効な手があるかチェック
    If Not HasValidMoves(WHITE) Then Exit Sub
    
    Select Case difficulty
        Case EASY
            Call CPUTurnEasy(bestRow, bestCol, evaluation)
        Case MEDIUM
            Call CPUTurnMedium(bestRow, bestCol, evaluation)
        Case HARD
            Call CPUTurnHard(bestRow, bestCol, evaluation)
    End Select
    
    ' 最良手を実行
    If bestRow > 0 And bestCol > 0 And IsValidMove(bestRow, bestCol, WHITE) Then
        Call SaveGameState(bestRow, bestCol, WHITE, evaluation)
        Call MakeMove(bestRow, bestCol, WHITE)
        Call RecordMove(bestRow, bestCol, WHITE)
    Else
        ' 無効な手が選ばれた場合は初級で代替
        Call CPUTurnEasy(bestRow, bestCol, evaluation)
        If bestRow > 0 And bestCol > 0 And IsValidMove(bestRow, bestCol, WHITE) Then
            Call SaveGameState(bestRow, bestCol, WHITE, evaluation)
            Call MakeMove(bestRow, bestCol, WHITE)
            Call RecordMove(bestRow, bestCol, WHITE)
        End If
    End If
End Sub

' 評価値を統一範囲に正規化する関数
Function NormalizeEvaluation(rawEval As Integer, difficulty As DifficultyLevel) As Integer
    Select Case difficulty
        Case EASY
            ' 初級はそのまま（1-100程度の範囲）
            NormalizeEvaluation = rawEval
        Case MEDIUM, HARD
            ' 中級・上級は-100～+100に正規化
            If rawEval >= 2000 Then
                NormalizeEvaluation = 100      ' 勝勢
            ElseIf rawEval >= 1000 Then
                NormalizeEvaluation = 80       ' 大きく有利
            ElseIf rawEval >= 500 Then
                NormalizeEvaluation = 50       ' 有利
            ElseIf rawEval >= 100 Then
                NormalizeEvaluation = 20       ' やや有利
            ElseIf rawEval > -100 Then
                NormalizeEvaluation = rawEval / 5  ' 微調整（-20～+20）
            ElseIf rawEval > -500 Then
                NormalizeEvaluation = -20      ' やや不利
            ElseIf rawEval > -1000 Then
                NormalizeEvaluation = -50      ' 不利
            ElseIf rawEval > -2000 Then
                NormalizeEvaluation = -80      ' 大きく不利
            Else
                NormalizeEvaluation = -100     ' 劣勢
            End If
    End Select
End Function

' 初級（評価関数のみ）
Sub CPUTurnEasy(ByRef bestRow As Integer, ByRef bestCol As Integer, ByRef bestEval As Integer)
    Dim validMoves() As ValidMove
    Dim bestScore As Integer
    Dim i As Integer
    
    validMoves = GetValidMoves(WHITE)
    bestScore = -1000
    bestRow = 0
    bestCol = 0
    
    ' 有効手が存在する場合のみ処理
    If UBound(validMoves) > 0 And validMoves(1).row > 0 Then
        For i = 1 To UBound(validMoves)
            If validMoves(i).Score > bestScore Then
                bestScore = validMoves(i).Score
                bestRow = validMoves(i).row
                bestCol = validMoves(i).col
            End If
        Next i
    End If
    
    ' 正規化して返す
    bestEval = NormalizeEvaluation(bestScore, EASY)
End Sub

' 中級（α-β探索 深度3）
Sub CPUTurnMedium(ByRef bestRow As Integer, ByRef bestCol As Integer, ByRef bestEval As Integer)
    Dim validMoves() As ValidMove
    Dim bestScore As Integer, Score As Integer
    Dim i As Integer
    Dim tempBoard(1 To 20, 1 To 20) As Integer
    
    validMoves = GetValidMoves(WHITE)
    bestScore = -9999
    bestRow = 0
    bestCol = 0
    
    Call CopyBoard(gameBoard, tempBoard)
    
    ' 有効手が存在する場合のみ処理
    If UBound(validMoves) > 0 And validMoves(1).row > 0 Then
        For i = 1 To UBound(validMoves)
            ' ボードを復元してから手を試す
            Call CopyBoard(tempBoard, gameBoard)
            Call MakeMove(validMoves(i).row, validMoves(i).col, WHITE)
            
            ' α-β探索（深度3）
            Score = AlphaBeta(3, -9999, 9999, False)
            
            If Score > bestScore Then
                bestScore = Score
                bestRow = validMoves(i).row
                bestCol = validMoves(i).col
            End If
            
            ' ボードを元に戻す
            Call CopyBoard(tempBoard, gameBoard)
        Next i
    End If
    
    ' 最終的にボード状態を復元
    Call CopyBoard(tempBoard, gameBoard)
    ' 正規化して返す
    bestEval = NormalizeEvaluation(bestScore, MEDIUM)
End Sub

' 上級（α-β探索 深度5）
Sub CPUTurnHard(ByRef bestRow As Integer, ByRef bestCol As Integer, ByRef bestEval As Integer)
    Dim validMoves() As ValidMove
    Dim bestScore As Integer, Score As Integer
    Dim i As Integer
    Dim tempBoard(1 To 20, 1 To 20) As Integer
    
    validMoves = GetValidMoves(WHITE)
    bestScore = -9999
    bestRow = 0
    bestCol = 0

    Call CopyBoard(gameBoard, tempBoard)
    
    ' 有効手が存在する場合のみ処理
    If UBound(validMoves) > 0 And validMoves(1).row > 0 Then
        For i = 1 To UBound(validMoves)
            ' ボードを復元してから手を試す
            Call CopyBoard(tempBoard, gameBoard)
            Call MakeMove(validMoves(i).row, validMoves(i).col, WHITE)
            
            ' α-β探索（深度5）
            Score = AlphaBeta(5, -9999, 9999, False)
            
            If Score > bestScore Then
                bestScore = Score
                bestRow = validMoves(i).row
                bestCol = validMoves(i).col
            End If
            
            ' ボードを元に戻す
            Call CopyBoard(tempBoard, gameBoard)
        Next i
    End If
    
    ' 最終的にボード状態を復元
    Call CopyBoard(tempBoard, gameBoard)
    ' 正規化して返す
    bestEval = NormalizeEvaluation(bestScore, HARD)
End Sub

' α-β探索
Function AlphaBeta(depth As Integer, alpha As Integer, beta As Integer, isMaximizing As Boolean) As Integer
    If depth = 0 Or CheckGameEnd() Then
        AlphaBeta = EvaluateBoardPosition()
        Exit Function
    End If
    
    Dim validMoves() As ValidMove
    Dim bestScore As Integer, Score As Integer
    Dim i As Integer
    Dim tempBoard(1 To 20, 1 To 20) As Integer
    Dim Player As Integer
    
    ' 現在のボード状態をバックアップ
    Call CopyBoard(gameBoard, tempBoard)
    
    If isMaximizing Then
        bestScore = -9999
        Player = WHITE
    Else
        bestScore = 9999
        Player = BLACK
    End If
    
    ' 有効手リストを取得
    validMoves = GetValidMoves(Player)
    
    ' 有効手が存在しない場合
    If UBound(validMoves) = 0 Or validMoves(1).row = 0 Then
        ' パスして相手のターン
        AlphaBeta = AlphaBeta(depth - 1, alpha, beta, Not isMaximizing)
        Call CopyBoard(tempBoard, gameBoard)
        Exit Function
    End If
    
    For i = 1 To UBound(validMoves)
        ' ボードを復元してから手を試す
        Call CopyBoard(tempBoard, gameBoard)
        Call MakeMove(validMoves(i).row, validMoves(i).col, Player)
        
        Score = AlphaBeta(depth - 1, alpha, beta, Not isMaximizing)
        
        If isMaximizing Then
            If Score > bestScore Then bestScore = Score
            If Score > alpha Then alpha = Score
            If beta <= alpha Then
                ' ボードを元に戻してからプルーニング
                Call CopyBoard(tempBoard, gameBoard)
                Exit For  ' β-カット
            End If
        Else
            If Score < bestScore Then bestScore = Score
            If Score < beta Then beta = Score
            If beta <= alpha Then
                ' ボードを元に戻してからプルーニング
                Call CopyBoard(tempBoard, gameBoard)
                Exit For  ' α-カット
            End If
        End If
        
        ' ボードを元に戻す
        Call CopyBoard(tempBoard, gameBoard)
    Next i
    
    ' 最終的にボード状態を復元
    Call CopyBoard(tempBoard, gameBoard)
    AlphaBeta = bestScore
End Function

' ボード全体の評価
Function EvaluateBoardPosition() As Integer
    Dim Score As Integer, i As Integer, j As Integer
    Dim whiteCount As Integer, blackCount As Integer
    Dim whiteMobility As Integer, blackMobility As Integer
    
    ' 石の数と位置価値を評価
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = WHITE Then
                whiteCount = whiteCount + 1
                Score = Score + GetPositionValue(i, j)
            ElseIf gameBoard(i, j) = BLACK Then
                blackCount = blackCount + 1
                Score = Score - GetPositionValue(i, j)
            End If
        Next j
    Next i
    
    ' 機動力（有効手の数）を評価
    whiteMobility = CountValidMoves(WHITE)
    blackMobility = CountValidMoves(BLACK)
    Score = Score + (whiteMobility - blackMobility) * 10
    
    ' ゲーム終盤では石数を重視
    If whiteCount + blackCount > BOARD_SIZE * BOARD_SIZE * 0.8 Then
        Score = Score + (whiteCount - blackCount) * 50
    End If
    
    EvaluateBoardPosition = Score
End Function

' 位置の価値を取得
Function GetPositionValue(row As Integer, col As Integer) As Integer
    ' 角
    If (row = 1 Or row = BOARD_SIZE) And (col = 1 Or col = BOARD_SIZE) Then
        GetPositionValue = 100
    ' 角の隣（危険地帯）
    ElseIf ((row = 1 Or row = BOARD_SIZE) And (col = 2 Or col = BOARD_SIZE - 1)) Or _
           ((col = 1 Or col = BOARD_SIZE) And (row = 2 Or row = BOARD_SIZE - 1)) Then
        GetPositionValue = -20
    ' 辺
    ElseIf row = 1 Or row = BOARD_SIZE Or col = 1 Or col = BOARD_SIZE Then
        GetPositionValue = 10
    ' 中央付近
    Else
        GetPositionValue = 1
    End If
End Function

' ボードコピー
Sub CopyBoard(sourceBoard() As Integer, targetBoard() As Integer)
    Dim i As Integer, j As Integer
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            targetBoard(i, j) = sourceBoard(i, j)
        Next j
    Next i
End Sub

' ゲーム状態保存
Sub SaveGameState(row As Integer, col As Integer, Player As Integer, evaluation As Integer)
    historyCount = historyCount + 1
    
    ' 配列サイズを拡張
    If historyCount > UBound(gameHistory) Then
        ReDim Preserve gameHistory(0 To historyCount + 50)
    End If
    
    With gameHistory(historyCount)
        ' ボード状態をコピー
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
        .evaluation = evaluation
    End With
End Sub

' 手の記録
Sub RecordMove(row As Integer, col As Integer, Player As Integer)
    moveCount = moveCount + 1
    
    If moveCount > UBound(moveHistory) Then
        ReDim Preserve moveHistory(1 To moveCount + 50)
    End If
    
    Dim playerName As String
    If Player = BLACK Then playerName = "黒" Else playerName = "白"
    
    Dim moveDescription As String
    If row = 0 And col = 0 Then
        ' スキップの場合
        moveDescription = "スキップ"
    Else
        Dim colLetter As String
        colLetter = Chr(64 + col)  ' 1->A, 2->B, etc.
        moveDescription = colLetter & row
    End If
    
    moveHistory(moveCount) = Format(moveCount, "000") & ": " & playerName & " " & _
                            moveDescription & " (" & Format(Now, "hh:mm:ss") & ")"
End Sub

' 棋譜保存
Sub SaveGameRecord()
    Dim recordSheet As Worksheet
    Dim lastRow As Integer, i As Integer
    Dim gameStartTime As String
    
    ' ゲーム開始時刻をゲームIDとして使用
    gameStartTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    ' 棋譜シートを作成または取得
    On Error Resume Next
    Set recordSheet = targetWorkbook.Worksheets("棋譜")
    On Error GoTo 0
    
    If recordSheet Is Nothing Then
        Set recordSheet = targetWorkbook.Worksheets.Add
        recordSheet.Name = "棋譜"
        
        ' ヘッダー設定
        recordSheet.Cells(1, 1).value = "ゲームID"
        recordSheet.Cells(1, 2).value = "ターン"
        recordSheet.Cells(1, 3).value = "プレイヤー"
        recordSheet.Cells(1, 4).value = "座標"
        recordSheet.Cells(1, 5).value = "時刻"
        recordSheet.Cells(1, 6).value = "評価値"
        recordSheet.Cells(1, 7).value = "難易度"
        recordSheet.Cells(1, 8).value = "ボードサイズ"
        recordSheet.Cells(1, 9).value = "スコア"
        recordSheet.Cells(1, 10).value = "勝者"
        
        ' ヘッダー書式設定
        With recordSheet.Range("A1:J1")
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' 列幅調整
        recordSheet.Columns("A").ColumnWidth = 18  ' ゲームID
        recordSheet.Columns("B").ColumnWidth = 6   ' ターン
        recordSheet.Columns("C").ColumnWidth = 8   ' プレイヤー
        recordSheet.Columns("D").ColumnWidth = 8   ' 座標
        recordSheet.Columns("E").ColumnWidth = 10  ' 時刻
        recordSheet.Columns("F").ColumnWidth = 8   ' 評価値
        recordSheet.Columns("G").ColumnWidth = 8   ' 難易度
        recordSheet.Columns("H").ColumnWidth = 10  ' ボードサイズ
        recordSheet.Columns("I").ColumnWidth = 15  ' スコア
        recordSheet.Columns("J").ColumnWidth = 12  ' 勝者
    End If
    
    ' 最終結果を取得
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
        winner = "黒"
    ElseIf finalWhiteCount > finalBlackCount Then
        winner = "白"
    Else
        winner = "引き分け"
    End If
    
    Dim diffName As String
    Select Case difficulty
        Case EASY: diffName = "初級"
        Case MEDIUM: diffName = "中級"
        Case HARD: diffName = "上級"
    End Select
    
    ' 各手を個別の行に記録
    lastRow = recordSheet.Cells(recordSheet.Rows.count, 1).End(xlUp).row
    
    For i = 1 To moveCount
        lastRow = lastRow + 1
        
        ' 手の情報を解析
        Dim moveInfo As String, playerName As String, movePos As String
        Dim timeInfo As String, moveNum As Integer
        Dim evalValue As Integer
        Dim currentScore As String

        moveInfo = moveHistory(i)
        
        ' ターンを抽出
        moveNum = Val(Left(moveInfo, 3))
        
        ' プレイヤー名を抽出
        If InStr(moveInfo, "黒") > 0 Then
            playerName = "黒"
        Else
            playerName = "白"
        End If
        
        ' 座標を抽出
        If InStr(moveInfo, "スキップ") > 0 Then
            movePos = "スキップ"
            evalValue = 0
        Else
            ' 座標部分を抽出（例：A4, B3など）
            Dim startPos As Integer, endPos As Integer
            startPos = InStr(moveInfo, playerName) + Len(playerName) + 1
            endPos = InStr(startPos, moveInfo, " (") - 1
            movePos = Trim(Mid(moveInfo, startPos, endPos - startPos + 1))
            
            ' 評価値を取得（履歴から）
            If i <= historyCount Then
                evalValue = gameHistory(i).evaluation
            Else
                evalValue = 0
            End If
        End If
        
        ' 時刻を抽出
        startPos = InStr(moveInfo, "(") + 1
        endPos = InStr(moveInfo, ")") - 1
        timeInfo = Mid(moveInfo, startPos, endPos - startPos + 1)
        
                
        ' その時点でのスコアを計算
        If i <= historyCount Then
            currentScore = CalculateScoreAtMove(i)
        Else
            currentScore = "0-0"  ' エラー時のデフォルト値
        End If
        
        ' データを記録
        With recordSheet
            .Cells(lastRow, 1).value = gameStartTime           ' ゲームID
            .Cells(lastRow, 2).value = moveNum                 ' ターン
            .Cells(lastRow, 3).value = playerName              ' プレイヤー
            .Cells(lastRow, 4).value = movePos                 ' 座標
            .Cells(lastRow, 5).value = timeInfo                ' 時刻
            .Cells(lastRow, 6).value = evalValue               ' 評価値
            .Cells(lastRow, 7).value = diffName                ' 難易度
            .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE  ' ボードサイズ
            .Cells(lastRow, 9).value = currentScore            ' その時点でのスコア
            .Cells(lastRow, 10).value = IIf(i = moveCount, winner, "-")  ' 最終手のみ勝者を表示
        End With
    Next i
    
    ' ゲーム終了行を追加
    lastRow = lastRow + 1
    With recordSheet
        .Cells(lastRow, 1).value = gameStartTime
        .Cells(lastRow, 2).value = "---"
        .Cells(lastRow, 3).value = "ゲーム終了"
        .Cells(lastRow, 4).value = "---"
        .Cells(lastRow, 5).value = Format(Now, "hh:mm:ss")
        .Cells(lastRow, 6).value = ""
        .Cells(lastRow, 7).value = diffName
        .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE
        .Cells(lastRow, 9).value = "黒" & finalBlackCount & "-" & finalWhiteCount & "白"  ' 最終スコア
        .Cells(lastRow, 10).value = winner
        
        ' 区切り行の書式設定
        .Range(.Cells(lastRow, 1), .Cells(lastRow, 10)).Interior.color = RGB(240, 240, 240)
        .Range(.Cells(lastRow, 1), .Cells(lastRow, 10)).Font.Bold = True
    End With
    
    ' 空行を追加（次のゲームとの区切り）
    lastRow = lastRow + 1
    
    MsgBox "棋譜を保存しました。" & vbCrLf & _
           "シート「棋譜」で確認できます。" & vbCrLf, vbInformation, "棋譜保存完了"
End Sub

' 指定した手番時点でのスコアを計算する関数
Function CalculateScoreAtMove(moveIndex As Integer) As String
    Dim blackCount As Integer, whiteCount As Integer
    Dim i As Integer, j As Integer
    
    ' 履歴が存在しない場合はエラー処理
    If moveIndex < 1 Or moveIndex > historyCount Then
        CalculateScoreAtMove = "エラー"
        Exit Function
    End If
    
    ' 指定した手番後のボード状態から石の数を数える
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
    
    ' スコア文字列を作成
    CalculateScoreAtMove = "黒" & blackCount & "-" & whiteCount & "白"
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

' 有効手の数をカウント
Function CountValidMoves(Player As Integer) As Integer
    Dim count As Integer, i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If IsValidMove(i, j, Player) Then count = count + 1
        Next j
    Next i
    
    CountValidMoves = count
End Function

' スキップ処理を管理
Sub CheckAndHandleSkip()
    Dim playerHasMoves As Boolean, cpuHasMoves As Boolean
    Dim skipMessage As String
    Dim skipCount As Integer
    
    ' 最大2回までの連続スキップを許可
    skipCount = 0
    
    Do While skipCount < 2
        playerHasMoves = HasValidMoves(BLACK)
        cpuHasMoves = HasValidMoves(WHITE)
        
        ' 両方に手がない場合はゲーム終了
        If Not playerHasMoves And Not cpuHasMoves Then
            gameOver = True
            Exit Sub
        End If
        
        ' 現在のプレイヤーに手がない場合
        If CurrentPlayer = BLACK And Not playerHasMoves Then
            skipMessage = "黒に有効な手がありません。白のターンです。"
            Call ShowMessage(skipMessage)
            Call RecordMove(0, 0, BLACK)  ' スキップを記録
            Call UpdateGameRecordDisplay  ' 棋譜表示更新
            MsgBox skipMessage, vbInformation, "ターンスキップ"
            
            CurrentPlayer = WHITE
            skipCount = skipCount + 1
            
            ' 白のターンをチェック
            If HasValidMoves(WHITE) Then
                Exit Sub  ' 白に手があるのでスキップ処理終了
            End If
            
        ElseIf CurrentPlayer = WHITE And Not cpuHasMoves Then
            skipMessage = "白に有効な手がありません。黒のターンです。"
            Call ShowMessage(skipMessage)
            Call RecordMove(0, 0, WHITE)  ' スキップを記録
            Call UpdateGameRecordDisplay  ' 棋譜表示更新
            MsgBox skipMessage, vbInformation, "ターンスキップ"
            
            CurrentPlayer = BLACK
            skipCount = skipCount + 1
            
            ' 黒のターンをチェック
            If HasValidMoves(BLACK) Then
                Exit Sub  ' 黒に手があるのでスキップ処理終了
            End If
            
        Else
            ' 現在のプレイヤーに手がある場合は終了
            Exit Sub
        End If
    Loop
    
    ' 2回連続でスキップが発生した場合はゲーム終了
    If skipCount >= 2 Then
        gameOver = True
    End If
End Sub

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

' 手の評価
Function EvaluateMove(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    
    ' 角の評価
    If (row = 1 Or row = BOARD_SIZE) And (col = 1 Or col = BOARD_SIZE) Then
        Score = 100
    ' 辺の評価
    ElseIf row = 1 Or row = BOARD_SIZE Or col = 1 Or col = BOARD_SIZE Then
        Score = 10
    Else
        Score = 1
    End If
    
    ' 取れる石の数を加算
    Score = Score + CountFlips(row, col, Player)
    
    ' 初級レベルとして正規化
    EvaluateMove = NormalizeEvaluation(Score, EASY)
End Function

' ひっくり返せる石の数をカウント
Function CountFlips(row As Integer, col As Integer, Player As Integer) As Integer
    Dim directions As Variant
    directions = Array(Array(-1, -1), Array(-1, 0), Array(-1, 1), _
                      Array(0, -1), Array(0, 1), _
                      Array(1, -1), Array(1, 0), Array(1, 1))
    
    Dim totalFlips As Integer, i As Integer
    Dim dr As Integer, dc As Integer, r As Integer, c As Integer
    Dim opponent As Integer, flipCount As Integer
    
    opponent = 3 - Player
    
    For i = 0 To 7
        dr = directions(i)(0)
        dc = directions(i)(1)
        r = row + dr
        c = col + dc
        flipCount = 0
        
        If r >= 1 And r <= BOARD_SIZE And c >= 1 And c <= BOARD_SIZE Then
            If gameBoard(r, c) = opponent Then
                Do
                    flipCount = flipCount + 1
                    r = r + dr
                    c = c + dc
                    If r < 1 Or r > BOARD_SIZE Or c < 1 Or c > BOARD_SIZE Then Exit Do
                    If gameBoard(r, c) = CELL_EMPTY Then Exit Do
                    If gameBoard(r, c) = Player Then
                        totalFlips = totalFlips + flipCount
                        Exit Do
                    End If
                Loop
            End If
        End If
    Next i
    
    CountFlips = totalFlips
End Function

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
    
    ' イベントハンドラをクリーンアップ
    Call CleanupEvents
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = BLACK Then blackCount = blackCount + 1
            If gameBoard(i, j) = WHITE Then whiteCount = whiteCount + 1
        Next j
    Next i
    
    If blackCount > whiteCount Then
        result = "黒の勝ちです！"
    ElseIf whiteCount > blackCount Then
        result = "白の勝ちです！"
    Else
        result = "引き分けです！"
    End If
    
    Call ShowMessage(result)
    targetSheet.Cells(BOARD_SIZE + 5, 1).value = _
        "最終スコア - 黒: " & blackCount & "  白: " & whiteCount
    
    ' 最終棋譜表示更新
    Call UpdateGameRecordDisplay
    
    ' 自動で棋譜保存
    Call SaveGameRecord
    
    MsgBox result & vbCrLf & vbCrLf & _
           "最終スコア" & vbCrLf & _
           "黒: " & blackCount & vbCrLf & _
           "白: " & whiteCount & vbCrLf & vbCrLf & _
           "棋譜を自動保存しました。" & vbCrLf & _
           "新しいゲームを始めるには リセットボタン をクリックしてください。", _
           vbInformation, "ゲーム終了"
End Sub





