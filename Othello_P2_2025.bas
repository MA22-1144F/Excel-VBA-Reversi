Attribute VB_Name = "Othello_P2_2025"
' ==========================================================
' VBA �I�Z���Q�[�� - 2�l�ΐ��
' ==========================================================

Option Explicit

' �Q�[���ݒ�
Public Const BOARD_SIZE As Integer = 8    ' �{�[�h�T�C�Y�i���������j
Public Const CELL_EMPTY As Integer = 0    ' ��̃Z��
Public Const BLACK As Integer = 1         ' �v���C���[1�i���j
Public Const WHITE As Integer = 2         ' �v���C���[2�i���j

' �v���C���[���
Public blackPlayerName As String          ' ���v���C���[��
Public whitePlayerName As String          ' ���v���C���[��

' �������Ԋ֘A
Public blackTotalTime As Double           ' ���̑��l�����ԁi�b�j
Public whiteTotalTime As Double           ' ���̑��l�����ԁi�b�j
Public currentTurnStartTime As Double     ' ���݂̃^�[���J�n����

' �q���g�@�\
Public showHints As Boolean

' �����\���ݒ�
Public RECORD_START_COL As Integer        ' �����\���J�n��
Public Const RECORD_WIDTH As Integer = 6  ' �����\�����i�񐔁j

' �Q�[����ԕۑ��p
Type GameState
    board(1 To 20, 1 To 20) As Integer
    Player As Integer
    moveRow As Integer
    moveCol As Integer
    movePlayer As Integer
    timestamp As String
    thinkingTime As Double
End Type

' �L����i�[�p
Type ValidMove
    row As Integer
    col As Integer
End Type

' �Q�[���ϐ�
Public gameBoard(1 To 20, 1 To 20) As Integer
Public CurrentPlayer As Integer
Public gameOver As Boolean
Public targetSheet As Worksheet
Public targetWorkbook As Workbook
Public gameHistory() As GameState
Public historyCount As Integer
Public moveHistory() As String
Public moveCount As Integer

' ���{������Ăяo�����C���֐�
Sub StartReversiGame2P()
    On Error GoTo ErrorHandler
    Call CreateNewGameWorkbook
    If targetWorkbook Is Nothing Or targetSheet Is Nothing Then
        MsgBox "���[�N�u�b�N�܂��̓��[�N�V�[�g�̍쐬�Ɏ��s���܂����B�ēx���s���Ă��������B", vbCritical
        Exit Sub
    End If
    
    Call InitializeGame2P
    Exit Sub
    
ErrorHandler:
    MsgBox "�Q�[���J�n���ɃG���[���������܂���: " & Err.description, vbCritical
End Sub

' �V�������[�N�u�b�N���쐬
Sub CreateNewGameWorkbook()
    Dim newWb As Workbook
    Dim ws As Worksheet
    Set newWb = Workbooks.Add
    Set ws = newWb.ActiveSheet
    If newWb Is Nothing Or ws Is Nothing Then
        MsgBox "�G���[: ���[�N�u�b�N�܂��̓��[�N�V�[�g�̍쐬�Ɏ��s���܂����B", vbCritical
        Exit Sub
    End If
    ws.Name = "�Q�[����"
    Set targetWorkbook = newWb
    Set targetSheet = ws
    Call SetupWorksheetEvents
End Sub

' ���[�N�V�[�g�C�x���g�̓��I�ݒ�
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

' VBA�v���W�F�N�g�A�N�Z�X�\���`�F�b�N
Function CheckVBAProjectAccess() As Boolean
    On Error GoTo AccessDenied
    Dim testName As String
    testName = targetWorkbook.VBProject.Name
    
    CheckVBAProjectAccess = True
    Exit Function
    
AccessDenied:
    CheckVBAProjectAccess = False
End Function

' �ȈՔ�
Sub SetupSimpleEvents()
    On Error Resume Next
    Application.OnSheetSelectionChange = ""
    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange2P"
    
    On Error GoTo 0
End Sub

' �A�v���P�[�V�������x����SelectionChange�C�x���g
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

' �Z���N���b�N�����̓����֐�
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
        Call ShowMessage("�����ɂ͒u���܂���B�ʂ̏ꏊ��I��ł��������B")
    End If
End Sub

' �v���C���[�̎�����s
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

' ���̃v���C���[�Ɍ��
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
            skipMessage = blackPlayerName & "�i���j�ɗL���Ȏ肪����܂���B" & vbCrLf & _
            GetCurrentPlayerName() & "�����s���܂��B"
        Else
            skipMessage = whitePlayerName & "�i���j�ɗL���Ȏ肪����܂���B" & vbCrLf & _
            GetCurrentPlayerName() & "�����s���܂��B"
        End If
        
        Call ShowMessage(skipMessage)
        Call RecordMove(0, 0, nextPlayer, 0)
        Call UpdateGameRecordDisplay
        MsgBox skipMessage, vbInformation, "�^�[���X�L�b�v"
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

' ���݂̃v���C���[�����擾
Function GetCurrentPlayerName() As String
    If CurrentPlayer = BLACK Then
        GetCurrentPlayerName = blackPlayerName & "�i���j"
    Else
        GetCurrentPlayerName = whitePlayerName & "�i���j"
    End If
End Function

' �^�[���\�����X�V
Sub UpdateTurnDisplay()
    Dim message As String
    message = GetCurrentPlayerName() & "�̃^�[���ł��B"
    Call ShowMessage(message)
End Sub

' �Q�[��������
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
    
    MsgBox "�I�Z��2�l�ΐ�Q�[�����J�n���܂����B" & vbCrLf & _
           "�{�[�h�T�C�Y: " & BOARD_SIZE & "x" & BOARD_SIZE & vbCrLf & _
           "��: " & blackPlayerName & vbCrLf & _
           "��: " & whitePlayerName & vbCrLf & vbCrLf & _
           "�Z���N���b�N�܂��͎蓮���͂Ńv���C���Ă��������B", vbInformation

    currentTurnStartTime = Timer
    
    Exit Sub
    
InitError:
    MsgBox "�Q�[�����������ɃG���[���������܂���: " & Err.description & vbCrLf & vbCrLf & _
           "�u�蓮���́v�{�^���Ńv���C�𑱍s�ł��܂��B", vbExclamation
    Resume Next
End Sub

' �v���C���[���̐ݒ�
Sub SetupPlayerNames()
    blackPlayerName = InputBox("���i���j�̃v���C���[������͂��Ă�������:", "�v���C���[���ݒ�", "�v���C���[1")
    If blackPlayerName = "" Then blackPlayerName = "�v���C���[1"
    
    whitePlayerName = InputBox("���i���j�̃v���C���[������͂��Ă�������:", "�v���C���[���ݒ�", "�v���C���[2")
    If whitePlayerName = "" Then whitePlayerName = "�v���C���[2"
End Sub

' �Q�[���@�\�̐ݒ�
Sub SetupGameFeatures()
    Dim response As String

    response = MsgBox("�L����̃q���g��\�����܂����H", vbYesNo + vbQuestion, "�@�\�ݒ�")
    showHints = (response = vbYes)
End Sub

' �Q�[���{�[�h������
Sub InitializeGameBoard()
    Dim i As Integer, j As Integer, center As Integer

    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            gameBoard(i, j) = CELL_EMPTY
        Next j
    Next i
    
    CurrentPlayer = BLACK
    gameOver = False
    
    ' �����Δz�u�i����4�}�X�j
    center = BOARD_SIZE / 2
    gameBoard(center, center) = WHITE
    gameBoard(center + 1, center + 1) = WHITE
    gameBoard(center, center + 1) = BLACK
    gameBoard(center + 1, center) = BLACK
End Sub

' ����������
Sub InitializeHistory()
    ReDim gameHistory(0 To 400)
    ReDim moveHistory(1 To 400)
    historyCount = 0
    moveCount = 0
End Sub

' UI�ݒ�
Sub SetupUI2P()
    Dim i As Integer, j As Integer

    If targetSheet Is Nothing Then
        MsgBox "�G���[: ���[�N�V�[�g�I�u�W�F�N�g���ݒ肳��Ă��܂���B", vbCritical
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

' �Q�[���p�{�^���쐬
Sub CreateGameButtons()
    Call DeleteAllButtons
    
    Call CreateResetButton2P        ' �Q�[�����Z�b�g�{�^����z�u
    Call CreateManualInputButton2P  ' �蓮���̓{�^����z�u
    Call CreateHintToggleButton     ' �q���g�\���{�^����z�u
    Call CreateHelpButton           ' �w���v�{�^����z�u
End Sub

' �����\���G���A�̐ݒ�
Sub SetupGameRecordArea()
    RECORD_START_COL = BOARD_SIZE + 2
    
    With targetSheet
        .Cells(1, RECORD_START_COL).value = "�^�[��"
        .Cells(1, RECORD_START_COL + 1).value = "�v���C���["
        .Cells(1, RECORD_START_COL + 2).value = "���W"
        .Cells(1, RECORD_START_COL + 3).value = "����"
        .Cells(1, RECORD_START_COL + 4).value = "�l������"
        .Cells(1, RECORD_START_COL + 5).value = "�݌v����"

        With .Range(.Cells(1, RECORD_START_COL), .Cells(1, RECORD_START_COL + 5))
            .Font.Bold = False
            .Font.Size = 10
            .Interior.color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns(RECORD_START_COL).ColumnWidth = 6      ' �^�[��
        .Columns(RECORD_START_COL + 1).ColumnWidth = 10 ' �v���C���[
        .Columns(RECORD_START_COL + 2).ColumnWidth = 6  ' ���W
        .Columns(RECORD_START_COL + 3).ColumnWidth = 8  ' ����
        .Columns(RECORD_START_COL + 4).ColumnWidth = 8  ' �l������
        .Columns(RECORD_START_COL + 5).ColumnWidth = 8  ' �݌v����
    End With
End Sub

' ���A���^�C�������\��������
Sub InitializeGameRecordDisplay()
    Dim i As Integer

    With targetSheet
        For i = 2 To BOARD_SIZE + 20
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 5)).ClearContents
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 5)).Interior.color = xlNone
        Next i
    End With
End Sub

' ���A���^�C�������\���X�V
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

        If InStr(moveInfo, "��") > 0 Then
            playerName = blackPlayerName & "�i���j"
        Else
            playerName = whitePlayerName & "�i���j"
        End If

        If InStr(moveInfo, "�X�L�b�v") > 0 Then
            movePos = "�X�L�b�v"
            thinkTime = "-"
            totalTime = "-"
        Else
            Dim startPos As Integer, endPos As Integer
            startPos = InStr(moveInfo, "�j") + 2
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

                If InStr(playerName, "��") > 0 Then
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

' ���Ԃ̃t�H�[�}�b�g
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

' �݌v���Ԃ��v�Z
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

' ���Z�b�g�{�^�����쐬
Sub CreateResetButton2P()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 9, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!RestartGame2P"
        .caption = "���Z�b�g"
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �Q�[���ĊJ�n
Sub RestartGame2P()
    Call CleanupEvents
    Call InitializeGame2P
End Sub

' �蓮���̓{�^�����쐬
Sub CreateManualInputButton2P()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 10, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 10, BOARD_SIZE))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ManualInput2P"
        .caption = "�蓮����"
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �蓮���͋@�\
Sub ManualInput2P()
    Dim userInput As String
    Dim col As Integer, row As Integer
    Dim colChar As String
    
    If gameOver Then
        MsgBox "�Q�[���͏I�����Ă��܂��B", vbInformation
        Exit Sub
    End If
    
    userInput = InputBox("���W����͂��Ă������� (��: A4, B3):" & vbCrLf & _
                        "���݂̃^�[��: " & GetCurrentPlayerName(), "�蓮����", "")
    If userInput = "" Then Exit Sub

    userInput = UCase(Trim(userInput))
    If Len(userInput) < 2 Then
        MsgBox "�������`���œ��͂��Ă������� (��: A4, B3)", vbExclamation
        Exit Sub
    End If
    
    colChar = Left(userInput, 1)
    row = Val(Mid(userInput, 2))
    col = Asc(colChar) - 64

    If col < 1 Or col > BOARD_SIZE Or row < 1 Or row > BOARD_SIZE Then
        MsgBox "���W���͈͊O�ł��BA1����" & Chr(64 + BOARD_SIZE) & BOARD_SIZE & "�͈̔͂œ��͂��Ă��������B", vbExclamation
        Exit Sub
    End If

    If IsValidMove(row, col, CurrentPlayer) Then
        Call ExecutePlayerMove(row, col)
    Else
        MsgBox "�����ɂ͒u���܂���B�ʂ̏ꏊ��I��ł��������B", vbExclamation
    End If
End Sub

' �q���g�\���؂�ւ��{�^�����쐬
Sub CreateHintToggleButton()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 11, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 11, BOARD_SIZE / 2))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ToggleHints"
        .caption = IIf(showHints, "�q���g: ON", "�q���g: OFF")
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �q���g�\�����[�h�؂�ւ�
Sub ToggleHints()
    showHints = Not showHints

    Call UpdateHintButtonCaption

    Call UpdateDisplay
End Sub

' �q���g�\���{�^���̃L���v�V�����X�V
Sub UpdateHintButtonCaption()
    Dim btn As Button
    Dim i As Integer
    
    On Error Resume Next
    For i = 1 To targetSheet.Buttons.count
        Set btn = targetSheet.Buttons(i)
        If InStr(btn.caption, "�q���g") > 0 Then
            If showHints Then
                btn.caption = "�q���g: ON"
            Else
                btn.caption = "�q���g: OFF"
            End If
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' �w���v�{�^�����쐬
Sub CreateHelpButton()
    Dim btn As Button
    Dim btnRange As Range
    
    If targetSheet Is Nothing Then Exit Sub
    
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 11, BOARD_SIZE / 2 + 1), _
                                    targetSheet.Cells(BOARD_SIZE + 11, BOARD_SIZE))
    
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ShowHelp"
        .caption = "�w���v"
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �w���v�\��
Sub ShowHelp()
    Dim helpText As String
    helpText = "�y�I�Z��2�l�ΐ� ������@�z" & vbCrLf & vbCrLf & _
               "�� �΂̒u����:" & vbCrLf & _
               "�E�Ֆʂ̋󂢂Ă���}�X���N���b�N" & vbCrLf & _
               "�E�u�蓮���́v�{�^���ō��W���́i��FA4�j" & vbCrLf & vbCrLf & _
               "�� �@�\�{�^��:" & vbCrLf & _
               "�E�V�����Q�[���F�Q�[�������Z�b�g" & vbCrLf & _
               "�E�q���g�F�L����̈ʒu��\��" & vbCrLf & vbCrLf & _
               "�� ����:" & vbCrLf & _
               "�E�E���ɍŋ߂̎肪�L�^����܂��B" & vbCrLf & _
               "�E�l�����ԂƗ݌v���Ԃ��L�^����܂��B" & vbCrLf & _
               "�E�Q�[���I�����Ɏ����ۑ�����܂��B" & vbCrLf & vbCrLf & _
               "�� ���[��:" & vbCrLf & _
               "�E����̐΂�����Ŏ����̐΂ɂ���" & vbCrLf & _
               "�E�u����ꏊ���Ȃ��ꍇ�̓X�L�b�v" & vbCrLf & _
               "�E�Ֆʂ����܂邩���҂Ƃ��u���Ȃ��Ȃ�����I��"
    
    MsgBox helpText, vbInformation, "�w���v"
End Sub

' �C�x���g�n���h�����N���[���A�b�v
Sub CleanupEvents()
    On Error Resume Next
    Application.OnSheetSelectionChange = ""
    On Error GoTo 0
End Sub

' �S�{�^�����폜
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

' ��ʕ\���X�V
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
                                .value = "��"
                                .Font.Size = 10
                                .Font.color = RGB(100, 100, 100)
                                .Font.Bold = False
                            End If
                        End If
                        
                    Case BLACK
                        .value = "��"
                        .Font.color = RGB(0, 0, 0)
                        .Font.Size = 14
                        .Font.Bold = True
                    Case WHITE
                        .value = "��"
                        .Font.color = RGB(255, 255, 255)
                        .Font.Size = 14
                        .Font.Bold = True
                End Select
            End With
        Next j
    Next i
    
    Call ShowScore
End Sub

' ���b�Z�[�W�\��
Sub ShowMessage(msg As String)
    targetSheet.Cells(BOARD_SIZE + 2, 1).value = msg
End Sub

' �X�R�A�\��
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
        "�X�R�A - " & blackPlayerName & "�i���j: " & blackCount & "  " & whitePlayerName & "�i���j: " & whiteCount
    
    targetSheet.Cells(BOARD_SIZE + 4, 1).value = "�^�[��: " & moveCount

    Dim statusText As String
    statusText = "�@�\: "
    If showHints Then statusText = statusText & "�q���gON"
    targetSheet.Cells(BOARD_SIZE + 5, 1).value = statusText
End Sub

' �Q�[����ԕۑ�
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

' ��̋L�^
Sub RecordMove(row As Integer, col As Integer, Player As Integer, thinkingTime As Double)
    moveCount = moveCount + 1
    
    If moveCount > UBound(moveHistory) Then
        ReDim Preserve moveHistory(1 To moveCount + 50)
    End If
    
    Dim playerName As String
    If Player = BLACK Then
        playerName = blackPlayerName & "�i���j"
    Else
        playerName = whitePlayerName & "�i���j"
    End If
    
    Dim moveDescription As String
    If row = 0 And col = 0 Then
        moveDescription = "�X�L�b�v"
    Else
        Dim colLetter As String
        colLetter = Chr(64 + col)
        moveDescription = colLetter & row
    End If
    
    moveHistory(moveCount) = Format(moveCount, "000") & ": " & playerName & " " & _
                            moveDescription & " (" & Format(Now, "hh:mm:ss") & ")"
End Sub

' �����ۑ�
Sub SaveGameRecord2P()
    Dim recordSheet As Worksheet
    Dim lastRow As Integer, i As Integer
    Dim gameStartTime As String
    gameStartTime = Format(Now, "yyyy/mm/dd hh:mm:ss")

    On Error Resume Next
    Set recordSheet = targetWorkbook.Worksheets("����")
    On Error GoTo 0
    
    If recordSheet Is Nothing Then
        Set recordSheet = targetWorkbook.Worksheets.Add
        recordSheet.Name = "����"

        recordSheet.Cells(1, 1).value = "�Q�[��ID"
        recordSheet.Cells(1, 2).value = "�^�[��"
        recordSheet.Cells(1, 3).value = "�v���C���["
        recordSheet.Cells(1, 4).value = "���W"
        recordSheet.Cells(1, 5).value = "����"
        recordSheet.Cells(1, 6).value = "�l������"
        recordSheet.Cells(1, 7).value = "�݌v����"
        recordSheet.Cells(1, 8).value = "�{�[�h�T�C�Y"
        recordSheet.Cells(1, 9).value = "�X�R�A"
        recordSheet.Cells(1, 10).value = "����"

        With recordSheet.Range("A1:J1")
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With

        recordSheet.Columns("A").ColumnWidth = 18  ' �Q�[��ID
        recordSheet.Columns("B").ColumnWidth = 6   ' �^�[��
        recordSheet.Columns("C").ColumnWidth = 12  ' �v���C���[
        recordSheet.Columns("D").ColumnWidth = 8   ' ���W
        recordSheet.Columns("E").ColumnWidth = 10  ' ����
        recordSheet.Columns("F").ColumnWidth = 10  ' �l������
        recordSheet.Columns("G").ColumnWidth = 10  ' �݌v����
        recordSheet.Columns("H").ColumnWidth = 10  ' �{�[�h�T�C�Y
        recordSheet.Columns("I").ColumnWidth = 15  ' �X�R�A
        recordSheet.Columns("J").ColumnWidth = 15  ' ����
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
        winner = blackPlayerName & "�i���j"
    ElseIf finalWhiteCount > finalBlackCount Then
        winner = whitePlayerName & "�i���j"
    Else
        winner = "��������"
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

        If InStr(moveInfo, "��") > 0 Then
            playerName = blackPlayerName & "�i���j"
        Else
            playerName = whitePlayerName & "�i���j"
        End If

        If InStr(moveInfo, "�X�L�b�v") > 0 Then
            movePos = "�X�L�b�v"
            thinkTime = "-"
            totalTime = "-"
        Else

            Dim startPos As Integer, endPos As Integer
            startPos = InStr(moveInfo, "�j") + 2
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
            .Cells(lastRow, 1).value = gameStartTime           ' �Q�[��ID
            .Cells(lastRow, 2).value = moveNum                 ' �^�[��
            .Cells(lastRow, 3).value = playerName              ' �v���C���[
            .Cells(lastRow, 4).value = movePos                 ' ���W
            .Cells(lastRow, 5).value = timeInfo                ' ����
            .Cells(lastRow, 6).value = thinkTime               ' �l������
            .Cells(lastRow, 7).value = totalTime               ' �݌v����
            .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE  ' �{�[�h�T�C�Y
            .Cells(lastRow, 9).value = currentScore            ' ���̎��_�ł̃X�R�A
            .Cells(lastRow, 10).value = IIf(i = moveCount, winner, "-")  ' �ŏI��̂ݏ��҂�\��
        End With
    Next i

    lastRow = lastRow + 1
    With recordSheet
        .Cells(lastRow, 1).value = gameStartTime
        .Cells(lastRow, 2).value = "---"
        .Cells(lastRow, 3).value = "�Q�[���I��"
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
    
    MsgBox "������ۑ����܂����B" & vbCrLf & _
           "�V�[�g�u�����v�Ŋm�F�ł��܂��B" & vbCrLf, vbInformation, "�����ۑ�����"
End Sub

' �w�肵����Ԏ��_�ł̃X�R�A���v�Z����֐�
Function CalculateScoreAtMove2P(moveIndex As Integer) As String
    Dim blackCount As Integer, whiteCount As Integer
    Dim i As Integer, j As Integer

    If moveIndex < 1 Or moveIndex > historyCount Then
        CalculateScoreAtMove2P = "�G���["
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

' �w��v���C���[�ɗL���Ȏ肪���邩�`�F�b�N
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

' �L���蔻��
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

' ������s
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

' �Q�[���I������
Function CheckGameEnd() As Boolean
    Dim hasBlackMove As Boolean, hasWhiteMove As Boolean
    
    hasBlackMove = HasValidMoves(BLACK)
    hasWhiteMove = HasValidMoves(WHITE)
    
    CheckGameEnd = Not (hasBlackMove Or hasWhiteMove)
    gameOver = CheckGameEnd
End Function

' ���ʕ\��
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
        result = blackPlayerName & "�i���j�̏����ł��I"
    ElseIf whiteCount > blackCount Then
        result = whitePlayerName & "�i���j�̏����ł��I"
    Else
        result = "���������ł��I"
    End If
    
    Call ShowMessage(result)
    targetSheet.Cells(BOARD_SIZE + 6, 1).value = _
        "�ŏI�X�R�A - " & blackPlayerName & "�i���j: " & blackCount & "  " & whitePlayerName & "�i���j: " & whiteCount

    Call UpdateGameRecordDisplay
    Call SaveGameRecord2P
    
    Dim finalMessage As String
    finalMessage = result & vbCrLf & vbCrLf & _
           "�ŏI�X�R�A" & vbCrLf & _
           blackPlayerName & "�i���j: " & blackCount & vbCrLf & _
           whitePlayerName & "�i���j: " & whiteCount & vbCrLf & vbCrLf & _
           "���l������" & vbCrLf & _
           blackPlayerName & ": " & FormatTime(blackTotalTime) & vbCrLf & _
           whitePlayerName & ": " & FormatTime(whiteTotalTime) & vbCrLf & vbCrLf & _
           "�����������ۑ����܂����B" & vbCrLf & _
           "�V�����Q�[�����n�߂�ɂ́u�V�����Q�[���v�{�^�����N���b�N���Ă��������B"
    
    MsgBox finalMessage, vbInformation, "�Q�[���I��"
End Sub

