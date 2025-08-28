Attribute VB_Name = "Othello_2025"
' ==========================================================
' VBA �I�Z���Q�[��
' ==========================================================

Option Explicit

' �Q�[���ݒ�
Public Const BOARD_SIZE As Integer = 8    ' �{�[�h�T�C�Y�i���������j
Public Const CELL_EMPTY As Integer = 0    ' ��̃Z��
Public Const BLACK As Integer = 1         ' �v���C���[�i���j
Public Const WHITE As Integer = 2         ' CPU�i���j
Public showEvaluationMode As Boolean  ' �]���\�����[�h�̃I��/�I�t

' �����\���ݒ�
Public RECORD_START_COL As Integer             ' �����\���J�n��
Public Const RECORD_WIDTH As Integer = 5       ' �����\�����i�񐔁j

' ��Փx�ݒ�
Public Enum DifficultyLevel
    EASY = 1        ' �ȒP�i�]���֐��̂݁j
    MEDIUM = 2      ' ���ʁi��-���T�� �[�x3�j
    HARD = 3        ' ����i��-���T�� �[�x5�j
End Enum

' �Q�[����ԕۑ��p
Type GameState
    board(1 To 20, 1 To 20) As Integer
    Player As Integer
    moveRow As Integer
    moveCol As Integer
    movePlayer As Integer
    timestamp As String
    evaluation As Integer
End Type

' �L����i�[�p
Type ValidMove
    row As Integer
    col As Integer
    Score As Integer
End Type

' �Q�[���i�K�̒�`
Public Const PHASE_OPENING As Integer = 1     ' ���Ձi�`20��j
Public Const PHASE_MIDGAME As Integer = 2     ' ���Ձi21�`45��j
Public Const PHASE_ENDGAME As Integer = 3     ' �I�Ձi46��`�j

' �Q�[���ϐ�
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

' ���{������Ăяo�����C���֐�
Sub StartReversiGame()
    On Error GoTo ErrorHandler
    
    ' �V�������[�N�u�b�N���쐬���ăQ�[�����J�n
    Call CreateNewGameWorkbook
    
    If targetWorkbook Is Nothing Or targetSheet Is Nothing Then
        MsgBox "���[�N�u�b�N�܂��̓��[�N�V�[�g�̍쐬�Ɏ��s���܂����B�ēx���s���Ă��������B", vbCritical
        Exit Sub
    End If
    
    Call InitializeGame
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
    
    ' ���[�N�V�[�g����ݒ�
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
               "        Application.Run """ & ThisWorkbook.Name & "!ProcessCellClick"", Target" & vbCrLf & _
               "    End If" & vbCrLf & _
               "End Sub"
    
    CodeModule.AddFromString eventCode

    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange"
    
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

' �ȈՔŁFApplication.OnSheetSelectionChange�݂̂��g�p
Sub SetupSimpleEvents()
    On Error Resume Next

    Application.OnSheetSelectionChange = ""
    Application.OnSheetSelectionChange = ThisWorkbook.Name & "!OnSheetSelectionChange"
    
    On Error GoTo 0
End Sub

' �A�v���P�[�V�������x����SelectionChange�C�x���g
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

' �Z���N���b�N�����̊֐�
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
        Call ShowMessage("�����ɂ͒u���܂���B�ʂ̏ꏊ��I��ł��������B")
    End If
End Sub

' ���[�N�V�[�g�C�x���g����Ăяo�����֐�
Public Sub HandleCellClickFromEvent(Target As Range)
    Call ProcessCellClick(Target)
End Sub

' �Q�[��������
Sub InitializeGame()
    Dim response As String
    Dim diffLevel As Integer
    
    On Error GoTo InitError
    
    ' �]���\�����[�h���������i�f�t�H���g�̓I�t�j
    showEvaluationMode = False
    
    ' ��Փx�I��
    response = InputBox("��Փx��I�����Ă�������:" & vbCrLf & _
                       "1: �����i�]���֐��̂݁j" & vbCrLf & _
                       "2: �����i��-���T�� �[�x3�j" & vbCrLf & _
                       "3: �㋉�i��-���T�� �[�x5�j", _
                       "��Փx�I��", "2")
    
    If response = "" Then
        ' �L�����Z�����ꂽ�ꍇ�̓��[�N�u�b�N�����
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
        Call ShowMessage("���̃^�[���ł��B�u�������ꏊ���N���b�N���Ă��������B")
    End If

    Call SetupWorksheetEvents
    
    Dim diffName As String
    Select Case difficulty
        Case EASY: diffName = "����"
        Case MEDIUM: diffName = "����"
        Case HARD: diffName = "�㋉"
    End Select
    
    MsgBox "�I�Z���Q�[�����J�n���܂����B" & vbCrLf & _
           "�{�[�h�T�C�Y: " & BOARD_SIZE & "x" & BOARD_SIZE & vbCrLf & _
           "��Փx: " & diffName & vbCrLf & _
           "���̃^�[���ł��B" & vbCrLf & vbCrLf & _
           "�Z���N���b�N�܂��͎蓮���͂Ńv���C���Ă��������B", vbInformation
    
    Exit Sub
    
InitError:
    MsgBox "�Q�[�����������ɃG���[���������܂���: " & Err.description & vbCrLf & vbCrLf & _
           "�u�蓮���́v�{�^���Ńv���C�𑱍s�ł��܂��B", vbExclamation
    Resume Next
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
    ReDim gameHistory(0 To 400)  ' �ő�400��i20x20�Ή��j
    ReDim moveHistory(1 To 400)
    historyCount = 0
    moveCount = 0
    
    ' ������Ԃ�ۑ�
    Call SaveGameState(0, 0, 0, 0)
End Sub

' UI�ݒ�
Sub SetupUI()
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
        
        ' �{�^���ނ�z�u
        Call CreateResetButton      ' �Q�[�����Z�b�g�{�^����z�u
        Call CreateManualInputButton        ' �蓮���̓{�^����z�u
        Call CreateEvaluationToggleButton       ' �]���l�\���{�^����z�u
        
        ' �����\���G���A�̐ݒ�
        Call SetupGameRecordArea
    End With
End Sub

' �����\���G���A�̐ݒ�
Sub SetupGameRecordArea()
    RECORD_START_COL = BOARD_SIZE + 2
    
    With targetSheet
        ' �����w�b�_�[�s�i1�s�ڂɔz�u�j
        .Cells(1, RECORD_START_COL).value = "�^�[��"
        .Cells(1, RECORD_START_COL + 1).value = "�v���C���["
        .Cells(1, RECORD_START_COL + 2).value = "���W"
        .Cells(1, RECORD_START_COL + 3).value = "����"
        .Cells(1, RECORD_START_COL + 4).value = "�]��"
        
        ' �w�b�_�[�s�̏����ݒ�
        With .Range(.Cells(1, RECORD_START_COL), .Cells(1, RECORD_START_COL + 4))
            .Font.Bold = False
            .Font.Size = 10
            .Interior.color = RGB(220, 220, 220)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        
        ' �񕝐ݒ�
        .Columns(RECORD_START_COL).ColumnWidth = 6      ' �^�[��
        .Columns(RECORD_START_COL + 1).ColumnWidth = 8  ' �v���C���[
        .Columns(RECORD_START_COL + 2).ColumnWidth = 6  ' ���W
        .Columns(RECORD_START_COL + 3).ColumnWidth = 8  ' ����
        .Columns(RECORD_START_COL + 4).ColumnWidth = 6  ' �]��
    End With
End Sub

' ���A���^�C�������\��������
Sub InitializeGameRecordDisplay()
    Dim i As Integer
    
    ' �����̊����f�[�^���N���A�i�w�b�_�[�͕ێ��j
    With targetSheet
        For i = 2 To BOARD_SIZE + 20
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 4)).ClearContents
            .Range(.Cells(i, RECORD_START_COL), .Cells(i, RECORD_START_COL + 4)).Interior.color = xlNone
        Next i
    End With
End Sub

' ���A���^�C�������\���X�V
Sub UpdateGameRecordDisplay()
    Dim i As Integer, displayRow As Integer
    Dim moveInfo As String, playerName As String, movePos As String
    Dim timeInfo As String, moveNum As Integer, evalValue As Integer
    Dim startPos As Integer, endPos As Integer
    Dim maxDisplayRows As Integer
    
    ' �����\�����N���A
    Call InitializeGameRecordDisplay
    
    ' �\���\�s�����v�Z�i�{�[�h�T�C�Y�ɉ����āj
    maxDisplayRows = BOARD_SIZE + 15
    
    ' �e����t���ŕ\���i�ŐV�肪��A�ŏ��̎肪���j
    For i = moveCount To 1 Step -1
        displayRow = (moveCount - i) + 2  ' �w�b�_�[�����I�t�Z�b�g
        
        If displayRow > maxDisplayRows Then Exit For  ' �\���͈͂𒴂����ꍇ
        
        moveInfo = moveHistory(i)
        moveNum = Val(Left(moveInfo, 3))
        If InStr(moveInfo, "��") > 0 Then
            playerName = "��"
        Else
            playerName = "��"
        End If
        If InStr(moveInfo, "�X�L�b�v") > 0 Then
            movePos = "�X�L�b�v"
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
            .Cells(displayRow, RECORD_START_COL + 4).value = IIf(evalValue = 0 And movePos = "�X�L�b�v", "-", evalValue)

            With .Range(.Cells(displayRow, RECORD_START_COL), .Cells(displayRow, RECORD_START_COL + 4))
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin

                If playerName = "��" Then
                    .Interior.color = RGB(240, 240, 240)
                Else
                    .Interior.color = RGB(255, 255, 255)
                End If
            End With
        End With
    Next i
    
    ' �ŐV����n�C���C�g
    If moveCount > 0 Then
        displayRow = 2  ' �ŐV��͏��2�s��
        With targetSheet.Range(targetSheet.Cells(displayRow, RECORD_START_COL), _
                              targetSheet.Cells(displayRow, RECORD_START_COL + 4))
            .Interior.color = RGB(255, 255, 150)  ' ���F�n�C���C�g
            .Font.Bold = True
        End With
    End If
End Sub

' ���݂̃Q�[���i�K�𔻒�
Function GetGamePhase() As Integer
    Dim totalMoves As Integer
    totalMoves = CountTotalStones() - 4  ' ����4�΂�����
    
    If totalMoves <= 20 Then
        GetGamePhase = PHASE_OPENING
    ElseIf totalMoves <= 45 Then
        GetGamePhase = PHASE_MIDGAME
    Else
        GetGamePhase = PHASE_ENDGAME
    End If
End Function

' �Տ�̐΂̑������J�E���g
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

' �]���֐�
Function GetUnifiedEvaluation(row As Integer, col As Integer, Player As Integer) As Integer
    Dim evaluation As Integer
    Dim tempBoard(1 To 20, 1 To 20) As Integer
    Dim phase As Integer
    
    ' �L����łȂ��ꍇ��0��Ԃ�
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

' ���K������
Function NormalizeEvaluationSmooth(rawEval As Integer, phase As Integer) As Integer
    Dim normalizedValue As Single
    
    Select Case phase
        Case PHASE_OPENING
            ' ���ՁF�@���͏d���A�p�̉��l������
            normalizedValue = SigmoidNormalization(rawEval, 800, 1.2)
            
        Case PHASE_MIDGAME
            ' ���ՁF�o�����X�d��
            normalizedValue = SigmoidNormalization(rawEval, 1000, 1)
            
        Case PHASE_ENDGAME
            ' �I�ՁF�ׂ��������d���A���q���Ȕ���
            normalizedValue = TanhNormalization(rawEval, 600, 0.8)
    End Select
    
    ' -100�`+100�͈̔͂ɐ���
    If normalizedValue > 100 Then
        NormalizeEvaluationSmooth = 100
    ElseIf normalizedValue < -100 Then
        NormalizeEvaluationSmooth = -100
    Else
        NormalizeEvaluationSmooth = Round(normalizedValue, 0)
    End If
End Function

' �V�O���C�h�֐��ɂ�鐳�K��
Function SigmoidNormalization(value As Integer, scaleParam As Single, sensitivity As Single) As Single

    Dim x As Single
    Dim result As Single
    
    ' �[�����Z��h��
    If scaleParam = 0 Then scaleParam = 1
    
    x = value / scaleParam
    
    ' Exp�֐��̈������傫������ꍇ�̑΍�
    If x > 50 Then
        result = 1
    ElseIf x < -50 Then
        result = -1
    Else
        result = 2 / (1 + Exp(-x)) - 1
    End If
    
    ' 100�{���� -100�`+100 �͈̔͂ɂ��A���x�𒲐�
    SigmoidNormalization = result * 100 * sensitivity
End Function

' �n�C�p�[�{���b�N�^���W�F���g�֐��ɂ�鐳�K��
Function TanhNormalization(value As Integer, scaleParam As Single, sensitivity As Single) As Single
    Dim x As Single
    Dim result As Single
    Dim expX As Single, expNegX As Single
    
    ' �[�����Z��h��
    If scaleParam = 0 Then scaleParam = 1
    
    x = value / scaleParam
    
    ' Exp�֐��̈������傫������ꍇ�̑΍�
    If x > 50 Then
        result = 1
    ElseIf x < -50 Then
        result = -1
    Else
        expX = Exp(x)
        expNegX = Exp(-x)
        result = (expX - expNegX) / (expX + expNegX)
    End If
    
    ' 100�{���� -100�`+100 �͈̔͂ɂ��A���x�𒲐�
    TanhNormalization = result * 100 * sensitivity
End Function

' ��荂�x�Ȑ��K���i�����̊֐���g�ݍ��킹�j
Function AdvancedNormalization(value As Integer, phase As Integer) As Integer
    Dim result As Double
    Dim absValue As Double
    absValue = Abs(value)
    
    Select Case phase
        Case PHASE_OPENING
            ' ���ՁF�p�̉��l�������������`�֐�
            If absValue <= 100 Then
                ' �������l�͐��`
                result = value * 0.3
            ElseIf absValue <= 1000 Then
                ' �����x�̒l�͑ΐ��I���k
                result = Sgn(value) * (30 + 40 * Log(absValue / 100) / Log(10))
            Else
                ' �傫���l�i�p�Ȃǁj�͎w���I�ɋ���
                result = Sgn(value) * (70 + 30 * (1 - Exp(-(absValue - 1000) / 2000)))
            End If
            
        Case PHASE_MIDGAME
            ' ���ՁF�o�����X�̎�ꂽ�]��
            If absValue <= 200 Then
                result = value * 0.2
            Else
                result = Sgn(value) * (40 + 60 * (1 - Exp(-absValue / 1500)))
            End If
            
        Case PHASE_ENDGAME
            ' �I�ՁF�ΐ������d���������`���̋����֐�
            result = 100 * (1 - Exp(-Abs(value) / 1200)) * Sgn(value)
    End Select
    
    ' -100�`+100�͈̔͂ɐ���
    If result > 100 Then
        AdvancedNormalization = 100
    ElseIf result < -100 Then
        AdvancedNormalization = -100
    Else
        AdvancedNormalization = Round(result, 0)
    End If
End Function

' �p�����[�^�����\�Ȑ��K���֐�
Function ParametricNormalization(value As Integer, phase As Integer) As Integer
    Dim scaleParam As Double, steepness As Double, threshold As Double
    Dim result As Double
    
    ' �i�K�ʃp�����[�^�ݒ�
    Select Case phase
        Case PHASE_OPENING
            scaleParam = 1200      ' ���傫�Ȓl�܂ōl��
            steepness = 0.8   ' ���ɂ₩�ȕω�
            threshold = 100   ' 臒l
            
        Case PHASE_MIDGAME
            scaleParam = 1000      ' �W���I�ȃX�P�[��
            steepness = 1#    ' �W���I�ȋ}�s��
            threshold = 150   ' �����x��臒l
            
        Case PHASE_ENDGAME
            scaleParam = 800       ' ���q���ɔ���
            steepness = 1.2   ' ���}�s�ȕω�
            threshold = 200   ' ��荂��臒l
    End Select
    
    ' �����\�ȃV�O���C�h�֐�
    Dim adjustedValue As Double
    adjustedValue = (value - threshold) * steepness / scaleParam
    
    ' tanh�֐��Ő��K��
    Dim expPos As Double, expNeg As Double
    expPos = Exp(adjustedValue)
    expNeg = Exp(-adjustedValue)
    result = 100 * (expPos - expNeg) / (expPos + expNeg)
    
    ' �͈͐���
    If result > 100 Then
        ParametricNormalization = 100
    ElseIf result < -100 Then
        ParametricNormalization = -100
    Else
        ParametricNormalization = Round(result, 0)
    End If
End Function

' ���K��������I������֐�
Function NormalizeEvaluationImproved(rawEval As Integer, phase As Integer) As Integer
    ' 3�̐��K����������I���\
    ' 1. ���炩�ȃV�O���C�h/tanh�֐�
    ' NormalizeEvaluationImproved = NormalizeEvaluationSmooth(rawEval, phase)
    ' 2. ��荂�x�ȕ����֐�
    NormalizeEvaluationImproved = AdvancedNormalization(rawEval, phase)
    ' 3. �����\�ȃp�����[�^�t���֐�
    ' NormalizeEvaluationImproved = ParametricNormalization(rawEval, phase)
End Function


' ���Օ]���i�@���͏d���A�ΐ��͍T���߂Ɂj
Function EvaluateOpening(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' 1. �p�̉��l
    If IsCorner(row, col) Then
        Score = Score + 5000
    End If
    
    ' 2. ����Ɋp��^�����ւ̌���
    Dim cornerGift As Integer
    cornerGift = CheckCornerGift(row, col, Player)
    If cornerGift > 0 Then
        Score = Score - (cornerGift * 3000)  ' 1�̊p��-3000�_
    End If
    
    ' 3. �댯�Ȋp���ӂւ̌���
    Dim dangerLevel As Integer
    dangerLevel = GetCornerDangerLevel(row, col)
    Score = Score - (dangerLevel * 200)
    
    ' 4. �@���͂̕]��
    Dim myMobility As Integer, oppMobility As Integer
    myMobility = CountValidMoves(Player)
    oppMobility = CountValidMoves(opponent)
    Score = Score + (myMobility - oppMobility) * 100
    
    ' 5. ��������̉��l
    If IsCenterRegion(row, col) Then
        Score = Score + 80
    End If
    
    ' 6. �ΐ��͍T���߂�
    Dim myStones As Integer, oppStones As Integer
    Call CountStones(myStones, oppStones, Player)
    Score = Score - (myStones - oppStones) * 30
    
    ' 7. �Ђ�����Ԃ��ΐ�
    Dim flippedCount As Integer
    flippedCount = CountFlips(row, col, Player)
    Score = Score + flippedCount * 15
    
    EvaluateOpening = Score
End Function

' ���Օ]���i�o�����X�d���j
Function EvaluateMidgame(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' 1. �p�̉��l
    If IsCorner(row, col) Then
        Score = Score + 4000
    End If
    
    ' 2. ����Ɋp��^�����ւ̌���
    Dim cornerGift As Integer
    cornerGift = CheckCornerGift(row, col, Player)
    If cornerGift > 0 Then
        Score = Score - (cornerGift * 2500)
    End If
    
    ' 3. �댯�Ȋp���ӂւ̔���
    Dim dangerLevel As Integer
    dangerLevel = GetCornerDangerLevel(row, col)
    Score = Score - (dangerLevel * 150)
    
    ' 4. ���S�ȕӂ̐헪�I���l
    If IsEdge(row, col) And dangerLevel = 0 Then
        Score = Score + GetEdgeValue(row, col)
    End If
    
    ' 5. �@���́i�d�v�j
    Dim myMobility As Integer, oppMobility As Integer
    myMobility = CountValidMoves(Player)
    oppMobility = CountValidMoves(opponent)
    Score = Score + (myMobility - oppMobility) * 80
    
    ' 6. �m��΂̕]��
    Score = Score + CountStableStones(Player) * 50
    
    ' 7. �ΐ��͒����I�ɕ]��
    Dim myStones As Integer, oppStones As Integer
    Call CountStones(myStones, oppStones, Player)
    Score = Score + (myStones - oppStones) * 10
    
    EvaluateMidgame = Score
End Function

' �I�Օ]���i�ΐ��d���j
Function EvaluateEndgame(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' 1. �p�̉��l
    If IsCorner(row, col) Then
        Score = Score + 3000
    End If
    
    ' 2. ����Ɋp��^�����ւ̌���
    Dim cornerGift As Integer
    cornerGift = CheckCornerGift(row, col, Player)
    If cornerGift > 0 Then
        Score = Score - (cornerGift * 2000)
    End If
    
    ' 3. �ΐ����ŏd�v
    Dim myStones As Integer, oppStones As Integer
    Call CountStones(myStones, oppStones, Player)
    Dim stoneDiff As Integer
    stoneDiff = myStones - oppStones
    Score = Score + stoneDiff * 150
    
    ' 4. �m��΂̉��l
    Dim myStable As Integer, oppStable As Integer
    myStable = CountStableStones(Player)
    oppStable = CountStableStones(opponent)
    Score = Score + (myStable - oppStable) * 100
    
    ' 5. �@����
    Dim myMobility As Integer, oppMobility As Integer
    myMobility = CountValidMoves(Player)
    oppMobility = CountValidMoves(opponent)
    Score = Score + (myMobility - oppMobility) * 60
    
    ' 6. �p���e�B�i�c��萔�̋��j
    Dim emptySquares As Integer
    emptySquares = BOARD_SIZE * BOARD_SIZE - CountTotalStones()
    If emptySquares <= 12 Then
        Dim parityBonus As Integer
        parityBonus = CalculateParityBonus(emptySquares, myMobility, oppMobility)
        Score = Score + parityBonus
    End If
    
    ' 7. �Ђ�����Ԃ��ΐ�
    Dim flippedCount As Integer
    flippedCount = CountFlips(row, col, Player)
    Score = Score + flippedCount * 25
    
    EvaluateEndgame = Score
End Function

' �p���ǂ�������
Function IsCorner(row As Integer, col As Integer) As Boolean
    IsCorner = (row = 1 Or row = BOARD_SIZE) And (col = 1 Or col = BOARD_SIZE)
End Function

' �ӂ��ǂ�������
Function IsEdge(row As Integer, col As Integer) As Boolean
    IsEdge = (row = 1 Or row = BOARD_SIZE Or col = 1 Or col = BOARD_SIZE) And Not IsCorner(row, col)
End Function

' �����̈悩�ǂ�������
Function IsCenterRegion(row As Integer, col As Integer) As Boolean
    Dim center As Integer
    center = BOARD_SIZE / 2
    IsCenterRegion = (row >= center - 1 And row <= center + 2) And (col >= center - 1 And col <= center + 2)
End Function

' �ӂ̐헪�I���l���v�Z
Function GetEdgeValue(row As Integer, col As Integer) As Integer
    Dim value As Integer
    value = 200  ' ��{�I�ȕӂ̉��l
    
    ' �p�ɋ߂��ӂقǉ��l������
    If IsCornerAdjacent(row, col) Then
        value = value + 100
    End If
    
    GetEdgeValue = value
End Function

' �p�ɗאڂ���ӂ��ǂ���
Function IsCornerAdjacent(row As Integer, col As Integer) As Boolean
    ' �p�̒�����ɂ����
    IsCornerAdjacent = (row = 1 And (col = 1 Or col = BOARD_SIZE)) Or _
                      (row = BOARD_SIZE And (col = 1 Or col = BOARD_SIZE)) Or _
                      (col = 1 And (row = 1 Or row = BOARD_SIZE)) Or _
                      (col = BOARD_SIZE And (row = 1 Or row = BOARD_SIZE))
End Function

' �m��΁i��΂ɂЂ�����Ԃ���Ȃ��΁j���J�E���g
Function CountStableStones(Player As Integer) As Integer
    Dim stableCount As Integer
    Dim i As Integer, j As Integer
    
    ' �ȈՔŁF�p�Ƃ��̗אڂ���m��΂̂݃J�E���g
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

' �w��ʒu�̐΂��m��΂��ǂ�������
Function IsStableStone(row As Integer, col As Integer, Player As Integer) As Boolean
    ' �p�͏�Ɋm���
    If IsCorner(row, col) Then
        IsStableStone = True
        Exit Function
    End If
    
    ' �p����A������ӂ̐΂ŁA�Ԃɋ󂫂��Ȃ��ꍇ
    If IsEdge(row, col) Then
        IsStableStone = IsStableEdge(row, col, Player)
    Else
        IsStableStone = False
    End If
End Function

' �ӂ̐΂��m��΂��ǂ�������
Function IsStableEdge(row As Integer, col As Integer, Player As Integer) As Boolean
    ' �ȈՔ���F�p�ɗאڂ��A�p�������v���C���[�̐΂̏ꍇ
    If row = 1 Then  ' ���
        If col > 1 And gameBoard(1, 1) = Player Then
            IsStableEdge = True
        ElseIf col < BOARD_SIZE And gameBoard(1, BOARD_SIZE) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    ElseIf row = BOARD_SIZE Then  ' ����
        If col > 1 And gameBoard(BOARD_SIZE, 1) = Player Then
            IsStableEdge = True
        ElseIf col < BOARD_SIZE And gameBoard(BOARD_SIZE, BOARD_SIZE) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    ElseIf col = 1 Then  ' ����
        If row > 1 And gameBoard(1, 1) = Player Then
            IsStableEdge = True
        ElseIf row < BOARD_SIZE And gameBoard(BOARD_SIZE, 1) = Player Then
            IsStableEdge = True
        Else
            IsStableEdge = False
        End If
    ElseIf col = BOARD_SIZE Then  ' �E��
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

' ����Ɋp��^����肩�ǂ������`�F�b�N
Function CheckCornerGift(row As Integer, col As Integer, Player As Integer) As Integer
    Dim giftCount As Integer
    Dim opponent As Integer
    opponent = 3 - Player
    
    ' ���̎��ł�����A���肪�p������悤�ɂȂ邩�`�F�b�N
    Dim corners As Variant
    corners = Array(Array(1, 1), Array(1, BOARD_SIZE), Array(BOARD_SIZE, 1), Array(BOARD_SIZE, BOARD_SIZE))
    
    Dim i As Integer
    For i = 0 To 3
        Dim cornerRow As Integer, cornerCol As Integer
        cornerRow = corners(i)(0)
        cornerCol = corners(i)(1)
        
        ' ���̊p���󂢂Ă��āA���肪����悤�ɂȂ邩�`�F�b�N
        If gameBoard(cornerRow, cornerCol) = CELL_EMPTY Then
            If IsValidMove(cornerRow, cornerCol, opponent) Then
                giftCount = giftCount + 1
            End If
        End If
    Next i
    
    CheckCornerGift = giftCount
End Function

' �p���ӂ̊댯�x���x�����ڍׂɔ���
Function GetCornerDangerLevel(row As Integer, col As Integer) As Integer
    Dim dangerLevel As Integer
    
    ' �e�p�ɂ��Ċ댯�x���`�F�b�N
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, 1, 1)                    ' ����
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, 1, BOARD_SIZE)          ' �E��
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, BOARD_SIZE, 1)          ' ����
    dangerLevel = dangerLevel + CheckSingleCornerDanger(row, col, BOARD_SIZE, BOARD_SIZE) ' �E��
    
    GetCornerDangerLevel = dangerLevel
End Function

' ����̊p�ɑ΂���댯�x���`�F�b�N
Function CheckSingleCornerDanger(row As Integer, col As Integer, cornerRow As Integer, cornerCol As Integer) As Integer
    ' ���̊p�����ɖ��܂��Ă���ꍇ�͊댯�Ȃ�
    If gameBoard(cornerRow, cornerCol) <> CELL_EMPTY Then
        CheckSingleCornerDanger = 0
        Exit Function
    End If
    
    ' X-square�i�Ίp���אځj: �ł��댯
    If row = cornerRow + IIf(cornerRow = 1, 1, -1) And col = cornerCol + IIf(cornerCol = 1, 1, -1) Then
        CheckSingleCornerDanger = 5
    ' C-square�i�����אځj: ���Ɋ댯
    ElseIf (row = cornerRow And Abs(col - cornerCol) = 1) Or (col = cornerCol And Abs(row - cornerRow) = 1) Then
        CheckSingleCornerDanger = 4
    ' A-square�i�p����2�ڂ̕Ӂj: ���댯
    ElseIf (row = cornerRow And Abs(col - cornerCol) = 2) Or (col = cornerCol And Abs(row - cornerRow) = 2) Then
        CheckSingleCornerDanger = 2
    ' B-square�iX-square�ׁ̗j: �����댯
    ElseIf Abs(row - cornerRow) = 2 And Abs(col - cornerCol) = 1 Then
        CheckSingleCornerDanger = 1
    ElseIf Abs(row - cornerRow) = 1 And Abs(col - cornerCol) = 2 Then
        CheckSingleCornerDanger = 1
    Else
        CheckSingleCornerDanger = 0
    End If
End Function

' �p���e�B�{�[�i�X�̏ڍ׌v�Z
Function CalculateParityBonus(emptySquares As Integer, myMobility As Integer, oppMobility As Integer) As Integer
    Dim bonus As Integer
    
    ' ��{�I�ȃp���e�B
    If emptySquares Mod 2 = 0 Then
        bonus = 80  ' ������c��͗L��
    Else
        bonus = -40 ' ���c��͕s��
    End If
    
    ' �@���͂Ƃ̑g�ݍ��킹
    If myMobility > oppMobility Then
        bonus = bonus + 30  ' �I���������������L��
    ElseIf myMobility < oppMobility Then
        bonus = bonus - 30
    End If
    
    ' �c��萔�����Ȃ��ꍇ�͂��d�v
    If emptySquares <= 6 Then
        bonus = bonus * 2
    ElseIf emptySquares <= 3 Then
        bonus = bonus * 3
    End If
    
    CalculateParityBonus = bonus
End Function

' �ΐ����J�E���g
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

' �Q�[���i�K�����擾
Function GetPhaseName(phase As Integer) As String
    Select Case phase
        Case PHASE_OPENING: GetPhaseName = "����"
        Case PHASE_MIDGAME: GetPhaseName = "����"
        Case PHASE_ENDGAME: GetPhaseName = "�I��"
        Case Else: GetPhaseName = "�s��"
    End Select
End Function


' ���Z�b�g�{�^�����쐬
Sub CreateResetButton()
    Dim btn As Button
    Dim btnRange As Range
    
    ' �Ώۃ��[�N�V�[�g���ݒ肳��Ă��邩�`�F�b�N
    If targetSheet Is Nothing Then Exit Sub
    
    ' �S�{�^�����폜���Ă���쐬
    Call DeleteAllButtons
    
    ' �{�^���̔z�u�ꏊ��ݒ�i�����G���A�������j
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 7, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 6, BOARD_SIZE))
    
    ' �{�^�����쐬
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!RestartGame"  ' ���I�Ƀ��[�N�u�b�N�����w��
        .caption = "���Z�b�g"     ' �{�^���̕\���e�L�X�g
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �Q�[���ĊJ�n�i���Z�b�g�{�^���p�j
Sub RestartGame()
    ' �����̃C�x���g�n���h�����N���A
    Call CleanupEvents
    ' �]���\�����[�h�����Z�b�g
    showEvaluationMode = False
    Call InitializeGame
End Sub

' �C�x���g�n���h�����N���[���A�b�v
Sub CleanupEvents()
    On Error Resume Next
    Application.OnSheetSelectionChange = ""
    On Error GoTo 0
End Sub

' �蓮���̓{�^�����쐬
Sub CreateManualInputButton()
    Dim btn As Button
    Dim btnRange As Range
    
    ' �Ώۃ��[�N�V�[�g���ݒ肳��Ă��邩�`�F�b�N
    If targetSheet Is Nothing Then Exit Sub
    
    ' �{�^���̔z�u�ꏊ��ݒ�
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 8, 1), _
                                    targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE / 2))
    
    ' �{�^�����쐬
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ManualInput"  ' ���I�Ƀ��[�N�u�b�N�����w��
        .caption = "�蓮����"  ' �{�^���̕\���e�L�X�g
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �蓮���͋@�\
Sub ManualInput()
    Dim userInput As String
    Dim col As Integer, row As Integer
    Dim colChar As String
    
    If gameOver Or CurrentPlayer <> BLACK Then
        MsgBox "���݂͎蓮���͂ł��܂���B", vbInformation
        Exit Sub
    End If
    
    userInput = InputBox("���W����͂��Ă������� (��: A4, B3):", "�蓮����", "")
    If userInput = "" Then Exit Sub
    
    ' ���͂����
    userInput = UCase(Trim(userInput))
    If Len(userInput) < 2 Then
        MsgBox "�������`���œ��͂��Ă������� (��: A4, B3)", vbExclamation
        Exit Sub
    End If
    
    colChar = Left(userInput, 1)
    row = Val(Mid(userInput, 2))
    col = Asc(colChar) - 64  ' A=1, B=2, etc.
    
    ' �͈̓`�F�b�N
    If col < 1 Or col > BOARD_SIZE Or row < 1 Or row > BOARD_SIZE Then
        MsgBox "���W���͈͊O�ł��BA1����" & Chr(64 + BOARD_SIZE) & BOARD_SIZE & "�͈̔͂œ��͂��Ă��������B", vbExclamation
        Exit Sub
    End If
    
    ' �L����`�F�b�N�Ǝ��s
    If IsValidMove(row, col, BLACK) Then
        Dim playerEvaluation As Integer
        playerEvaluation = GetUnifiedEvaluation(row, col, BLACK)

        Call SaveGameState(row, col, BLACK, playerEvaluation)
        Call MakeMove(row, col, BLACK)
        Call UpdateDisplay
        
        ' ��̋L�^
        Call RecordMove(row, col, BLACK)
        Call UpdateGameRecordDisplay
        
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If
        
        ' CPU�^�[��
        Call ExecuteCPUTurn
        
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If
        
        ' �v���C���[�̃^�[���ɖ߂�
        Call SwitchToPlayerTurn
    Else
        MsgBox "�����ɂ͒u���܂���B�ʂ̏ꏊ��I��ł��������B", vbExclamation
    End If
End Sub

' �]���\���؂�ւ��{�^�����쐬
Sub CreateEvaluationToggleButton()
    Dim btn As Button
    Dim btnRange As Range
    
    ' �Ώۃ��[�N�V�[�g���ݒ肳��Ă��邩�`�F�b�N
    If targetSheet Is Nothing Then Exit Sub
    
    ' �{�^���̔z�u�ꏊ��ݒ�
    Set btnRange = targetSheet.Range(targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE / 2 + 1), _
                                    targetSheet.Cells(BOARD_SIZE + 8, BOARD_SIZE))
    
    ' �{�^�����쐬
    Set btn = targetSheet.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.width, btnRange.height)
    
    With btn
        .OnAction = ThisWorkbook.Name & "!ToggleEvaluationDisplay"  ' ���I�Ƀ��[�N�u�b�N�����w��
        .caption = "�]���\��: OFF"     ' �{�^���̕\���e�L�X�g
        .Font.Size = 10
        .Font.Bold = False
    End With
End Sub

' �]���\�����[�h�؂�ւ�
Sub ToggleEvaluationDisplay()
    showEvaluationMode = Not showEvaluationMode
    
    ' �{�^���̃L���v�V�������X�V
    Call UpdateEvaluationButtonCaption
    
    ' ��ʕ\�����X�V
    Call UpdateDisplay
End Sub
' �]���\���{�^���̃L���v�V�����X�V
Sub UpdateEvaluationButtonCaption()
    Dim btn As Button
    Dim i As Integer
    
    On Error Resume Next
    ' �]���\���{�^����T���čX�V
    For i = 1 To targetSheet.Buttons.count
        Set btn = targetSheet.Buttons(i)
        If InStr(btn.caption, "�]���\��") > 0 Then
            If showEvaluationMode Then
                btn.caption = "�]���\��: ON"
            Else
                btn.caption = "�]���\��: OFF"
            End If
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

' �S�{�^�����폜
Sub DeleteAllButtons()
    Dim i As Integer
    
    On Error Resume Next  ' �G���[���������Ă����s
    
    ' �Ώۃ��[�N�V�[�g���ݒ肳��Ă���ꍇ�̂ݎ��s
    If Not targetSheet Is Nothing Then
        ' �t���őS�{�^�����폜
        For i = targetSheet.Buttons.count To 1 Step -1
            targetSheet.Buttons(i).Delete
        Next i
    End If
    
    On Error GoTo 0  ' �G���[�n���h�����O�����ɖ߂�
End Sub

' ��ʕ\���X�V
Sub UpdateDisplay()
    Dim i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            With targetSheet.Cells(i, j)
                Select Case gameBoard(i, j)
                    Case CELL_EMPTY
                        ' ��̃Z���̏���
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

' �]���l�ɉ������F����
Function GetEvaluationColor(evalValue As Integer) As Long
    If evalValue >= 90 Then
        GetEvaluationColor = RGB(255, 0, 0)        ' �ԁF�ŗǎ�
    ElseIf evalValue >= 70 Then
        GetEvaluationColor = RGB(255, 100, 0)      ' �Ԟ�F���ɗǂ���
    ElseIf evalValue >= 50 Then
        GetEvaluationColor = RGB(255, 165, 0)      ' �I�����W�F�ǂ���
    ElseIf evalValue >= 30 Then
        GetEvaluationColor = RGB(255, 200, 0)      ' ����F���ǂ���
    ElseIf evalValue >= 10 Then
        GetEvaluationColor = RGB(255, 255, 0)      ' ���F���ʂ̎�
    ElseIf evalValue >= -10 Then
        GetEvaluationColor = RGB(255, 255, 255)    ' ���F�݊p
    ElseIf evalValue >= -30 Then
        GetEvaluationColor = RGB(200, 200, 200)    ' ���D�F��∫��
    ElseIf evalValue >= -50 Then
        GetEvaluationColor = RGB(128, 128, 128)    ' �D�F������
    ElseIf evalValue >= -70 Then
        GetEvaluationColor = RGB(100, 100, 100)    ' �Z�D�F���Ɉ���
    Else
        GetEvaluationColor = RGB(0, 0, 0)          ' ���F�댯�Ȏ�
    End If
End Function

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
        "�X�R�A - ��: " & blackCount & "  ��: " & whiteCount & "  �^�[��: " & moveCount
    
    Dim diffName As String
    Select Case difficulty
        Case EASY: diffName = "����"
        Case MEDIUM: diffName = "����"
        Case HARD: diffName = "�㋉"
    End Select
    
    Dim modeStatus As String
    If showEvaluationMode Then
        modeStatus = " | �]���\��: ON"
    Else
        modeStatus = " | �]���\��: OFF"
    End If
    
    targetSheet.Cells(BOARD_SIZE + 4, 1).value = "��Փx: " & diffName & modeStatus
End Sub

' CPU�^�[�������s
Sub ExecuteCPUTurn()
    CurrentPlayer = WHITE
    
    ' CPU���X�L�b�v����K�v�����邩�`�F�b�N
    If Not HasValidMoves(WHITE) Then
        Call ShowMessage("���ɗL���Ȏ肪����܂���B�X�L�b�v���܂��B")
        Call RecordMove(0, 0, WHITE)  ' �X�L�b�v���L�^
        Call UpdateGameRecordDisplay  ' �����\���X�V
        MsgBox "���ɗL���Ȏ肪����܂���B���̃^�[���ł��B", vbInformation, "�^�[���X�L�b�v"
        Exit Sub
    End If
    
    ' CPU��������s
    Call ShowMessage("���̃^�[���ł��B")
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")
    
    Call CPUTurn
    Call UpdateDisplay
    Call UpdateGameRecordDisplay  ' �����\���X�V
End Sub

' �v���C���[�^�[���ɐ؂�ւ�
Sub SwitchToPlayerTurn()
    CurrentPlayer = BLACK
    
    ' �v���C���[���X�L�b�v����K�v�����邩�`�F�b�N
    If Not HasValidMoves(BLACK) Then
        Call ShowMessage("���ɗL���Ȏ肪����܂���B�X�L�b�v���܂��B")
        Call RecordMove(0, 0, BLACK)  ' �X�L�b�v���L�^
        Call UpdateGameRecordDisplay  ' �����\���X�V
        MsgBox "���ɗL���Ȏ肪����܂���B���̃^�[���ł��B", vbInformation, "�^�[���X�L�b�v"

        ' �ēxCPU�^�[�������s
        Call ExecuteCPUTurn
        Call UpdateDisplay
        Call UpdateGameRecordDisplay  ' �����\���X�V
        
        ' �Q�[���I���`�F�b�N
        If CheckGameEnd() Then
            Call ShowResult
            Exit Sub
        End If
        
        ' ������x�v���C���[�^�[�����`�F�b�N
        Call SwitchToPlayerTurn
    Else
        ' �v���C���[�ɗL���Ȏ肪����ꍇ
        Call ShowMessage("���F�̃^�[���ł��B")
        ' �]���\�����[�h�̏ꍇ�͉�ʂ��X�V���ĕ]���l��\��
        If showEvaluationMode Then
            Call UpdateDisplay
        End If
    End If
End Sub

' �Z���N���b�N����
Public Sub HandleCellClick(Target As Range)
    Call ProcessCellClick(Target)
End Sub

' �L���胊�X�g���擾
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
    
    ' �z��T�C�Y�𒲐�
    If moveCount > 0 Then
        ReDim Preserve validMoves(1 To moveCount)
    Else
        ReDim validMoves(1 To 1)  ' ��z��
    End If
    
    GetValidMoves = validMoves
End Function

' CPU�̎�i��Փx�ɉ����ď����𕪊�j
Sub CPUTurn()
    Dim bestRow As Integer, bestCol As Integer
    Dim evaluation As Integer
    
    ' CPU�ɗL���Ȏ肪���邩�`�F�b�N
    If Not HasValidMoves(WHITE) Then Exit Sub
    
    Select Case difficulty
        Case EASY
            Call CPUTurnEasy(bestRow, bestCol, evaluation)
        Case MEDIUM
            Call CPUTurnMedium(bestRow, bestCol, evaluation)
        Case HARD
            Call CPUTurnHard(bestRow, bestCol, evaluation)
    End Select
    
    ' �ŗǎ�����s
    If bestRow > 0 And bestCol > 0 And IsValidMove(bestRow, bestCol, WHITE) Then
        Call SaveGameState(bestRow, bestCol, WHITE, evaluation)
        Call MakeMove(bestRow, bestCol, WHITE)
        Call RecordMove(bestRow, bestCol, WHITE)
    Else
        ' �����Ȏ肪�I�΂ꂽ�ꍇ�͏����ő��
        Call CPUTurnEasy(bestRow, bestCol, evaluation)
        If bestRow > 0 And bestCol > 0 And IsValidMove(bestRow, bestCol, WHITE) Then
            Call SaveGameState(bestRow, bestCol, WHITE, evaluation)
            Call MakeMove(bestRow, bestCol, WHITE)
            Call RecordMove(bestRow, bestCol, WHITE)
        End If
    End If
End Sub

' �]���l�𓝈�͈͂ɐ��K������֐�
Function NormalizeEvaluation(rawEval As Integer, difficulty As DifficultyLevel) As Integer
    Select Case difficulty
        Case EASY
            ' �����͂��̂܂܁i1-100���x�͈̔́j
            NormalizeEvaluation = rawEval
        Case MEDIUM, HARD
            ' �����E�㋉��-100�`+100�ɐ��K��
            If rawEval >= 2000 Then
                NormalizeEvaluation = 100      ' ����
            ElseIf rawEval >= 1000 Then
                NormalizeEvaluation = 80       ' �傫���L��
            ElseIf rawEval >= 500 Then
                NormalizeEvaluation = 50       ' �L��
            ElseIf rawEval >= 100 Then
                NormalizeEvaluation = 20       ' ���L��
            ElseIf rawEval > -100 Then
                NormalizeEvaluation = rawEval / 5  ' �������i-20�`+20�j
            ElseIf rawEval > -500 Then
                NormalizeEvaluation = -20      ' ���s��
            ElseIf rawEval > -1000 Then
                NormalizeEvaluation = -50      ' �s��
            ElseIf rawEval > -2000 Then
                NormalizeEvaluation = -80      ' �傫���s��
            Else
                NormalizeEvaluation = -100     ' ��
            End If
    End Select
End Function

' �����i�]���֐��̂݁j
Sub CPUTurnEasy(ByRef bestRow As Integer, ByRef bestCol As Integer, ByRef bestEval As Integer)
    Dim validMoves() As ValidMove
    Dim bestScore As Integer
    Dim i As Integer
    
    validMoves = GetValidMoves(WHITE)
    bestScore = -1000
    bestRow = 0
    bestCol = 0
    
    ' �L���肪���݂���ꍇ�̂ݏ���
    If UBound(validMoves) > 0 And validMoves(1).row > 0 Then
        For i = 1 To UBound(validMoves)
            If validMoves(i).Score > bestScore Then
                bestScore = validMoves(i).Score
                bestRow = validMoves(i).row
                bestCol = validMoves(i).col
            End If
        Next i
    End If
    
    ' ���K�����ĕԂ�
    bestEval = NormalizeEvaluation(bestScore, EASY)
End Sub

' �����i��-���T�� �[�x3�j
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
    
    ' �L���肪���݂���ꍇ�̂ݏ���
    If UBound(validMoves) > 0 And validMoves(1).row > 0 Then
        For i = 1 To UBound(validMoves)
            ' �{�[�h�𕜌����Ă���������
            Call CopyBoard(tempBoard, gameBoard)
            Call MakeMove(validMoves(i).row, validMoves(i).col, WHITE)
            
            ' ��-���T���i�[�x3�j
            Score = AlphaBeta(3, -9999, 9999, False)
            
            If Score > bestScore Then
                bestScore = Score
                bestRow = validMoves(i).row
                bestCol = validMoves(i).col
            End If
            
            ' �{�[�h�����ɖ߂�
            Call CopyBoard(tempBoard, gameBoard)
        Next i
    End If
    
    ' �ŏI�I�Ƀ{�[�h��Ԃ𕜌�
    Call CopyBoard(tempBoard, gameBoard)
    ' ���K�����ĕԂ�
    bestEval = NormalizeEvaluation(bestScore, MEDIUM)
End Sub

' �㋉�i��-���T�� �[�x5�j
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
    
    ' �L���肪���݂���ꍇ�̂ݏ���
    If UBound(validMoves) > 0 And validMoves(1).row > 0 Then
        For i = 1 To UBound(validMoves)
            ' �{�[�h�𕜌����Ă���������
            Call CopyBoard(tempBoard, gameBoard)
            Call MakeMove(validMoves(i).row, validMoves(i).col, WHITE)
            
            ' ��-���T���i�[�x5�j
            Score = AlphaBeta(5, -9999, 9999, False)
            
            If Score > bestScore Then
                bestScore = Score
                bestRow = validMoves(i).row
                bestCol = validMoves(i).col
            End If
            
            ' �{�[�h�����ɖ߂�
            Call CopyBoard(tempBoard, gameBoard)
        Next i
    End If
    
    ' �ŏI�I�Ƀ{�[�h��Ԃ𕜌�
    Call CopyBoard(tempBoard, gameBoard)
    ' ���K�����ĕԂ�
    bestEval = NormalizeEvaluation(bestScore, HARD)
End Sub

' ��-���T��
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
    
    ' ���݂̃{�[�h��Ԃ��o�b�N�A�b�v
    Call CopyBoard(gameBoard, tempBoard)
    
    If isMaximizing Then
        bestScore = -9999
        Player = WHITE
    Else
        bestScore = 9999
        Player = BLACK
    End If
    
    ' �L���胊�X�g���擾
    validMoves = GetValidMoves(Player)
    
    ' �L���肪���݂��Ȃ��ꍇ
    If UBound(validMoves) = 0 Or validMoves(1).row = 0 Then
        ' �p�X���đ���̃^�[��
        AlphaBeta = AlphaBeta(depth - 1, alpha, beta, Not isMaximizing)
        Call CopyBoard(tempBoard, gameBoard)
        Exit Function
    End If
    
    For i = 1 To UBound(validMoves)
        ' �{�[�h�𕜌����Ă���������
        Call CopyBoard(tempBoard, gameBoard)
        Call MakeMove(validMoves(i).row, validMoves(i).col, Player)
        
        Score = AlphaBeta(depth - 1, alpha, beta, Not isMaximizing)
        
        If isMaximizing Then
            If Score > bestScore Then bestScore = Score
            If Score > alpha Then alpha = Score
            If beta <= alpha Then
                ' �{�[�h�����ɖ߂��Ă���v���[�j���O
                Call CopyBoard(tempBoard, gameBoard)
                Exit For  ' ��-�J�b�g
            End If
        Else
            If Score < bestScore Then bestScore = Score
            If Score < beta Then beta = Score
            If beta <= alpha Then
                ' �{�[�h�����ɖ߂��Ă���v���[�j���O
                Call CopyBoard(tempBoard, gameBoard)
                Exit For  ' ��-�J�b�g
            End If
        End If
        
        ' �{�[�h�����ɖ߂�
        Call CopyBoard(tempBoard, gameBoard)
    Next i
    
    ' �ŏI�I�Ƀ{�[�h��Ԃ𕜌�
    Call CopyBoard(tempBoard, gameBoard)
    AlphaBeta = bestScore
End Function

' �{�[�h�S�̂̕]��
Function EvaluateBoardPosition() As Integer
    Dim Score As Integer, i As Integer, j As Integer
    Dim whiteCount As Integer, blackCount As Integer
    Dim whiteMobility As Integer, blackMobility As Integer
    
    ' �΂̐��ƈʒu���l��]��
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
    
    ' �@���́i�L����̐��j��]��
    whiteMobility = CountValidMoves(WHITE)
    blackMobility = CountValidMoves(BLACK)
    Score = Score + (whiteMobility - blackMobility) * 10
    
    ' �Q�[���I�Ղł͐ΐ����d��
    If whiteCount + blackCount > BOARD_SIZE * BOARD_SIZE * 0.8 Then
        Score = Score + (whiteCount - blackCount) * 50
    End If
    
    EvaluateBoardPosition = Score
End Function

' �ʒu�̉��l���擾
Function GetPositionValue(row As Integer, col As Integer) As Integer
    ' �p
    If (row = 1 Or row = BOARD_SIZE) And (col = 1 Or col = BOARD_SIZE) Then
        GetPositionValue = 100
    ' �p�ׁ̗i�댯�n�сj
    ElseIf ((row = 1 Or row = BOARD_SIZE) And (col = 2 Or col = BOARD_SIZE - 1)) Or _
           ((col = 1 Or col = BOARD_SIZE) And (row = 2 Or row = BOARD_SIZE - 1)) Then
        GetPositionValue = -20
    ' ��
    ElseIf row = 1 Or row = BOARD_SIZE Or col = 1 Or col = BOARD_SIZE Then
        GetPositionValue = 10
    ' �����t��
    Else
        GetPositionValue = 1
    End If
End Function

' �{�[�h�R�s�[
Sub CopyBoard(sourceBoard() As Integer, targetBoard() As Integer)
    Dim i As Integer, j As Integer
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            targetBoard(i, j) = sourceBoard(i, j)
        Next j
    Next i
End Sub

' �Q�[����ԕۑ�
Sub SaveGameState(row As Integer, col As Integer, Player As Integer, evaluation As Integer)
    historyCount = historyCount + 1
    
    ' �z��T�C�Y���g��
    If historyCount > UBound(gameHistory) Then
        ReDim Preserve gameHistory(0 To historyCount + 50)
    End If
    
    With gameHistory(historyCount)
        ' �{�[�h��Ԃ��R�s�[
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

' ��̋L�^
Sub RecordMove(row As Integer, col As Integer, Player As Integer)
    moveCount = moveCount + 1
    
    If moveCount > UBound(moveHistory) Then
        ReDim Preserve moveHistory(1 To moveCount + 50)
    End If
    
    Dim playerName As String
    If Player = BLACK Then playerName = "��" Else playerName = "��"
    
    Dim moveDescription As String
    If row = 0 And col = 0 Then
        ' �X�L�b�v�̏ꍇ
        moveDescription = "�X�L�b�v"
    Else
        Dim colLetter As String
        colLetter = Chr(64 + col)  ' 1->A, 2->B, etc.
        moveDescription = colLetter & row
    End If
    
    moveHistory(moveCount) = Format(moveCount, "000") & ": " & playerName & " " & _
                            moveDescription & " (" & Format(Now, "hh:mm:ss") & ")"
End Sub

' �����ۑ�
Sub SaveGameRecord()
    Dim recordSheet As Worksheet
    Dim lastRow As Integer, i As Integer
    Dim gameStartTime As String
    
    ' �Q�[���J�n�������Q�[��ID�Ƃ��Ďg�p
    gameStartTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    ' �����V�[�g���쐬�܂��͎擾
    On Error Resume Next
    Set recordSheet = targetWorkbook.Worksheets("����")
    On Error GoTo 0
    
    If recordSheet Is Nothing Then
        Set recordSheet = targetWorkbook.Worksheets.Add
        recordSheet.Name = "����"
        
        ' �w�b�_�[�ݒ�
        recordSheet.Cells(1, 1).value = "�Q�[��ID"
        recordSheet.Cells(1, 2).value = "�^�[��"
        recordSheet.Cells(1, 3).value = "�v���C���["
        recordSheet.Cells(1, 4).value = "���W"
        recordSheet.Cells(1, 5).value = "����"
        recordSheet.Cells(1, 6).value = "�]���l"
        recordSheet.Cells(1, 7).value = "��Փx"
        recordSheet.Cells(1, 8).value = "�{�[�h�T�C�Y"
        recordSheet.Cells(1, 9).value = "�X�R�A"
        recordSheet.Cells(1, 10).value = "����"
        
        ' �w�b�_�[�����ݒ�
        With recordSheet.Range("A1:J1")
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 200)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' �񕝒���
        recordSheet.Columns("A").ColumnWidth = 18  ' �Q�[��ID
        recordSheet.Columns("B").ColumnWidth = 6   ' �^�[��
        recordSheet.Columns("C").ColumnWidth = 8   ' �v���C���[
        recordSheet.Columns("D").ColumnWidth = 8   ' ���W
        recordSheet.Columns("E").ColumnWidth = 10  ' ����
        recordSheet.Columns("F").ColumnWidth = 8   ' �]���l
        recordSheet.Columns("G").ColumnWidth = 8   ' ��Փx
        recordSheet.Columns("H").ColumnWidth = 10  ' �{�[�h�T�C�Y
        recordSheet.Columns("I").ColumnWidth = 15  ' �X�R�A
        recordSheet.Columns("J").ColumnWidth = 12  ' ����
    End If
    
    ' �ŏI���ʂ��擾
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
        winner = "��"
    ElseIf finalWhiteCount > finalBlackCount Then
        winner = "��"
    Else
        winner = "��������"
    End If
    
    Dim diffName As String
    Select Case difficulty
        Case EASY: diffName = "����"
        Case MEDIUM: diffName = "����"
        Case HARD: diffName = "�㋉"
    End Select
    
    ' �e����ʂ̍s�ɋL�^
    lastRow = recordSheet.Cells(recordSheet.Rows.count, 1).End(xlUp).row
    
    For i = 1 To moveCount
        lastRow = lastRow + 1
        
        ' ��̏������
        Dim moveInfo As String, playerName As String, movePos As String
        Dim timeInfo As String, moveNum As Integer
        Dim evalValue As Integer
        Dim currentScore As String

        moveInfo = moveHistory(i)
        
        ' �^�[���𒊏o
        moveNum = Val(Left(moveInfo, 3))
        
        ' �v���C���[���𒊏o
        If InStr(moveInfo, "��") > 0 Then
            playerName = "��"
        Else
            playerName = "��"
        End If
        
        ' ���W�𒊏o
        If InStr(moveInfo, "�X�L�b�v") > 0 Then
            movePos = "�X�L�b�v"
            evalValue = 0
        Else
            ' ���W�����𒊏o�i��FA4, B3�Ȃǁj
            Dim startPos As Integer, endPos As Integer
            startPos = InStr(moveInfo, playerName) + Len(playerName) + 1
            endPos = InStr(startPos, moveInfo, " (") - 1
            movePos = Trim(Mid(moveInfo, startPos, endPos - startPos + 1))
            
            ' �]���l���擾�i��������j
            If i <= historyCount Then
                evalValue = gameHistory(i).evaluation
            Else
                evalValue = 0
            End If
        End If
        
        ' �����𒊏o
        startPos = InStr(moveInfo, "(") + 1
        endPos = InStr(moveInfo, ")") - 1
        timeInfo = Mid(moveInfo, startPos, endPos - startPos + 1)
        
                
        ' ���̎��_�ł̃X�R�A���v�Z
        If i <= historyCount Then
            currentScore = CalculateScoreAtMove(i)
        Else
            currentScore = "0-0"  ' �G���[���̃f�t�H���g�l
        End If
        
        ' �f�[�^���L�^
        With recordSheet
            .Cells(lastRow, 1).value = gameStartTime           ' �Q�[��ID
            .Cells(lastRow, 2).value = moveNum                 ' �^�[��
            .Cells(lastRow, 3).value = playerName              ' �v���C���[
            .Cells(lastRow, 4).value = movePos                 ' ���W
            .Cells(lastRow, 5).value = timeInfo                ' ����
            .Cells(lastRow, 6).value = evalValue               ' �]���l
            .Cells(lastRow, 7).value = diffName                ' ��Փx
            .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE  ' �{�[�h�T�C�Y
            .Cells(lastRow, 9).value = currentScore            ' ���̎��_�ł̃X�R�A
            .Cells(lastRow, 10).value = IIf(i = moveCount, winner, "-")  ' �ŏI��̂ݏ��҂�\��
        End With
    Next i
    
    ' �Q�[���I���s��ǉ�
    lastRow = lastRow + 1
    With recordSheet
        .Cells(lastRow, 1).value = gameStartTime
        .Cells(lastRow, 2).value = "---"
        .Cells(lastRow, 3).value = "�Q�[���I��"
        .Cells(lastRow, 4).value = "---"
        .Cells(lastRow, 5).value = Format(Now, "hh:mm:ss")
        .Cells(lastRow, 6).value = ""
        .Cells(lastRow, 7).value = diffName
        .Cells(lastRow, 8).value = BOARD_SIZE & "x" & BOARD_SIZE
        .Cells(lastRow, 9).value = "��" & finalBlackCount & "-" & finalWhiteCount & "��"  ' �ŏI�X�R�A
        .Cells(lastRow, 10).value = winner
        
        ' ��؂�s�̏����ݒ�
        .Range(.Cells(lastRow, 1), .Cells(lastRow, 10)).Interior.color = RGB(240, 240, 240)
        .Range(.Cells(lastRow, 1), .Cells(lastRow, 10)).Font.Bold = True
    End With
    
    ' ��s��ǉ��i���̃Q�[���Ƃ̋�؂�j
    lastRow = lastRow + 1
    
    MsgBox "������ۑ����܂����B" & vbCrLf & _
           "�V�[�g�u�����v�Ŋm�F�ł��܂��B" & vbCrLf, vbInformation, "�����ۑ�����"
End Sub

' �w�肵����Ԏ��_�ł̃X�R�A���v�Z����֐�
Function CalculateScoreAtMove(moveIndex As Integer) As String
    Dim blackCount As Integer, whiteCount As Integer
    Dim i As Integer, j As Integer
    
    ' ���������݂��Ȃ��ꍇ�̓G���[����
    If moveIndex < 1 Or moveIndex > historyCount Then
        CalculateScoreAtMove = "�G���["
        Exit Function
    End If
    
    ' �w�肵����Ԍ�̃{�[�h��Ԃ���΂̐��𐔂���
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
    
    ' �X�R�A��������쐬
    CalculateScoreAtMove = "��" & blackCount & "-" & whiteCount & "��"
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

' �L����̐����J�E���g
Function CountValidMoves(Player As Integer) As Integer
    Dim count As Integer, i As Integer, j As Integer
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If IsValidMove(i, j, Player) Then count = count + 1
        Next j
    Next i
    
    CountValidMoves = count
End Function

' �X�L�b�v�������Ǘ�
Sub CheckAndHandleSkip()
    Dim playerHasMoves As Boolean, cpuHasMoves As Boolean
    Dim skipMessage As String
    Dim skipCount As Integer
    
    ' �ő�2��܂ł̘A���X�L�b�v������
    skipCount = 0
    
    Do While skipCount < 2
        playerHasMoves = HasValidMoves(BLACK)
        cpuHasMoves = HasValidMoves(WHITE)
        
        ' �����Ɏ肪�Ȃ��ꍇ�̓Q�[���I��
        If Not playerHasMoves And Not cpuHasMoves Then
            gameOver = True
            Exit Sub
        End If
        
        ' ���݂̃v���C���[�Ɏ肪�Ȃ��ꍇ
        If CurrentPlayer = BLACK And Not playerHasMoves Then
            skipMessage = "���ɗL���Ȏ肪����܂���B���̃^�[���ł��B"
            Call ShowMessage(skipMessage)
            Call RecordMove(0, 0, BLACK)  ' �X�L�b�v���L�^
            Call UpdateGameRecordDisplay  ' �����\���X�V
            MsgBox skipMessage, vbInformation, "�^�[���X�L�b�v"
            
            CurrentPlayer = WHITE
            skipCount = skipCount + 1
            
            ' ���̃^�[�����`�F�b�N
            If HasValidMoves(WHITE) Then
                Exit Sub  ' ���Ɏ肪����̂ŃX�L�b�v�����I��
            End If
            
        ElseIf CurrentPlayer = WHITE And Not cpuHasMoves Then
            skipMessage = "���ɗL���Ȏ肪����܂���B���̃^�[���ł��B"
            Call ShowMessage(skipMessage)
            Call RecordMove(0, 0, WHITE)  ' �X�L�b�v���L�^
            Call UpdateGameRecordDisplay  ' �����\���X�V
            MsgBox skipMessage, vbInformation, "�^�[���X�L�b�v"
            
            CurrentPlayer = BLACK
            skipCount = skipCount + 1
            
            ' ���̃^�[�����`�F�b�N
            If HasValidMoves(BLACK) Then
                Exit Sub  ' ���Ɏ肪����̂ŃX�L�b�v�����I��
            End If
            
        Else
            ' ���݂̃v���C���[�Ɏ肪����ꍇ�͏I��
            Exit Sub
        End If
    Loop
    
    ' 2��A���ŃX�L�b�v�����������ꍇ�̓Q�[���I��
    If skipCount >= 2 Then
        gameOver = True
    End If
End Sub

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

' ��̕]��
Function EvaluateMove(row As Integer, col As Integer, Player As Integer) As Integer
    Dim Score As Integer
    
    ' �p�̕]��
    If (row = 1 Or row = BOARD_SIZE) And (col = 1 Or col = BOARD_SIZE) Then
        Score = 100
    ' �ӂ̕]��
    ElseIf row = 1 Or row = BOARD_SIZE Or col = 1 Or col = BOARD_SIZE Then
        Score = 10
    Else
        Score = 1
    End If
    
    ' ����΂̐������Z
    Score = Score + CountFlips(row, col, Player)
    
    ' �������x���Ƃ��Đ��K��
    EvaluateMove = NormalizeEvaluation(Score, EASY)
End Function

' �Ђ�����Ԃ���΂̐����J�E���g
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
    
    ' �C�x���g�n���h�����N���[���A�b�v
    Call CleanupEvents
    
    For i = 1 To BOARD_SIZE
        For j = 1 To BOARD_SIZE
            If gameBoard(i, j) = BLACK Then blackCount = blackCount + 1
            If gameBoard(i, j) = WHITE Then whiteCount = whiteCount + 1
        Next j
    Next i
    
    If blackCount > whiteCount Then
        result = "���̏����ł��I"
    ElseIf whiteCount > blackCount Then
        result = "���̏����ł��I"
    Else
        result = "���������ł��I"
    End If
    
    Call ShowMessage(result)
    targetSheet.Cells(BOARD_SIZE + 5, 1).value = _
        "�ŏI�X�R�A - ��: " & blackCount & "  ��: " & whiteCount
    
    ' �ŏI�����\���X�V
    Call UpdateGameRecordDisplay
    
    ' �����Ŋ����ۑ�
    Call SaveGameRecord
    
    MsgBox result & vbCrLf & vbCrLf & _
           "�ŏI�X�R�A" & vbCrLf & _
           "��: " & blackCount & vbCrLf & _
           "��: " & whiteCount & vbCrLf & vbCrLf & _
           "�����������ۑ����܂����B" & vbCrLf & _
           "�V�����Q�[�����n�߂�ɂ� ���Z�b�g�{�^�� ���N���b�N���Ă��������B", _
           vbInformation, "�Q�[���I��"
End Sub





