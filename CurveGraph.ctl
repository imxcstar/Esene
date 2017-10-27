VERSION 5.00
Begin VB.UserControl CurveGraph 
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   2700
   ScaleWidth      =   7200
   Begin VB.PictureBox picGraph 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   480
      ScaleHeight     =   1935
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   360
      Width           =   6255
      Begin VB.Shape ShapeLegend 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   240
         Index           =   1
         Left            =   1200
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblLegend 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   3720
         TabIndex        =   1
         Top             =   80
         Width           =   90
      End
   End
End
Attribute VB_Name = "CurveGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***********************************************************************************
' ��    �ܣ�����ʵʱ���ߣ����� Windows ���������CPUʹ�������ߣ�
' ʹ�÷�����
' ��    �ߣ�����������
' ��    Ȩ������������
' �������ڣ�2009-05-01
' ��    վ��http://hewanglan34512.cech.com.cn
' E - mail��hewanglan34512@163.com
' ��    �������ޣ�
' ��    �£�
' 2��2009-09-25~2009-09-26
'              (1) ��������������ԣ���һ���������ϣ�һ�οɻ��������ߣ�1��MAX_CURVECOUNT �������á���
'                  ע�⣺����������ÿ�����߶��ử����ʹû���� AddValue ���ֵ����� blnShouldDrawCurve ���������Ƿ�Ҫ��������
'                  �����������뻭�����پ����ö��٣���Ҫ�ж���ġ�
'              (2) ������ɫ ���ԸĶ���֧�ֶ������߲�ͬ��ɫ
'              (3) �������� ���ԸĶ���֧�ֶ������߲�ͬ����
'              (4) ������ݺ�����AddValue�������Ӳ���ȷ����ӵ���һ�����ߣ���һ�����������ݣ�����ÿ�����߶�������ݣ�����
'              (5) ���ͼ��˵��������FixLegend������ �� A����Ҫ�޸�ʱ�����ô˺������ɣ�
'                  ���е�Windowsϵͳ������Ϊ����ʱ������ �� ������ֻ����ʾΪ�����ߣ���Ϊ��������Ϳ���������ʾ��
'              (6) �Ƿ���ʾͼ��˵�����ɵ�������ĳһ������ͼ��
'              (7) ���ͼ���������ԣ�֧�ֲ�ͬͼ�����岻ͬ
'              (8) ����ͼ��˵�����Ե����ƶ���������϶���
' 1��2009-05-17
'              (1) ��Ӻ������������ = ClearAll�����ԭ�������ߣ�
'              (2) ���Ӵ�ֱ������Сֵ���� = MinVertical������֧�ָ�����
'              (3) Ϊ���������������˵�����֣������û�ʹ�ã�
' * (******* �����뱣��������Ϣ *******)
' **************************************************************************************

Option Explicit
' ### API �������� -------------------------------------------------------------------------------------

' ### ö������ -----------------------------------------------------------------------------------------
' �߿�����' ע�⣺ͼƬ��߿��������� 0���ؼ��߿������� UserControl �Ľ��п��ƣ�
' �� UserControl.Appearance ����Ϊ 0 - Flat Ҳ������ picGraph ͼƬ��߿������ơ�
Public Enum BoderStyleEnum
    [None] = 0
    [Fixed Single] = 1
End Enum
' ��ֱ�������ƶ���ʽ��Ҳ�������ƶ�����
Public Enum MovingGridEnum
    [Not Moving] = 0
    [Left to Right] = 1
    [Right to Left] = 2
End Enum
' �������͡�ʵ�ߣ����ߣ�
Public Enum CurveTypeEnum
    [Solid] = 0
    [Dot] = 1
End Enum
'Download by http://www.NewXing.com
' ### �������� -----------------------------------------------------------------------------------------
Private Const LAST_LINE_TOLERANCE As Single = 0.0001    ' �����޸���ĳЩ��������һ���߿̶�����
Private Const MAX_CURVECOUNT As Integer = 10            ' �����������Ϊ 10 ����

' ��Ա���� ���� =============================================================================
Private m_ShowGrid As Boolean                           ' �Ƿ���ʾ����
Private m_MovingGrid As MovingGridEnum                  ' ��ֱ�������ƶ���ʽ�����ҽ��������߷Ǿ�ֹʱ��Ҳ�������ƶ�����
Private m_MovingCurve As MovingGridEnum                 ' �����߾�ֹʱ�������ƶ�����ע�⣺���������߾�ֹʱ��Ч������
Private m_HorizontalSplits As Long                      ' ˮƽ����ֳɶ��ٷݣ����������á�
Private m_VerticalSplits As Long                        ' Ǧ������ֳɶ��ٷݣ����������á�
Private m_MaxVertical As Single                         ' Ǧ���������ֵ��
Private m_MinVertical As Single                         ' Ǧ���������ֵ��

Private m_HorizontalGridColor As OLE_COLOR              ' ˮƽ������ ��ɫ��Ĭ��Ϊ����ɫ����RGB(0, 130, 0)
Private m_VerticalGridColor As OLE_COLOR                ' ��ֱ������ ��ɫ
Private m_CurveLineColor() As OLE_COLOR                   ' ������ɫ��Ĭ��Ϊ��ɫ��RGB(0, 130, 0) ����ɫ��
Private m_CurveLineType() As CurveTypeEnum                ' �������ͣ�ʵ�ߣ����ߣ�
Private m_AxesTextColor As OLE_COLOR                    ' ������������ɫ
Private m_ShowAxesText As Boolean                       ' �Ƿ���ʾ���������֣�
Private m_xBarNowTimeFormat As String                   ' X ��ʱ���ʽ��

Private m_CurveCount As Integer                         ' ������������һ���������ϣ�һ�οɻ��������ߣ�
Private m_ShowLegend() As Boolean                         ' �Ƿ���ʾͼ��˵�����ɵ�������ĳһ������ͼ��

' --- ˽�б��� ���� --------------------------------------------------------
Private picGraphHeight As Long                          ' ͼƬ��߶�
Private picGraphWidth As Long                           ' ���
Private GridPosition As Long                            ' ��ֱ������λ�ã�
Private StartPosition As Long                           ' ��ʼ�����ߵ�λ�ã�ͳһ�仯�����ߴ��ҵ��������Ҿ��廭��ʱ���ڿ��ǲ�ͬ������ʼһ����ͼʱ������Ҫ��ʾ��ֵ��Needed to not to display first zero values when starting a new diagram��

Private Type YValues
    yValueArray() As Single                         ' ����Y������ֵ�����顣
End Type
Private yV() As YValues

Private xBarText() As String                            ' ����X����ʾ�����ֵ����飡
'Private StartXBarText As Long                           ' X ��������λ��
Private VerticalGridIndex As Long                       ' ��ֱ��������ţ�X��������λ����š�

Private strLegendLine As String                         ' ͼ���ϵ����ͣ�
Private blnShouldDrawCurve() As Boolean                 ' �Ƿ�Ҫ���������ߣ����ʼ�������¶�������ʱ��һ��Ԫ�س�ʼ��Ϊ True��������ʹ��Ĭ�ϳ�ʼ��ֵ False��

' --- �ṹ�� ����
' === �ƶ�ͼ��˵�� =============================
Private Type PointXY ' ����� X��Y����λ�á�
    x As Single
    y As Single
End Type
Private IsMovingControl() As Boolean   ' ��ʶ���Ƿ������ƶ��ؼ���
Private ptTopLeft() As PointXY         ' �ؼ����Ͻǵ�һ��
Private ptBottomRight() As PointXY     ' �ؼ����Ͻǵ�һ��
Private ptOffset() As PointXY          ' ����ڿؼ��ϰ���ʱ����ǰ������ؼ����Ͻǵ�Ĳ
Private IsControlMoved() As Boolean    ' ĳ��ͼ���Ƿ��Ѿ��ƶ��ˣ�
' ע�⣺ֻҪ��һ���ƶ��ˣ��� UserControl_Resize �Ͳ��ı�ͼ��˵����λ���ˣ�
' === �ƶ�ͼ��˵�� =============================

' --- �¼����� ---------------------------------------------------------------
'Event Declarations:
Event DblClick() 'MappingInfo=picGraph,picGraph,-1,DblClick
Attribute DblClick.VB_Description = "���û���һ�������ϰ��²��ͷ���갴ť���ٴΰ��²��ͷ���갴ťʱ������"
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picGraph,picGraph,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picGraph,picGraph,-1,MouseUp
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picGraph,picGraph,-1,MouseDown


' ######################################################################################################
' ### �������� ###
' ######################################################################################################
' ��ʾ���ڶԻ���
'Public Sub About()
'    MsgBox "����_ʵʱ���߿ؼ� -> ����ʵʱ���ߣ����� Windows ���������CPUʹ�������ߣ���" & vbCrLf & vbCrLf _
'         & "��Ȩ����(C) 2009 ����������" & vbCrLf & vbCrLf _
'         & "��ҳ: http://hewanglan.ys168.com", vbInformation + vbSystemModal, "[����] - ����������"
'End Sub

' ������ݣ����ڻ����ߣ�
Public Sub AddValue(ByVal yValue As Single, Optional ByVal Index As Integer = 1) ', Optional ByVal xBarString As String)
Attribute AddValue.VB_Description = "���Y�������ݣ����ڻ����ߣ�"
    Dim i As Long
    
    ' ���Ӵ�ֱ������Сֵ���� = MinVertical������֧�ָ�����
    yValue = yValue - m_MinVertical
    
    ' �Ƿ�Ҫ���������ߣ�ע�⣺�����Ķ�������Ϊ False����������������ÿ��ֻ�ửһ�����ߣ���
'    For I = 1 To m_CurveCount
'        blnShouldDrawCurve(I) = False
'    Next I
    blnShouldDrawCurve(Index) = True
    ' ����ͼ��˵��
    lblLegend(Index).Visible = m_ShowLegend(Index) And blnShouldDrawCurve(Index)
    
    ' ����������Ԫ��ǰ�ƣ��±��һλ�ĸ�ֵ����һλ�ģ�
    For i = 1 To picGraphWidth - 1
        yV(Index).yValueArray(i - 1) = yV(Index).yValueArray(i)
    Next i

    ' ����ֵ��ӵ��������һ��Ԫ�أ�ע�⣺������ I-1����Ϊ����ѭ�������� I �����Զ� +1 ��
    yV(Index).yValueArray(i - 1) = picGraphHeight - ((yValue / (m_MaxVertical - m_MinVertical)) * picGraphHeight)
    
    ' �Ӹ��жϣ�ֻ�ڵ�һ��ʱ����һ�Σ���������ͼʱÿ�����߶���̫�ԣ�X��ʱ��Ҳ���ԡ�
    ' ��ʼ�����ߵ�λ�ã�
    If Index = 1 Then
        If StartPosition >= 1 Then StartPosition = StartPosition - 1
    
        ' ��ֱ������λ���ƶ���
        If m_MovingGrid = [Right to Left] Then      ' ���ҵ���
            GridPosition = GridPosition - 1
        ElseIf m_MovingGrid = [Left to Right] Then  ' �����ң�
            GridPosition = GridPosition + 1
        End If
        
        ' === X ������ʾ������ ==================================================================
        If Len(xBarText(0)) = 0 Then xBarText(0) = Format$(Now, m_xBarNowTimeFormat)
    End If
'    StartXBarText = StartXBarText - 1           ' X ���������ƶ�
'    Dim b As Boolean
'    For I = m_HorizontalSplits To 0 Step -1
'        If Len(xBarText(I)) > 0 Then
'            b = True
'        Else
'            b = False
'        End If
'        Debug.Print I & " == " & xBarText(I)
'    Next I
'    If Not b Then
'        ' ����������Ԫ��ǰ��
'        For I = 1 To m_HorizontalSplits
'            xBarText(I - 1) = xBarText(I)
'        Next I
'        xBarText(I - 1) = xBarString
'    End If
End Sub
' ���ͼ��˵������ �� A������Ҫ�޸�ʱ�����ô˺������ɣ�
Public Sub FixLegend(ByVal strText As String, Optional ByVal Index As Integer = 1)
Attribute FixLegend.VB_Description = "���ͼ��˵������ �� A������Ҫ�޸�ʱ�����ô˺������ɣ�"
    lblLegend(Index).Caption = strLegendLine & strText
    lblLegend(Index).ForeColor = m_CurveLineColor(Index)
    
    ' ��ʼ�� Shape ͼ���ƶ�Ч����
    Call InitLegendShape(Index)
End Sub

' �Ȼ�ˮƽ����ֱ���������ߣ��ٻ����ߣ�
Public Function DrawGridCurve()
Attribute DrawGridCurve.VB_Description = "���ĺ���������ˮƽ����ֱ���������ߣ��������ߣ�����X��Y�������֡�"
    Dim x As Single
    Dim y As Single
    Dim i As Long
    ' �� KK ������
    Dim KK As Integer
    
    ' �����ͼƬ��
    picGraph.Cls

    ' 1����������
    If m_ShowGrid Then
        ' 1.1 ˮƽ������
        For y = 0 To (picGraphHeight - 1) Step ((picGraphHeight - 1) / (m_VerticalSplits)) - LAST_LINE_TOLERANCE
            picGraph.Line (0, y)-(picGraphWidth, y), m_HorizontalGridColor
        Next y
        ' 1.2 ��ֱ�����ߣ�ע�⣺��3��������ƶ������ƶ������ҵ��󣿴����ң���
        If m_MovingGrid = [Not Moving] Then ' ����Ҫ�ƶ�����̬�����ߡ�
            For x = 0 To (picGraphWidth - 1) Step ((picGraphWidth - 1) / (m_HorizontalSplits)) - LAST_LINE_TOLERANCE
                picGraph.Line (x, 0)-(x, picGraphHeight), m_VerticalGridColor
            Next x
        Else ' ���ҵ���' �����ң� һ���Ĵ��룬ֻ�� AddValue ������ GridPosition �仯���Ʋ�һ����GridPosition = GridPosition - 1 �� + 1
            For x = GridPosition To (picGraphWidth - 1) Step ((picGraphWidth - 1) / (m_HorizontalSplits)) - LAST_LINE_TOLERANCE
                picGraph.Line (x, 0)-(x, picGraphHeight), m_VerticalGridColor
            Next x
        End If
    End If
    ' 1.3 ����һ����ֱ�����߲���ʱ��������λ�ã���ע�⣺��2��������ƶ������ҵ��󣿴����ң���
    If m_MovingGrid = [Right to Left] Then      ' ���ҵ���
        If GridPosition <= -Int((picGraphWidth - 1) / m_HorizontalSplits) Then
            GridPosition = 0
            ' ���ʱ�䣬ע�⣺�� +1 �� 0 �ڳ�ʼ��ʱ��ӣ�
            VerticalGridIndex = VerticalGridIndex + 1
            xBarText(VerticalGridIndex) = Format$(Now, m_xBarNowTimeFormat)
        End If
    ElseIf m_MovingGrid = [Left to Right] Then  ' �����ң�
        If GridPosition >= Int((picGraphWidth - 1) / m_HorizontalSplits) Then
            GridPosition = 0
            ' ���ʱ�䣬ע�⣺�� +1 �� 0 �ڳ�ʼ��ʱ��ӣ�
            VerticalGridIndex = VerticalGridIndex + 1
            xBarText(VerticalGridIndex) = Format$(Now, m_xBarNowTimeFormat)
        End If
    End If
    ' ��������ֱ�߶���ʾ�����ֺ����㣬��ͷ������ʾ������ˣ�ÿ����ʾ���ͻ��ִӵ���������߿�ʼ���������ֱ������
    If VerticalGridIndex >= m_HorizontalSplits - 1 Then VerticalGridIndex = 0: xBarText(0) = Format$(Now, m_xBarNowTimeFormat)
'    If StartXBarText <= -(picGraphWidth - 1) Then StartXBarText = 0

    ' 2��������
    ' Draw line diagram only if there are 2 or more values defined
    If m_MovingGrid = [Right to Left] Then      ' ���ҵ���
        ' ----------------------------------------------------------------------------------------------
        If StartPosition <= picGraphWidth - 1 Then
            For KK = 1 To m_CurveCount
                If blnShouldDrawCurve(KK) Then ' �Ƿ�Ҫ���������ߣ�
                    If m_CurveLineType(KK) = [Solid] Then
                        For i = StartPosition + 1 To picGraphWidth - 2
                            picGraph.Line (i, yV(KK).yValueArray(i))-(i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                        Next i
                    Else
                        For i = StartPosition + 1 To picGraphWidth - 2
                            picGraph.PSet (i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                        Next i
                    End If
                End If
            Next KK
        End If
        ' ----------------------------------------------------------------------------------------------
    ElseIf m_MovingGrid = [Left to Right] Then  ' �����ң���������ȣ� I ǰ���� picGraphWidth - 1 -
        ' ----------------------------------------------------------------------------------------------
        If StartPosition <= picGraphWidth - 1 Then
            For KK = 1 To m_CurveCount
                If blnShouldDrawCurve(KK) Then ' �Ƿ�Ҫ���������ߣ�
                    If m_CurveLineType(KK) = [Solid] Then
                        For i = StartPosition + 1 To picGraphWidth - 2
                            picGraph.Line (picGraphWidth - 1 - i, yV(KK).yValueArray(i))-(picGraphWidth - 1 - i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                        Next i
                    Else
                        For i = StartPosition + 1 To picGraphWidth - 2
                            picGraph.PSet (picGraphWidth - 1 - i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                        Next i
                    End If
                End If
            Next KK
        End If
        ' ----------------------------------------------------------------------------------------------
    Else ' �����߾�ֹʱ�������ƶ�����
        ' **********************************************************************************************
        If m_MovingCurve = [Right to Left] Or m_MovingCurve = [Not Moving] Then   ' ���ҵ������߲��ܾ�ֹ��Ĭ�ϴ��ҵ����ƶ���
            If StartPosition <= picGraphWidth - 1 Then
                For KK = 1 To m_CurveCount
                    If blnShouldDrawCurve(KK) Then ' �Ƿ�Ҫ���������ߣ�
                        If m_CurveLineType(KK) = [Solid] Then
                            For i = StartPosition + 1 To picGraphWidth - 2
                                picGraph.Line (i, yV(KK).yValueArray(i))-(i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                            Next i
                        Else
                            For i = StartPosition + 1 To picGraphWidth - 2
                                picGraph.PSet (i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                            Next i
                        End If
                    End If
                Next KK
            End If
        Else  ' �����ң�
            If StartPosition <= picGraphWidth - 1 Then
                For KK = 1 To m_CurveCount
                    If blnShouldDrawCurve(KK) Then ' �Ƿ�Ҫ���������ߣ�
                        If m_CurveLineType(KK) = [Solid] Then
                            For i = StartPosition + 1 To picGraphWidth - 2
                                picGraph.Line (picGraphWidth - 1 - i, yV(KK).yValueArray(i))-(picGraphWidth - 1 - i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                            Next i
                        Else
                            For i = StartPosition + 1 To picGraphWidth - 2
                                picGraph.PSet (picGraphWidth - 1 - i + 1, yV(KK).yValueArray(i + 1)), m_CurveLineColor(KK)
                            Next i
                        End If
                    End If
                Next KK
            End If
        End If
        ' **********************************************************************************************
    End If
    
    ' 3������ֵ������ʾ
    If m_ShowAxesText Then
        With picGraph
            ' 3.1 Y �ᡣ���̶ȣ�
            .CurrentX = 3
            .CurrentY = 0
            picGraph.Print Format$(m_MaxVertical, "0.0#")
            ' �м�̶�
            For i = 1 To m_VerticalSplits - 1
                .CurrentX = 3
                .CurrentY = Int(picGraphHeight / m_VerticalSplits * i - picGraph.TextHeight(CStr(i)) / 2)
                picGraph.Print Format$(m_MaxVertical - ((m_MaxVertical - m_MinVertical) / m_VerticalSplits * i), "0.0#")
            Next i
            ' ��С�̶�
            .CurrentX = 3
            .CurrentY = picGraphHeight - picGraph.TextHeight(CStr(i))
            picGraph.Print m_MinVertical
            
            ' 3.2 X �����֡�
'            X = Abs(StartXBarText)
            If m_MovingGrid = [Left to Right] Then
                For i = 0 To VerticalGridIndex
                    ' X λ�������ȷ���������˽�1Сʱ��������
                    picGraph.CurrentX = GridPosition + i * (((picGraphWidth - 1) / (m_HorizontalSplits))) - picGraph.TextWidth(xBarText(VerticalGridIndex - i)) / 2
                    picGraph.CurrentY = picGraphHeight - picGraph.TextHeight(CStr(Int(x)))
                    picGraph.Print xBarText(VerticalGridIndex - i)
                Next i
            ElseIf m_MovingGrid = [Right to Left] Then
                For i = 0 To VerticalGridIndex
                    ' X λ�������ȷ���������˽�1Сʱ��������
                    picGraph.CurrentX = picGraphWidth + GridPosition - i * (((picGraphWidth - 1) / (m_HorizontalSplits))) - picGraph.TextWidth(xBarText(VerticalGridIndex - i)) / 2
                    picGraph.CurrentY = picGraphHeight - picGraph.TextHeight(CStr(Int(x)))
                    picGraph.Print xBarText(VerticalGridIndex - i)
                Next i
            End If
        End With
    End If
End Function

' ������ߡ�X ����ʱ��ȡ�
Public Sub ClearAll()
Attribute ClearAll.VB_Description = "������ߡ�X ����ʱ��ȡ�"
    ' ˽�б��� ��ʼ����ע�⣺Ҫ�� UserControl_Resize ����дһ�飿�����ʼ�����ܲ��ԣ�������
    picGraphHeight = picGraph.ScaleHeight   ' ͼƬ��߶�
    picGraphWidth = picGraph.ScaleWidth     ' ���
    GridPosition = 0                        ' ��ֱ������λ��
    StartPosition = picGraphWidth - 1       ' ��ʼ�����ߵ�λ�ã�
    VerticalGridIndex = 0                   ' ��ֱ��������ţ�X��������λ����š�
    
    ' �Ƿ�Ҫ���������ߣ�
    ReDim blnShouldDrawCurve(1 To m_CurveCount) As Boolean
    blnShouldDrawCurve(1) = True
    
    ' Allocate array to hold all diagram values (value per pixel)
    Dim i As Integer
    For i = 1 To m_CurveCount
        ReDim yV(i).yValueArray(picGraphWidth - 1)
    Next i
    
    ReDim xBarText(m_HorizontalSplits)
        
End Sub
' ######################################################################################################
' ### �������� ###
' ######################################################################################################



' ######################################################################################################
' ### ��������  ###
' ######################################################################################################
' ����: ͼƬ���Ƿ������Զ��ػ�
Public Property Get AutoRedraw() As Boolean ' �������ֵ
Attribute AutoRedraw.VB_Description = "����/���� �ؼ��Ƿ������Զ��ػ棡"
    AutoRedraw = picGraph.AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal New_Value As Boolean) ' ��������
    picGraph.AutoRedraw = New_Value
    PropertyChanged "AutoRedraw"
End Property

' ����: ͼƬ�򱳾���ɫ
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���� �ؼ�������ɫ��"
    BackColor = picGraph.BackColor
End Property
Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    picGraph.BackColor = New_Value
    ' ���߿���ɫ������ɫȡ��ɫ������
    Dim i As Integer
    For i = 1 To m_CurveCount
        ShapeLegend(i).BorderColor = ColorInverted(picGraph.BackColor)
    Next i
    PropertyChanged "BackColor"
End Property

' ����: �ؼ��߿�����
Public Property Get BorderStyle() As BoderStyleEnum
Attribute BorderStyle.VB_Description = "����/���� �ؼ��߿����͡�"
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_Value As BoderStyleEnum)
    UserControl.BorderStyle = New_Value
    PropertyChanged "BorderStyle"
End Property

' ����: �Ƿ���ʾ����
Public Property Get ShowGrid() As Boolean
Attribute ShowGrid.VB_Description = "����/���� �Ƿ���ʾˮƽ����ֱ�����ߡ�"
    ShowGrid = m_ShowGrid
End Property
Public Property Let ShowGrid(ByVal New_Value As Boolean)
    m_ShowGrid = New_Value
    PropertyChanged "ShowGrid"
End Property

' ����: ��ֱ�������ƶ���ʽ��Ҳ�������ƶ�����
Public Property Get MovingGrid() As MovingGridEnum
Attribute MovingGrid.VB_Description = "����/���� ��ֱ�������ƶ���ʽ��Ҳ�������ƶ�����"
    MovingGrid = m_MovingGrid
End Property
Public Property Let MovingGrid(ByVal New_Value As MovingGridEnum)
    m_MovingGrid = New_Value
    
    ' ���ú��� ��ʼ��������
    Call ClearAll
'    ' ��ʼ�����ߵ�λ�ã�
'    ' ��ֱ��������ţ�X��������λ����š�
'    VerticalGridIndex = 0
'    StartPosition = picGraphWidth - 1
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)

    PropertyChanged "MovingGrid"
End Property

' ����: �����߾�ֹʱ�������ƶ�����ע�⣺���������߾�ֹʱ��Ч������
Public Property Get MovingCurve() As MovingGridEnum
Attribute MovingCurve.VB_Description = "����/���� �����߾�ֹʱ�������ƶ�����ע�⣺���������߾�ֹʱ��Ч������"
    MovingCurve = m_MovingCurve
End Property
Public Property Let MovingCurve(ByVal New_Value As MovingGridEnum)
    m_MovingCurve = New_Value
    
    ' ���ú��� ��ʼ��������
    Call ClearAll
'    ' ��ʼ�����ߵ�λ�ã�
'    StartPosition = picGraphWidth - 1
'    ' ��ֱ��������ţ�X��������λ����š�
'    VerticalGridIndex = 0
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)
    
    PropertyChanged "MovingCurve"
End Property

' ����: ˮƽ�ֳɶ��ٷݣ����������á�
Public Property Get HorizontalSplits() As Long
Attribute HorizontalSplits.VB_Description = "����/���� ˮƽ�ֳɶ��ٷݣ����������á�"
    HorizontalSplits = m_HorizontalSplits
End Property
Public Property Let HorizontalSplits(ByVal New_Value As Long)
    ' ֵ����<=0
    If New_Value <= 0 Then New_Value = 1
    m_HorizontalSplits = New_Value
    PropertyChanged "HorizontalSplits"
End Property

' ����: Ǧ������ֳɶ��ٷݣ����������á�
Public Property Get VerticalSplits() As Long
Attribute VerticalSplits.VB_Description = "����/���� Ǧ������ֳɶ��ٷݣ����������á�"
    VerticalSplits = m_VerticalSplits
End Property
Public Property Let VerticalSplits(ByVal New_Value As Long)
    ' ֵ����<=0
    If New_Value <= 0 Then New_Value = 1
    m_VerticalSplits = New_Value
    PropertyChanged "VerticalSplits"
End Property

' ����: Ǧ���������ֵ��
Public Property Get MaxVertical() As Single
Attribute MaxVertical.VB_Description = "����/���� Ǧ���������ֵ��"
    MaxVertical = m_MaxVertical
End Property
Public Property Let MaxVertical(ByVal New_Value As Single)
    ' ���ֵ����<=0
    If New_Value <= 0 Then New_Value = 1
    m_MaxVertical = New_Value
    PropertyChanged "MaxVertical"
End Property

' ����: Ǧ��������Сֵ��
Public Property Get MinVertical() As Single
Attribute MinVertical.VB_Description = "����/���� Ǧ��������Сֵ��"
    MinVertical = m_MinVertical
End Property
Public Property Let MinVertical(ByVal New_Value As Single)
    ' ��Сֵ����>=���ֵ
    If New_Value >= m_MaxVertical Then New_Value = m_MaxVertical - 1
    m_MinVertical = New_Value
    PropertyChanged "MinVertical"
End Property

' ����: ˮƽ������ ��ɫ
Public Property Get HorizontalGridColor() As OLE_COLOR
Attribute HorizontalGridColor.VB_Description = "����/���� ˮƽ��������ɫ��"
    HorizontalGridColor = m_HorizontalGridColor
End Property
Public Property Let HorizontalGridColor(ByVal New_Value As OLE_COLOR)
    m_HorizontalGridColor = New_Value
    PropertyChanged "HorizontalGridColor"
End Property

' ����: ��ֱ������ ��ɫ
Public Property Get VerticalGridColor() As OLE_COLOR
Attribute VerticalGridColor.VB_Description = "����/���� ��ֱ��������ɫ��"
    VerticalGridColor = m_VerticalGridColor
End Property
Public Property Let VerticalGridColor(ByVal New_Value As OLE_COLOR)
    m_VerticalGridColor = New_Value
    PropertyChanged "VerticalGridColor"
End Property

' ����: ������ɫ
Public Property Get CurveLineColor(Optional ByVal Index As Integer = 1) As OLE_COLOR
Attribute CurveLineColor.VB_Description = "����/���� ������ɫ��"
    CurveLineColor = m_CurveLineColor(Index)
End Property
Public Property Let CurveLineColor(Optional ByVal Index As Integer = 1, ByVal New_Value As OLE_COLOR)
    m_CurveLineColor(Index) = New_Value
    ' ����ͼ����ʾ
    lblLegend(Index).ForeColor = m_CurveLineColor(Index)
    PropertyChanged "CurveLineColor"
End Property

' ����: �������ͣ�ʵ�ߣ����ߣ�
Public Property Get CurveLineType(Optional ByVal Index As Integer = 1) As CurveTypeEnum
Attribute CurveLineType.VB_Description = "����/���� �������ͣ�ʵ�ߣ����ߣ�"
    CurveLineType = m_CurveLineType(Index)
End Property
Public Property Let CurveLineType(Optional ByVal Index As Integer = 1, ByVal New_Value As CurveTypeEnum)
    m_CurveLineType(Index) = New_Value
    ' ����ͼ����ʾ
    On Error Resume Next
    strLegendLine = IIf(m_CurveLineType(Index) = [Solid], "�� ", "�� ")
    FixLegend Right$(lblLegend(Index).Caption, Len(lblLegend(Index).Caption) - 2), Index
    
    PropertyChanged "CurveLineType"
End Property

' ����: ������������ɫ
Public Property Get AxesTextColor() As OLE_COLOR
Attribute AxesTextColor.VB_Description = "����/���� ������������ɫ��"
    AxesTextColor = m_AxesTextColor
End Property
Public Property Let AxesTextColor(ByVal New_Value As OLE_COLOR)
    m_AxesTextColor = New_Value
    picGraph.ForeColor = m_AxesTextColor
    PropertyChanged "AxesTextColor"
End Property

' ����: �Ƿ���ʾ����������
Public Property Get ShowAxesText() As Boolean
Attribute ShowAxesText.VB_Description = "����/���� �Ƿ���ʾ���������֣�"
    ShowAxesText = m_ShowAxesText
End Property
Public Property Let ShowAxesText(ByVal New_Value As Boolean)
    m_ShowAxesText = New_Value
    PropertyChanged "ShowAxesText"
End Property

' ����: X ���������֣�ʱ�䣩��ʽ��
Public Property Get xBarNowTimeFormat() As String
Attribute xBarNowTimeFormat.VB_Description = "����/���� X ���������֣�ʱ�䣩��ʽ��"
    xBarNowTimeFormat = m_xBarNowTimeFormat
End Property
Public Property Let xBarNowTimeFormat(ByVal New_Value As String)
    m_xBarNowTimeFormat = New_Value
    PropertyChanged "xBarNowTimeFormat"
End Property

'' ����: ��������������
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picGraph,picGraph,-1,Font
Public Property Get AxesFont() As Font
Attribute AxesFont.VB_Description = "����/���� X �� Y �������������壡"
    Set AxesFont = picGraph.Font
End Property

Public Property Set AxesFont(ByVal New_AxesFont As Font)
    Set picGraph.Font = New_AxesFont
    PropertyChanged "AxesFont"
End Property

' ����: ������������һ���������ϣ�һ�οɻ��������ߣ�
Public Property Get CurveCount() As Integer
Attribute CurveCount.VB_Description = "����/���� ������������һ���������ϣ�һ�οɻ��������ߣ�"
    CurveCount = m_CurveCount
End Property
Public Property Let CurveCount(ByVal New_Value As Integer)
    ' �������� 1��MAX_CURVECOUNT ��
    If New_Value < 1 Then New_Value = 1
    If New_Value > MAX_CURVECOUNT Then New_Value = MAX_CURVECOUNT
    m_CurveCount = New_Value
    
    ' ���¶������飺������ɫ������
    ReDim Preserve m_CurveLineColor(1 To m_CurveCount) As OLE_COLOR      ' ������ɫ��Ĭ��Ϊ��ɫ��RGB(0, 130, 0) ����ɫ��
    ReDim Preserve m_CurveLineType(1 To m_CurveCount) As CurveTypeEnum   ' �������ͣ�ʵ�ߣ����ߣ�
    ReDim m_ShowLegend(1 To m_CurveCount) As Boolean                     ' �Ƿ���ʾͼ��˵�����ɵ�������ĳһ������ͼ��
    
    ReDim blnShouldDrawCurve(1 To m_CurveCount) As Boolean              ' �Ƿ�Ҫ���������ߣ�
    blnShouldDrawCurve(1) = True
    ReDim yV(1 To m_CurveCount) As YValues                              ' ����Y������ֵ�����顣
    
    ' === �ƶ�ͼ��˵�� =============================
    ReDim IsMovingControl(1 To m_CurveCount) As Boolean   ' ��ʶ���Ƿ��ƶ��ؼ���
    ReDim ptTopLeft(1 To m_CurveCount) As PointXY         ' �ؼ����Ͻǵ�һ��
    ReDim ptBottomRight(1 To m_CurveCount) As PointXY     ' �ؼ����Ͻǵ�һ��
    ReDim ptOffset(1 To m_CurveCount) As PointXY          ' ����ڿؼ��ϰ���ʱ����ǰ������ؼ����Ͻǵ�Ĳ
    ReDim IsControlMoved(1 To m_CurveCount) As Boolean    ' ĳ��ͼ���Ƿ��Ѿ��ƶ��ˣ�
    ' === �ƶ�ͼ��˵�� =============================
    
    Dim i As Integer
    For i = 1 To m_CurveCount
        ReDim yV(i).yValueArray(picGraphWidth - 1)
        m_CurveLineColor(i) = vbGreen              ' ������ɫ
        m_CurveLineType(i) = [Solid]               ' �������ͣ�ʵ�ߣ����ߣ�
        m_ShowLegend(i) = True
    Next i
    ' ��ʼ�� ͼ���ı���ǩ ����
    Call InitLegendLabel
    
    PropertyChanged "CurveCount"
End Property

' ����: �Ƿ���ʾͼ��˵����ע�⣺ And
Public Property Get ShowLegend(Optional ByVal Index As Integer = 1) As Boolean
Attribute ShowLegend.VB_Description = "����/���� �Ƿ���ʾĳ�����ߵ�ͼ��˵��������������û����AddValue�������ʱ����������Ч��"
    ShowLegend = m_ShowLegend(Index) And blnShouldDrawCurve(Index)
End Property
Public Property Let ShowLegend(Optional ByVal Index As Integer = 1, ByVal New_Value As Boolean)
    m_ShowLegend(Index) = New_Value

    ' ����ͼ��˵��
    lblLegend(Index).Visible = m_ShowLegend(Index) And blnShouldDrawCurve(Index)
    PropertyChanged "ShowLegend"
End Property

' ����: ͼ��˵����������
Public Property Get LegendFont(Optional ByVal Index As Integer = 1) As Font
Attribute LegendFont.VB_Description = "����/���� ͼ��˵����������"
    Set LegendFont = lblLegend(Index).Font
End Property

Public Property Set LegendFont(Optional ByVal Index As Integer = 1, ByVal New_LegendFont As Font)
    Set lblLegend(Index).Font = New_LegendFont
    PropertyChanged "LegendFont"
End Property
' ######################################################################################################
' ### ��������  ###
' ######################################################################################################


' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' --- �û��ؼ������¼� ---------------------------------------------------------------------------------
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' SSS �ڿؼ����������ʱ��������ʼ���������ؼ���ʵ������ʱ��ͻ���ã�Ҳ���ǿؼ��ڼ��ص�ʱ��ᱻִ��
Private Sub UserControl_Initialize()
    ' ��Ա���� ��ʼ��
    ' ͼƬ��������Ϊ��ͼ�̶ȣ��ƺ�û�б�Ҫ�������ؿ��Լ������滭������ѭ���Ĵ�������
    ' ���Ի���VBĬ�ϵģ��������Ļ������� picGraphHeight��picGraphWidth �ĸ�ֵ�仯Ϊ ScaleHeight, ScaleWidth ����
    picGraph.ScaleMode = vbPixels ' vbTwips
    picGraph.AutoRedraw = True              ' ͼƬ�������Զ��ػ�
    picGraph.BackColor = vbBlack            ' ͼƬ�򱳾�ɫ
    picGraph.BorderStyle = [None]           ' ע�⣺ͼƬ��߿��������� 0���ؼ��߿������� UserControl �ģ�
    UserControl.BorderStyle = [None]
    
    m_ShowGrid = True                       ' �Ƿ���ʾ����
    m_MovingGrid = [Right to Left]          ' ��ֱ�������ƶ���ʽ�����ҽ��������߷Ǿ�ֹʱ��Ҳ�������ƶ�����
    m_MovingCurve = [Right to Left]         ' �����߾�ֹʱ�������ƶ�����ע�⣺���������߾�ֹʱ��Ч������
    m_HorizontalSplits = 9                  ' ˮƽ����ֳɶ��ٷݣ����������á�
    m_VerticalSplits = 9                    ' Ǧ������ֳɶ��ٷݣ����������á�
    m_MaxVertical = 9                       ' Ǧ���������ֵ��
    m_MinVertical = 0                       ' Ǧ��������Сֵ��
    m_HorizontalGridColor = RGB(0, 130, 0)  ' ˮƽ������ ��ɫ
    m_VerticalGridColor = RGB(0, 130, 0)    ' ��ֱ������ ��ɫ
    
    m_CurveCount = 1                        ' ������������һ���������ϣ�һ�οɻ��������ߣ�
    
    
    ' ���¶������飺������ɫ������
    ReDim m_CurveLineColor(1 To m_CurveCount) As OLE_COLOR      ' ������ɫ��Ĭ��Ϊ��ɫ��RGB(0, 130, 0) ����ɫ��
    ReDim m_CurveLineType(1 To m_CurveCount) As CurveTypeEnum   ' �������ͣ�ʵ�ߣ����ߣ�
    ReDim m_ShowLegend(1 To m_CurveCount) As Boolean            ' �Ƿ���ʾͼ��˵�����ɵ�������ĳһ������ͼ��
    
    ReDim blnShouldDrawCurve(1 To m_CurveCount) As Boolean      ' �Ƿ�Ҫ���������ߣ�
    blnShouldDrawCurve(1) = True
    ReDim yV(1 To m_CurveCount) As YValues                      ' ����Y������ֵ�����顣
    
    ' === �ƶ�ͼ��˵�� =============================
    ReDim IsMovingControl(1 To m_CurveCount) As Boolean   ' ��ʶ���Ƿ��ƶ��ؼ���
    ReDim ptTopLeft(1 To m_CurveCount) As PointXY         ' �ؼ����Ͻǵ�һ��
    ReDim ptBottomRight(1 To m_CurveCount) As PointXY     ' �ؼ����Ͻǵ�һ��
    ReDim ptOffset(1 To m_CurveCount) As PointXY          ' ����ڿؼ��ϰ���ʱ����ǰ������ؼ����Ͻǵ�Ĳ
    ReDim IsControlMoved(1 To m_CurveCount) As Boolean    ' ĳ��ͼ���Ƿ��Ѿ��ƶ��ˣ�
    ' === �ƶ�ͼ��˵�� =============================
    
    Dim i As Integer
    For i = 1 To m_CurveCount
        m_CurveLineColor(i) = vbGreen              ' ������ɫ
        m_CurveLineType(i) = [Solid]               ' �������ͣ�ʵ�ߣ����ߣ�
        m_ShowLegend(i) = True                     ' �Ƿ���ʾͼ��˵�����ɵ�������ĳһ������ͼ��
    Next i
    ' ��ʼ�� ͼ���ı���ǩ ����
    Call InitLegendLabel
    
    m_AxesTextColor = vbWhite               ' ������������ɫ
    m_ShowAxesText = True                   ' �Ƿ���ʾ����������
    m_xBarNowTimeFormat = "hh:mm:ss"        ' X ��ʱ���ʽ��
    
    ' ���ú��� ��ʼ��������
    Call ClearAll
'    ' ˽�б��� ��ʼ����ע�⣺Ҫ�� UserControl_Resize ����дһ�飿�����ʼ�����ܲ��ԣ�������
'    picGraphHeight = picGraph.ScaleHeight   ' ͼƬ��߶�
'    picGraphWidth = picGraph.ScaleWidth     ' ���
'    GridPosition = 0                        ' ��ֱ������λ��
'    StartPosition = picGraphWidth - 1       ' ��ʼ�����ߵ�λ�ã�
'    VerticalGridIndex = 0                   ' ��ֱ��������ţ�X��������λ����š�
'
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)

End Sub

' SSS �ؼ�ʵ���ڶ��μ��Ժ����´���ʱ�������������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        AutoRedraw = .ReadProperty("AutoRedraw", True)
        BackColor = .ReadProperty("BackColor", vbBlack)
        BorderStyle = .ReadProperty("BorderStyle", [None])
        ShowGrid = .ReadProperty("ShowGrid", True)
        MovingGrid = .ReadProperty("MovingGrid", [Right to Left])
        MovingCurve = .ReadProperty("MovingCurve", [Right to Left])
        HorizontalSplits = .ReadProperty("HorizontalSplits", 9)
        VerticalSplits = .ReadProperty("VerticalSplits", 9)
        MaxVertical = .ReadProperty("MaxVertical", 9)
        MinVertical = .ReadProperty("MinVertical", 0)
        HorizontalGridColor = .ReadProperty("HorizontalGridColor", RGB(0, 130, 0))
        VerticalGridColor = .ReadProperty("VerticalGridColor", RGB(0, 130, 0))
        
        Dim i As Integer
        For i = 1 To m_CurveCount
            CurveLineColor(i) = .ReadProperty("CurveLineColor", vbGreen)
            CurveLineType(i) = .ReadProperty("CurveLineType", [Solid])
            Set lblLegend(i).Font = PropBag.ReadProperty("LegendFont", Ambient.Font)
            ShowLegend(i) = .ReadProperty("ShowLegend", True)
        Next i
        
        AxesTextColor = .ReadProperty("AxesTextColor", vbWhite)
        ShowAxesText = .ReadProperty("ShowAxesText", True)
        xBarNowTimeFormat = .ReadProperty("xBarNowTimeFormat", "hh:mm:ss")
        CurveCount = .ReadProperty("CurveCount", 1)
    End With
    Set picGraph.Font = PropBag.ReadProperty("AxesFont", Ambient.Font)
End Sub
' SSS �ؼ����ʱ��ʵ��������ʱ����������ʱ��������ؼ������ʱ���õ�����ֵ
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "AutoRedraw", picGraph.AutoRedraw, True
        .WriteProperty "BackColor", picGraph.BackColor, vbBlack
        .WriteProperty "BorderStyle", UserControl.BorderStyle, [None]
        .WriteProperty "ShowGrid", m_ShowGrid, True
        .WriteProperty "MovingGrid", m_MovingGrid, [Right to Left]
        .WriteProperty "MovingCurve", m_MovingCurve, [Right to Left]
        .WriteProperty "HorizontalSplits", m_HorizontalSplits, 9
        .WriteProperty "VerticalSplits", m_VerticalSplits, 9
        .WriteProperty "MaxVertical", m_MaxVertical, 9
        .WriteProperty "MinVertical", m_MinVertical, 0
        .WriteProperty "HorizontalGridColor", m_HorizontalGridColor, RGB(0, 130, 0)
        .WriteProperty "VerticalGridColor", m_VerticalGridColor, RGB(0, 130, 0)
        Dim i As Integer
        For i = 1 To m_CurveCount
            .WriteProperty "CurveLineColor", m_CurveLineColor(i), vbGreen
            .WriteProperty "CurveLineType", m_CurveLineType(i), [Solid]
            Call PropBag.WriteProperty("LegendFont", lblLegend(i).Font, Ambient.Font)
            .WriteProperty "ShowLegend", m_ShowLegend(i), True
        Next i
                
        .WriteProperty "AxesTextColor", m_AxesTextColor, vbWhite
        .WriteProperty "ShowAxesText", m_ShowAxesText, True
        .WriteProperty "xBarNowTimeFormat", m_xBarNowTimeFormat, "hh:mm:ss"
        .WriteProperty "CurveCount", m_CurveCount, 1
    End With
    Call PropBag.WriteProperty("AxesFont", picGraph.Font, Ambient.Font)
End Sub
' SSS �ؼ�ʵ��������֮ǰִ��
Private Sub UserControl_Terminate()
    Erase yV   ' ����Y������ֵ�����顣
    Erase xBarText      ' ����X����ʾ�����ֵ����飡
End Sub
' SSS ���û��ؼ���С�ı�ʱ�����ı�ؼ�ʵ����Сʱ����
Private Sub UserControl_Resize()
    ' ͼƬ��ؼ�λ�úʹ�С��ע�⣺��͸߼�ȥһ�������������һ���������޷���ʾ��
    picGraph.Move 0, 0, UserControl.Width - 52, UserControl.Height - 52
    ' ���� ͼ���ı���ǩ ����' ĳ��ͼ���Ƿ��Ѿ��ƶ��ˣ�
    Dim i As Integer
    If Not IsControlMoved(1) Then lblLegend(1).Left = (picGraph.ScaleWidth - lblLegend(1).Width) / 2
    ' ��ʼ�� Shape ͼ���ƶ�Ч����
    Call InitLegendShape(1)
    For i = 2 To m_CurveCount
        If Not IsControlMoved(i) Then lblLegend(i).Move lblLegend(1).Left, lblLegend(i - 1).Top + lblLegend(i).Height
        ' ��ʼ�� Shape ͼ���ƶ�Ч����
        Call InitLegendShape(i)
    Next i
    
    ' ���ú��� ��ʼ��������
    Call ClearAll
'    picGraphHeight = picGraph.ScaleHeight        ' ͼƬ��߶�
'    picGraphWidth = picGraph.ScaleWidth          ' ���
'    GridPosition = 0                             ' ��ֱ������λ��
'    StartPosition = picGraphWidth - 1            ' ��ʼ�����ߵ�λ�ã�
'    VerticalGridIndex = 0                        ' ��ֱ��������ţ�X��������λ����š�
'
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)

End Sub
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' --- �û��ؼ������¼� ---------------------------------------------------------------------------------
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


' --- �ؼ��ڲ��ķ��������ڷ�����˽�з�����--------------------------------------------------------------
' ��ʼ�� ͼ���ı���ǩ ����
Private Sub InitLegendLabel()
    ' ͼ���ı���ǩ��һ����λ�õ�������д�á�
    lblLegend(1).Caption = "�� 1"
    lblLegend(1).Move (picGraph.ScaleWidth - lblLegend(1).Width) / 2, 6
    lblLegend(1).Visible = m_ShowLegend(1) And blnShouldDrawCurve(1)
    ' ��ʼ�� Shape ͼ���ƶ�Ч����
    Call InitLegendShape(1)
        
    ' �����ı���ǩ����
    Dim i As Integer
    For i = 2 To m_CurveCount
        ' ����ͼ���ı���ǩ
        On Error Resume Next
        Load lblLegend(i)
        Load ShapeLegend(i)
        lblLegend(i).ForeColor = m_CurveLineColor(i)
        lblLegend(i).Caption = strLegendLine & CStr(i)
        lblLegend(i).Move lblLegend(1).Left, lblLegend(i - 1).Top + lblLegend(i).Height
        lblLegend(i).Visible = m_ShowLegend(i) And blnShouldDrawCurve(i)
        
        ' ��ʼ�� Shape ͼ���ƶ�Ч����
        Call InitLegendShape(i)
    Next i
End Sub
' ��ʼ�� Shape ͼ���ƶ�Ч����
Private Sub InitLegendShape(Optional ByVal Index As Integer)
    ' ��ʼ�������Ͻǵ㡢���½ǵ� ����������ע�⣺ShapeLegend(index).Move
    ShapeLegend(Index).Move lblLegend(Index).Left, lblLegend(Index).Top, lblLegend(Index).Width, lblLegend(Index).Height
    ptTopLeft(Index).x = ShapeLegend(Index).Left
    ptTopLeft(Index).y = ShapeLegend(Index).Top
    ptBottomRight(Index).x = ShapeLegend(Index).Left + ShapeLegend(Index).Width
    ptBottomRight(Index).y = ShapeLegend(Index).Top + ShapeLegend(Index).Height
    ' �������߿�
    ShapeLegend(Index).Visible = False
    ' ���߿���ɫ������ɫȡ��ɫ������
    ShapeLegend(Index).BorderColor = ColorInverted(picGraph.BackColor)
End Sub

' ��һ�� Long ����ɫֵ�õ� R G Bֵ�ֱ��Ƕ��٣�
Private Sub ColorToRGB(ByVal lngColor As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    R = lngColor Mod 256
    G = lngColor \ 256 Mod 256
    B = lngColor \ 65536
End Sub

' ȡ��ɫ����ɫ a = rgb(x1,y2,z3) a �ķ�ɫ = RGB(255 - x1, 255 - y2, 255 - z3)
Private Function ColorInverted(ByVal lngOldColor As Long) As Long
    Dim R As Integer, G As Integer, B As Integer
    ' 1����һ�� Long ����ɫֵ�õ� R G Bֵ�ֱ��Ƕ��٣�
    Call ColorToRGB(lngOldColor, R, G, B)
    ' 2��ȡ��ɫ
    ColorInverted = RGB(255 - R, 255 - G, 255 - B)
End Function
' --- �ؼ��ڲ��ķ��������ڷ�����˽�з�����--------------------------------------------------------------


' ######################################################################################################
' --- �ؼ� �����¼� ---------------------------------------------------------------------------------
' ######################################################################################################
Private Sub picGraph_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picGraph_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
' ######################################################################################################
' --- �ؼ� �����¼� ---------------------------------------------------------------------------------
' ######################################################################################################


' === �ƶ�ͼ��˵�� =============================
' ����ڿؼ��ϰ���ʱ��
' �ر�ע�⣺���굥λ��ͬ��������ͼƬ���ǡ�vbPixels������� X Y �� vbTwips
' ��ˣ�y / Screen.TwipsPerPixelY
Private Sub lblLegend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    IsMovingControl(Index) = False
    If Button = vbLeftButton Then   ' ��������������ʱ��
        IsMovingControl(Index) = True
        ' ����ڿؼ��ϰ���ʱ����ǰ������ؼ����Ͻǵ�Ĳ
        ptOffset(Index).y = (y / Screen.TwipsPerPixelY - ptTopLeft(Index).y)
        ptOffset(Index).x = (x / Screen.TwipsPerPixelX - ptTopLeft(Index).x)
        ' ��ʾ���߿�
        ShapeLegend(Index).Visible = True
    End If
End Sub
' ����ڿؼ��ϰ��£����϶�ʱ��
Private Sub lblLegend_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsMovingControl(Index) Then
        ' �ƶ����߿�
        ShapeLegend(Index).Top = (y / Screen.TwipsPerPixelY - ptOffset(Index).y)
        ShapeLegend(Index).Left = (x / Screen.TwipsPerPixelX - ptOffset(Index).x)
    End If
End Sub
' ����ڿؼ��ϵ���ʱ��
Private Sub lblLegend_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsMovingControl(Index) Then
        IsMovingControl(Index) = False
        ' �������߿�
        ShapeLegend(Index).Visible = False
        ' ���³�ʼ�������Ͻǵ㡢���½ǵ� ����������
        ptTopLeft(Index).y = ShapeLegend(Index).Top
        ptTopLeft(Index).x = ShapeLegend(Index).Left
        ptBottomRight(Index).y = ShapeLegend(Index).Top + ShapeLegend(Index).Height
        ptBottomRight(Index).x = ShapeLegend(Index).Left + ShapeLegend(Index).Width
        ' �ƶ��ؼ�
        lblLegend(Index).Move ptTopLeft(Index).x, ptTopLeft(Index).y, ptBottomRight(Index).x - ptTopLeft(Index).x, ptBottomRight(Index).y - ptTopLeft(Index).y
        
        ' ĳ��ͼ���Ƿ��Ѿ��ƶ��ˣ�
        IsControlMoved(Index) = True
    End If
End Sub
' === �ƶ�ͼ��˵�� =============================
