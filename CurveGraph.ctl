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
' 功    能：绘制实时曲线（类似 Windows 任务管理器CPU使用率曲线）
' 使用方法：
' 作    者：鹤望兰·流
' 版    权：鹤望兰·流
' 发布日期：2009-05-01
' 网    站：http://hewanglan34512.cech.com.cn
' E - mail：hewanglan34512@163.com
' 依    赖：（无）
' 更    新：
' 2、2009-09-25~2009-09-26
'              (1) 添加曲线条数属性（在一个坐标轴上，一次可画多条曲线，1～MAX_CURVECOUNT 条可设置。）
'                  注意：设置条数后，每条曲线都会画（即使没有用 AddValue 添加值！添加 blnShouldDrawCurve 变量决定是否要画出），
'                  不过，建议想画出多少就设置多少，不要有多余的。
'              (2) 曲线颜色 属性改动，支持多条曲线不同颜色
'              (3) 曲线类型 属性改动，支持多条曲线不同类型
'              (4) 添加数据函数（AddValue），增加参数确定添加到哪一条曲线（第一条必须有数据！建议每条曲线都添加数据！）。
'              (5) 添加图例说明函数（FixLegend），如 — A。需要修改时，调用此函数即可！
'                  在有的Windows系统中字体为宋体时，…… 和 ——都只能显示为两竖线，改为其他字体就可以正常显示！
'              (6) 是否显示图例说明？可单独设置某一条曲线图例
'              (7) 添加图例字体属性，支持不同图例字体不同
'              (8) 各个图例说明可以单独移动（用鼠标拖动）
' 1、2009-05-17
'              (1) 添加函数，清空曲线 = ClearAll。清除原来的曲线！
'              (2) 增加垂直方向最小值属性 = MinVertical，可以支持负数！
'              (3) 为各个属性添加描述说明文字，方便用户使用！
' * (******* 复制请保留以上信息 *******)
' **************************************************************************************

Option Explicit
' ### API 函数申明 -------------------------------------------------------------------------------------

' ### 枚举申明 -----------------------------------------------------------------------------------------
' 边框类型' 注意：图片框边框类型总是 0，控件边框类型用 UserControl 的进行控制！
' 把 UserControl.Appearance 设置为 0 - Flat 也不好用 picGraph 图片框边框来控制。
Public Enum BoderStyleEnum
    [None] = 0
    [Fixed Single] = 1
End Enum
' 竖直网格线移动方式，也是曲线移动方向
Public Enum MovingGridEnum
    [Not Moving] = 0
    [Left to Right] = 1
    [Right to Left] = 2
End Enum
' 曲线类型。实线？虚线？
Public Enum CurveTypeEnum
    [Solid] = 0
    [Dot] = 1
End Enum
'Download by http://www.NewXing.com
' ### 常数申明 -----------------------------------------------------------------------------------------
Private Const LAST_LINE_TOLERANCE As Single = 0.0001    ' 用来修复在某些情况下最后一条线刻度问题
Private Const MAX_CURVECOUNT As Integer = 10            ' 曲线条数最多为 10 条！

' 成员变量 申明 =============================================================================
Private m_ShowGrid As Boolean                           ' 是否显示网格？
Private m_MovingGrid As MovingGridEnum                  ' 竖直网格线移动方式，当且仅当网格线非静止时，也是曲线移动方向
Private m_MovingCurve As MovingGridEnum                 ' 网格线静止时，曲线移动方向！注意：仅在网格线静止时有效！！！
Private m_HorizontalSplits As Long                      ' 水平方向分成多少份？画网格线用。
Private m_VerticalSplits As Long                        ' 铅垂方向分成多少份？画网格线用。
Private m_MaxVertical As Single                         ' 铅垂方向最大值。
Private m_MinVertical As Single                         ' 铅垂方向最大值。

Private m_HorizontalGridColor As OLE_COLOR              ' 水平网格线 颜色（默认为深绿色！）RGB(0, 130, 0)
Private m_VerticalGridColor As OLE_COLOR                ' 竖直网格线 颜色
Private m_CurveLineColor() As OLE_COLOR                   ' 曲线颜色（默认为绿色）RGB(0, 130, 0) 深绿色！
Private m_CurveLineType() As CurveTypeEnum                ' 曲线类型？实线？虚线？
Private m_AxesTextColor As OLE_COLOR                    ' 坐标轴文字颜色
Private m_ShowAxesText As Boolean                       ' 是否显示坐标轴文字？
Private m_xBarNowTimeFormat As String                   ' X 轴时间格式！

Private m_CurveCount As Integer                         ' 曲线条数（在一个坐标轴上，一次可画多条曲线）
Private m_ShowLegend() As Boolean                         ' 是否显示图例说明？可单独设置某一条曲线图例

' --- 私有变量 申明 --------------------------------------------------------
Private picGraphHeight As Long                          ' 图片框高度
Private picGraphWidth As Long                           ' 宽度
Private GridPosition As Long                            ' 竖直网格线位置！
Private StartPosition As Long                           ' 起始画曲线的位置！统一变化，曲线从右到左或从左到右具体画的时候在考虑不同！当开始一个新图时，不需要显示零值（Needed to not to display first zero values when starting a new diagram）

Private Type YValues
    yValueArray() As Single                         ' 保存Y轴坐标值的数组。
End Type
Private yV() As YValues

Private xBarText() As String                            ' 保存X轴显示的文字的数组！
'Private StartXBarText As Long                           ' X 轴上文字位置
Private VerticalGridIndex As Long                       ' 竖直网格线序号，X轴上文字位置序号。

Private strLegendLine As String                         ' 图例上的线型！
Private blnShouldDrawCurve() As Boolean                 ' 是否要画那条曲线？其初始化在重新定义数组时第一个元素初始化为 True，其他的使用默认初始化值 False）

' --- 结构体 申明
' === 移动图例说明 =============================
Private Type PointXY ' 鼠标点的 X、Y坐标位置。
    x As Single
    y As Single
End Type
Private IsMovingControl() As Boolean   ' 标识：是否正在移动控件？
Private ptTopLeft() As PointXY         ' 控件左上角的一点
Private ptBottomRight() As PointXY     ' 控件右上角的一点
Private ptOffset() As PointXY          ' 鼠标在控件上按下时，当前鼠标点与控件左上角点的差！
Private IsControlMoved() As Boolean    ' 某个图例是否已经移动了？
' 注意：只要第一个移动了，在 UserControl_Resize 就不改变图例说明的位置了！
' === 移动图例说明 =============================

' --- 事件申明 ---------------------------------------------------------------
'Event Declarations:
Event DblClick() 'MappingInfo=picGraph,picGraph,-1,DblClick
Attribute DblClick.VB_Description = "当用户在一个对象上按下并释放鼠标按钮后再次按下并释放鼠标按钮时发生。"
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picGraph,picGraph,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picGraph,picGraph,-1,MouseUp
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=picGraph,picGraph,-1,MouseDown


' ######################################################################################################
' ### 公共方法 ###
' ######################################################################################################
' 显示关于对话框
'Public Sub About()
'    MsgBox "贺兰_实时曲线控件 -> 绘制实时曲线（类似 Windows 任务管理器CPU使用率曲线）。" & vbCrLf & vbCrLf _
'         & "版权所有(C) 2009 鹤望兰·流" & vbCrLf & vbCrLf _
'         & "主页: http://hewanglan.ys168.com", vbInformation + vbSystemModal, "[贺兰] - 鹤望兰·流"
'End Sub

' 添加数据，用于画曲线！
Public Sub AddValue(ByVal yValue As Single, Optional ByVal Index As Integer = 1) ', Optional ByVal xBarString As String)
Attribute AddValue.VB_Description = "添加Y坐标数据，用于画曲线！"
    Dim i As Long
    
    ' 增加垂直方向最小值属性 = MinVertical，可以支持负数！
    yValue = yValue - m_MinVertical
    
    ' 是否要画那条曲线？注意：其他的都先设置为 False，不能这样（这样每次只会画一条曲线！）
'    For I = 1 To m_CurveCount
'        blnShouldDrawCurve(I) = False
'    Next I
    blnShouldDrawCurve(Index) = True
    ' 更新图例说明
    lblLegend(Index).Visible = m_ShowLegend(Index) And blnShouldDrawCurve(Index)
    
    ' 将数组所有元素前移（下标高一位的赋值给低一位的）
    For i = 1 To picGraphWidth - 1
        yV(Index).yValueArray(i - 1) = yV(Index).yValueArray(i)
    Next i

    ' 将新值添加到数组最后一个元素，注意：这里用 I-1，因为上面循环结束后 I 会再自动 +1 ！
    yV(Index).yValueArray(i - 1) = picGraphHeight - ((yValue / (m_MaxVertical - m_MinVertical)) * picGraphHeight)
    
    ' 加个判断，只在第一次时设置一次！！！否则画图时每条曲线都不太对，X轴时间也不对。
    ' 起始画曲线的位置！
    If Index = 1 Then
        If StartPosition >= 1 Then StartPosition = StartPosition - 1
    
        ' 竖直网格线位置移动！
        If m_MovingGrid = [Right to Left] Then      ' 从右到左？
            GridPosition = GridPosition - 1
        ElseIf m_MovingGrid = [Left to Right] Then  ' 从左到右？
            GridPosition = GridPosition + 1
        End If
        
        ' === X 轴上显示的文字 ==================================================================
        If Len(xBarText(0)) = 0 Then xBarText(0) = Format$(Now, m_xBarNowTimeFormat)
    End If
'    StartXBarText = StartXBarText - 1           ' X 轴上文字移动
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
'        ' 将数组所有元素前移
'        For I = 1 To m_HorizontalSplits
'            xBarText(I - 1) = xBarText(I)
'        Next I
'        xBarText(I - 1) = xBarString
'    End If
End Sub
' 添加图例说明，如 — A。！需要修改时，调用此函数即可！
Public Sub FixLegend(ByVal strText As String, Optional ByVal Index As Integer = 1)
Attribute FixLegend.VB_Description = "添加图例说明，如 — A。！需要修改时，调用此函数即可！"
    lblLegend(Index).Caption = strLegendLine & strText
    lblLegend(Index).ForeColor = m_CurveLineColor(Index)
    
    ' 初始化 Shape 图例移动效果框
    Call InitLegendShape(Index)
End Sub

' 先画水平、竖直两种网格线，再画曲线！
Public Function DrawGridCurve()
Attribute DrawGridCurve.VB_Description = "核心函数，绘制水平、竖直两种网格线，绘制曲线，绘制X、Y坐标文字。"
    Dim x As Single
    Dim y As Single
    Dim i As Long
    ' 第 KK 条曲线
    Dim KK As Integer
    
    ' 先清空图片框
    picGraph.Cls

    ' 1、画网格线
    If m_ShowGrid Then
        ' 1.1 水平网格线
        For y = 0 To (picGraphHeight - 1) Step ((picGraphHeight - 1) / (m_VerticalSplits)) - LAST_LINE_TOLERANCE
            picGraph.Line (0, y)-(picGraphWidth, y), m_HorizontalGridColor
        Next y
        ' 1.2 竖直网格线（注意：分3种情况，移动：不移动？从右到左？从左到右？）
        If m_MovingGrid = [Not Moving] Then ' 不需要移动，静态网格线。
            For x = 0 To (picGraphWidth - 1) Step ((picGraphWidth - 1) / (m_HorizontalSplits)) - LAST_LINE_TOLERANCE
                picGraph.Line (x, 0)-(x, picGraphHeight), m_VerticalGridColor
            Next x
        Else ' 从右到左？' 从左到右？ 一样的代码，只是 AddValue 函数中 GridPosition 变化趋势不一样！GridPosition = GridPosition - 1 或 + 1
            For x = GridPosition To (picGraphWidth - 1) Step ((picGraphWidth - 1) / (m_HorizontalSplits)) - LAST_LINE_TOLERANCE
                picGraph.Line (x, 0)-(x, picGraphHeight), m_VerticalGridColor
            Next x
        End If
    End If
    ' 1.3 当第一个竖直网格线不见时，重置其位置！（注意：分2种情况，移动：从右到左？从左到右？）
    If m_MovingGrid = [Right to Left] Then      ' 从右到左？
        If GridPosition <= -Int((picGraphWidth - 1) / m_HorizontalSplits) Then
            GridPosition = 0
            ' 添加时间，注意：先 +1 ， 0 在初始化时添加！
            VerticalGridIndex = VerticalGridIndex + 1
            xBarText(VerticalGridIndex) = Format$(Now, m_xBarNowTimeFormat)
        End If
    ElseIf m_MovingGrid = [Left to Right] Then  ' 从左到右？
        If GridPosition >= Int((picGraphWidth - 1) / m_HorizontalSplits) Then
            GridPosition = 0
            ' 添加时间，注意：先 +1 ， 0 在初始化时添加！
            VerticalGridIndex = VerticalGridIndex + 1
            xBarText(VerticalGridIndex) = Format$(Now, m_xBarNowTimeFormat)
        End If
    End If
    ' 当所有竖直线都显示了文字后，清零，从头在来显示，正因此，每次显示完后就会又从第零个网格线开始，其他文字被清除！
    If VerticalGridIndex >= m_HorizontalSplits - 1 Then VerticalGridIndex = 0: xBarText(0) = Format$(Now, m_xBarNowTimeFormat)
'    If StartXBarText <= -(picGraphWidth - 1) Then StartXBarText = 0

    ' 2、画曲线
    ' Draw line diagram only if there are 2 or more values defined
    If m_MovingGrid = [Right to Left] Then      ' 从右到左？
        ' ----------------------------------------------------------------------------------------------
        If StartPosition <= picGraphWidth - 1 Then
            For KK = 1 To m_CurveCount
                If blnShouldDrawCurve(KK) Then ' 是否要画那条曲线？
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
    ElseIf m_MovingGrid = [Left to Right] Then  ' 从左到右？与上面相比， I 前加了 picGraphWidth - 1 -
        ' ----------------------------------------------------------------------------------------------
        If StartPosition <= picGraphWidth - 1 Then
            For KK = 1 To m_CurveCount
                If blnShouldDrawCurve(KK) Then ' 是否要画那条曲线？
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
    Else ' 网格线静止时，曲线移动方向！
        ' **********************************************************************************************
        If m_MovingCurve = [Right to Left] Or m_MovingCurve = [Not Moving] Then   ' 从右到左？曲线不能静止，默认从右到左移动！
            If StartPosition <= picGraphWidth - 1 Then
                For KK = 1 To m_CurveCount
                    If blnShouldDrawCurve(KK) Then ' 是否要画那条曲线？
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
        Else  ' 从左到右？
            If StartPosition <= picGraphWidth - 1 Then
                For KK = 1 To m_CurveCount
                    If blnShouldDrawCurve(KK) Then ' 是否要画那条曲线？
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
    
    ' 3、坐标值文字显示
    If m_ShowAxesText Then
        With picGraph
            ' 3.1 Y 轴。最大刻度！
            .CurrentX = 3
            .CurrentY = 0
            picGraph.Print Format$(m_MaxVertical, "0.0#")
            ' 中间刻度
            For i = 1 To m_VerticalSplits - 1
                .CurrentX = 3
                .CurrentY = Int(picGraphHeight / m_VerticalSplits * i - picGraph.TextHeight(CStr(i)) / 2)
                picGraph.Print Format$(m_MaxVertical - ((m_MaxVertical - m_MinVertical) / m_VerticalSplits * i), "0.0#")
            Next i
            ' 最小刻度
            .CurrentX = 3
            .CurrentY = picGraphHeight - picGraph.TextHeight(CStr(i))
            picGraph.Print m_MinVertical
            
            ' 3.2 X 轴文字。
'            X = Abs(StartXBarText)
            If m_MovingGrid = [Left to Right] Then
                For i = 0 To VerticalGridIndex
                    ' X 位置坐标的确定，调试了近1小时。。。。
                    picGraph.CurrentX = GridPosition + i * (((picGraphWidth - 1) / (m_HorizontalSplits))) - picGraph.TextWidth(xBarText(VerticalGridIndex - i)) / 2
                    picGraph.CurrentY = picGraphHeight - picGraph.TextHeight(CStr(Int(x)))
                    picGraph.Print xBarText(VerticalGridIndex - i)
                Next i
            ElseIf m_MovingGrid = [Right to Left] Then
                For i = 0 To VerticalGridIndex
                    ' X 位置坐标的确定，调试了近1小时。。。。
                    picGraph.CurrentX = picGraphWidth + GridPosition - i * (((picGraphWidth - 1) / (m_HorizontalSplits))) - picGraph.TextWidth(xBarText(VerticalGridIndex - i)) / 2
                    picGraph.CurrentY = picGraphHeight - picGraph.TextHeight(CStr(Int(x)))
                    picGraph.Print xBarText(VerticalGridIndex - i)
                Next i
            End If
        End With
    End If
End Function

' 清空曲线、X 坐标时间等。
Public Sub ClearAll()
Attribute ClearAll.VB_Description = "清空曲线、X 坐标时间等。"
    ' 私有变量 初始化（注意：要在 UserControl_Resize 中再写一遍？否则初始化可能不对！！！）
    picGraphHeight = picGraph.ScaleHeight   ' 图片框高度
    picGraphWidth = picGraph.ScaleWidth     ' 宽度
    GridPosition = 0                        ' 竖直网格线位置
    StartPosition = picGraphWidth - 1       ' 起始画曲线的位置！
    VerticalGridIndex = 0                   ' 竖直网格线序号，X轴上文字位置序号。
    
    ' 是否要画那条曲线？
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
' ### 公共方法 ###
' ######################################################################################################



' ######################################################################################################
' ### 公共属性  ###
' ######################################################################################################
' 属性: 图片框是否允许自动重绘
Public Property Get AutoRedraw() As Boolean ' 获得属性值
Attribute AutoRedraw.VB_Description = "返回/设置 控件是否允许自动重绘！"
    AutoRedraw = picGraph.AutoRedraw
End Property
Public Property Let AutoRedraw(ByVal New_Value As Boolean) ' 设置属性
    picGraph.AutoRedraw = New_Value
    PropertyChanged "AutoRedraw"
End Property

' 属性: 图片框背景颜色
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置 控件背景颜色。"
    BackColor = picGraph.BackColor
End Property
Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    picGraph.BackColor = New_Value
    ' 虚线框颜色，背景色取反色！！！
    Dim i As Integer
    For i = 1 To m_CurveCount
        ShapeLegend(i).BorderColor = ColorInverted(picGraph.BackColor)
    Next i
    PropertyChanged "BackColor"
End Property

' 属性: 控件边框类型
Public Property Get BorderStyle() As BoderStyleEnum
Attribute BorderStyle.VB_Description = "返回/设置 控件边框类型。"
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_Value As BoderStyleEnum)
    UserControl.BorderStyle = New_Value
    PropertyChanged "BorderStyle"
End Property

' 属性: 是否显示网格？
Public Property Get ShowGrid() As Boolean
Attribute ShowGrid.VB_Description = "返回/设置 是否显示水平和竖直网格线。"
    ShowGrid = m_ShowGrid
End Property
Public Property Let ShowGrid(ByVal New_Value As Boolean)
    m_ShowGrid = New_Value
    PropertyChanged "ShowGrid"
End Property

' 属性: 竖直网格线移动方式，也是曲线移动方向
Public Property Get MovingGrid() As MovingGridEnum
Attribute MovingGrid.VB_Description = "返回/设置 竖直网格线移动方式，也是曲线移动方向。"
    MovingGrid = m_MovingGrid
End Property
Public Property Let MovingGrid(ByVal New_Value As MovingGridEnum)
    m_MovingGrid = New_Value
    
    ' 调用函数 初始化变量？
    Call ClearAll
'    ' 起始画曲线的位置！
'    ' 竖直网格线序号，X轴上文字位置序号。
'    VerticalGridIndex = 0
'    StartPosition = picGraphWidth - 1
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)

    PropertyChanged "MovingGrid"
End Property

' 属性: 网格线静止时，曲线移动方向！注意：仅在网格线静止时有效！！！
Public Property Get MovingCurve() As MovingGridEnum
Attribute MovingCurve.VB_Description = "返回/设置 网格线静止时，曲线移动方向！注意：仅在网格线静止时有效！！！"
    MovingCurve = m_MovingCurve
End Property
Public Property Let MovingCurve(ByVal New_Value As MovingGridEnum)
    m_MovingCurve = New_Value
    
    ' 调用函数 初始化变量？
    Call ClearAll
'    ' 起始画曲线的位置！
'    StartPosition = picGraphWidth - 1
'    ' 竖直网格线序号，X轴上文字位置序号。
'    VerticalGridIndex = 0
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)
    
    PropertyChanged "MovingCurve"
End Property

' 属性: 水平分成多少份？画网格线用。
Public Property Get HorizontalSplits() As Long
Attribute HorizontalSplits.VB_Description = "返回/设置 水平分成多少份？画网格线用。"
    HorizontalSplits = m_HorizontalSplits
End Property
Public Property Let HorizontalSplits(ByVal New_Value As Long)
    ' 值不能<=0
    If New_Value <= 0 Then New_Value = 1
    m_HorizontalSplits = New_Value
    PropertyChanged "HorizontalSplits"
End Property

' 属性: 铅垂方向分成多少份？画网格线用。
Public Property Get VerticalSplits() As Long
Attribute VerticalSplits.VB_Description = "返回/设置 铅垂方向分成多少份？画网格线用。"
    VerticalSplits = m_VerticalSplits
End Property
Public Property Let VerticalSplits(ByVal New_Value As Long)
    ' 值不能<=0
    If New_Value <= 0 Then New_Value = 1
    m_VerticalSplits = New_Value
    PropertyChanged "VerticalSplits"
End Property

' 属性: 铅垂方向最大值。
Public Property Get MaxVertical() As Single
Attribute MaxVertical.VB_Description = "返回/设置 铅垂方向最大值。"
    MaxVertical = m_MaxVertical
End Property
Public Property Let MaxVertical(ByVal New_Value As Single)
    ' 最大值不能<=0
    If New_Value <= 0 Then New_Value = 1
    m_MaxVertical = New_Value
    PropertyChanged "MaxVertical"
End Property

' 属性: 铅垂方向最小值。
Public Property Get MinVertical() As Single
Attribute MinVertical.VB_Description = "返回/设置 铅垂方向最小值。"
    MinVertical = m_MinVertical
End Property
Public Property Let MinVertical(ByVal New_Value As Single)
    ' 最小值不能>=最大值
    If New_Value >= m_MaxVertical Then New_Value = m_MaxVertical - 1
    m_MinVertical = New_Value
    PropertyChanged "MinVertical"
End Property

' 属性: 水平网格线 颜色
Public Property Get HorizontalGridColor() As OLE_COLOR
Attribute HorizontalGridColor.VB_Description = "返回/设置 水平网格线颜色。"
    HorizontalGridColor = m_HorizontalGridColor
End Property
Public Property Let HorizontalGridColor(ByVal New_Value As OLE_COLOR)
    m_HorizontalGridColor = New_Value
    PropertyChanged "HorizontalGridColor"
End Property

' 属性: 竖直网格线 颜色
Public Property Get VerticalGridColor() As OLE_COLOR
Attribute VerticalGridColor.VB_Description = "返回/设置 竖直网格线颜色。"
    VerticalGridColor = m_VerticalGridColor
End Property
Public Property Let VerticalGridColor(ByVal New_Value As OLE_COLOR)
    m_VerticalGridColor = New_Value
    PropertyChanged "VerticalGridColor"
End Property

' 属性: 曲线颜色
Public Property Get CurveLineColor(Optional ByVal Index As Integer = 1) As OLE_COLOR
Attribute CurveLineColor.VB_Description = "返回/设置 曲线颜色。"
    CurveLineColor = m_CurveLineColor(Index)
End Property
Public Property Let CurveLineColor(Optional ByVal Index As Integer = 1, ByVal New_Value As OLE_COLOR)
    m_CurveLineColor(Index) = New_Value
    ' 更新图例显示
    lblLegend(Index).ForeColor = m_CurveLineColor(Index)
    PropertyChanged "CurveLineColor"
End Property

' 属性: 曲线类型：实线？虚线？
Public Property Get CurveLineType(Optional ByVal Index As Integer = 1) As CurveTypeEnum
Attribute CurveLineType.VB_Description = "返回/设置 曲线类型：实线？虚线？"
    CurveLineType = m_CurveLineType(Index)
End Property
Public Property Let CurveLineType(Optional ByVal Index As Integer = 1, ByVal New_Value As CurveTypeEnum)
    m_CurveLineType(Index) = New_Value
    ' 更新图例显示
    On Error Resume Next
    strLegendLine = IIf(m_CurveLineType(Index) = [Solid], "— ", "… ")
    FixLegend Right$(lblLegend(Index).Caption, Len(lblLegend(Index).Caption) - 2), Index
    
    PropertyChanged "CurveLineType"
End Property

' 属性: 坐标轴文字颜色
Public Property Get AxesTextColor() As OLE_COLOR
Attribute AxesTextColor.VB_Description = "返回/设置 坐标轴文字颜色。"
    AxesTextColor = m_AxesTextColor
End Property
Public Property Let AxesTextColor(ByVal New_Value As OLE_COLOR)
    m_AxesTextColor = New_Value
    picGraph.ForeColor = m_AxesTextColor
    PropertyChanged "AxesTextColor"
End Property

' 属性: 是否显示坐标轴文字
Public Property Get ShowAxesText() As Boolean
Attribute ShowAxesText.VB_Description = "返回/设置 是否显示坐标轴文字？"
    ShowAxesText = m_ShowAxesText
End Property
Public Property Let ShowAxesText(ByVal New_Value As Boolean)
    m_ShowAxesText = New_Value
    PropertyChanged "ShowAxesText"
End Property

' 属性: X 坐标轴文字（时间）格式！
Public Property Get xBarNowTimeFormat() As String
Attribute xBarNowTimeFormat.VB_Description = "返回/设置 X 坐标轴文字（时间）格式！"
    xBarNowTimeFormat = m_xBarNowTimeFormat
End Property
Public Property Let xBarNowTimeFormat(ByVal New_Value As String)
    m_xBarNowTimeFormat = New_Value
    PropertyChanged "xBarNowTimeFormat"
End Property

'' 属性: 坐标轴文字字体
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picGraph,picGraph,-1,Font
Public Property Get AxesFont() As Font
Attribute AxesFont.VB_Description = "返回/设置 X 和 Y 坐标轴文字字体！"
    Set AxesFont = picGraph.Font
End Property

Public Property Set AxesFont(ByVal New_AxesFont As Font)
    Set picGraph.Font = New_AxesFont
    PropertyChanged "AxesFont"
End Property

' 属性: 曲线条数（在一个坐标轴上，一次可画多条曲线）
Public Property Get CurveCount() As Integer
Attribute CurveCount.VB_Description = "返回/设置 曲线条数（在一个坐标轴上，一次可画多条曲线）"
    CurveCount = m_CurveCount
End Property
Public Property Let CurveCount(ByVal New_Value As Integer)
    ' 条数限制 1～MAX_CURVECOUNT 条
    If New_Value < 1 Then New_Value = 1
    If New_Value > MAX_CURVECOUNT Then New_Value = MAX_CURVECOUNT
    m_CurveCount = New_Value
    
    ' 重新定义数组：曲线颜色，线型
    ReDim Preserve m_CurveLineColor(1 To m_CurveCount) As OLE_COLOR      ' 曲线颜色（默认为绿色）RGB(0, 130, 0) 深绿色！
    ReDim Preserve m_CurveLineType(1 To m_CurveCount) As CurveTypeEnum   ' 曲线类型？实线？虚线？
    ReDim m_ShowLegend(1 To m_CurveCount) As Boolean                     ' 是否显示图例说明？可单独设置某一条曲线图例
    
    ReDim blnShouldDrawCurve(1 To m_CurveCount) As Boolean              ' 是否要画那条曲线？
    blnShouldDrawCurve(1) = True
    ReDim yV(1 To m_CurveCount) As YValues                              ' 保存Y轴坐标值的数组。
    
    ' === 移动图例说明 =============================
    ReDim IsMovingControl(1 To m_CurveCount) As Boolean   ' 标识：是否移动控件？
    ReDim ptTopLeft(1 To m_CurveCount) As PointXY         ' 控件左上角的一点
    ReDim ptBottomRight(1 To m_CurveCount) As PointXY     ' 控件右上角的一点
    ReDim ptOffset(1 To m_CurveCount) As PointXY          ' 鼠标在控件上按下时，当前鼠标点与控件左上角点的差！
    ReDim IsControlMoved(1 To m_CurveCount) As Boolean    ' 某个图例是否已经移动了？
    ' === 移动图例说明 =============================
    
    Dim i As Integer
    For i = 1 To m_CurveCount
        ReDim yV(i).yValueArray(picGraphWidth - 1)
        m_CurveLineColor(i) = vbGreen              ' 曲线颜色
        m_CurveLineType(i) = [Solid]               ' 曲线类型：实线？虚线？
        m_ShowLegend(i) = True
    Next i
    ' 初始化 图例文本标签 属性
    Call InitLegendLabel
    
    PropertyChanged "CurveCount"
End Property

' 属性: 是否显示图例说明，注意： And
Public Property Get ShowLegend(Optional ByVal Index As Integer = 1) As Boolean
Attribute ShowLegend.VB_Description = "返回/设置 是否显示某条曲线的图例说明。当那条曲线没有用AddValue添加数据时，本属性无效！"
    ShowLegend = m_ShowLegend(Index) And blnShouldDrawCurve(Index)
End Property
Public Property Let ShowLegend(Optional ByVal Index As Integer = 1, ByVal New_Value As Boolean)
    m_ShowLegend(Index) = New_Value

    ' 更新图例说明
    lblLegend(Index).Visible = m_ShowLegend(Index) And blnShouldDrawCurve(Index)
    PropertyChanged "ShowLegend"
End Property

' 属性: 图例说明文字字体
Public Property Get LegendFont(Optional ByVal Index As Integer = 1) As Font
Attribute LegendFont.VB_Description = "返回/设置 图例说明文字字体"
    Set LegendFont = lblLegend(Index).Font
End Property

Public Property Set LegendFont(Optional ByVal Index As Integer = 1, ByVal New_LegendFont As Font)
    Set lblLegend(Index).Font = New_LegendFont
    PropertyChanged "LegendFont"
End Property
' ######################################################################################################
' ### 公共属性  ###
' ######################################################################################################


' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' --- 用户控件自身事件 ---------------------------------------------------------------------------------
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' SSS 在控件添加至窗口时所做个初始化动作，控件被实例化的时候就会调用，也就是控件在加载的时候会被执行
Private Sub UserControl_Initialize()
    ' 成员变量 初始化
    ' 图片框以像素为绘图刻度（似乎没有必要！用像素可以减少下面画网格线循环的次数？）
    ' 可以换成VB默认的，但这样的话，下面 picGraphHeight、picGraphWidth 的赋值变化为 ScaleHeight, ScaleWidth ！）
    picGraph.ScaleMode = vbPixels ' vbTwips
    picGraph.AutoRedraw = True              ' 图片框允许自动重绘
    picGraph.BackColor = vbBlack            ' 图片框背景色
    picGraph.BorderStyle = [None]           ' 注意：图片框边框类型总是 0，控件边框类型用 UserControl 的！
    UserControl.BorderStyle = [None]
    
    m_ShowGrid = True                       ' 是否显示网格？
    m_MovingGrid = [Right to Left]          ' 竖直网格线移动方式，当且仅当网格线非静止时，也是曲线移动方向
    m_MovingCurve = [Right to Left]         ' 网格线静止时，曲线移动方向！注意：仅在网格线静止时有效！！！
    m_HorizontalSplits = 9                  ' 水平方向分成多少份？画网格线用。
    m_VerticalSplits = 9                    ' 铅垂方向分成多少份？画网格线用。
    m_MaxVertical = 9                       ' 铅垂方向最大值。
    m_MinVertical = 0                       ' 铅垂方向最小值。
    m_HorizontalGridColor = RGB(0, 130, 0)  ' 水平网格线 颜色
    m_VerticalGridColor = RGB(0, 130, 0)    ' 竖直网格线 颜色
    
    m_CurveCount = 1                        ' 曲线条数（在一个坐标轴上，一次可画多条曲线）
    
    
    ' 重新定义数组：曲线颜色，线型
    ReDim m_CurveLineColor(1 To m_CurveCount) As OLE_COLOR      ' 曲线颜色（默认为绿色）RGB(0, 130, 0) 深绿色！
    ReDim m_CurveLineType(1 To m_CurveCount) As CurveTypeEnum   ' 曲线类型？实线？虚线？
    ReDim m_ShowLegend(1 To m_CurveCount) As Boolean            ' 是否显示图例说明？可单独设置某一条曲线图例
    
    ReDim blnShouldDrawCurve(1 To m_CurveCount) As Boolean      ' 是否要画那条曲线？
    blnShouldDrawCurve(1) = True
    ReDim yV(1 To m_CurveCount) As YValues                      ' 保存Y轴坐标值的数组。
    
    ' === 移动图例说明 =============================
    ReDim IsMovingControl(1 To m_CurveCount) As Boolean   ' 标识：是否移动控件？
    ReDim ptTopLeft(1 To m_CurveCount) As PointXY         ' 控件左上角的一点
    ReDim ptBottomRight(1 To m_CurveCount) As PointXY     ' 控件右上角的一点
    ReDim ptOffset(1 To m_CurveCount) As PointXY          ' 鼠标在控件上按下时，当前鼠标点与控件左上角点的差！
    ReDim IsControlMoved(1 To m_CurveCount) As Boolean    ' 某个图例是否已经移动了？
    ' === 移动图例说明 =============================
    
    Dim i As Integer
    For i = 1 To m_CurveCount
        m_CurveLineColor(i) = vbGreen              ' 曲线颜色
        m_CurveLineType(i) = [Solid]               ' 曲线类型：实线？虚线？
        m_ShowLegend(i) = True                     ' 是否显示图例说明？可单独设置某一条曲线图例
    Next i
    ' 初始化 图例文本标签 属性
    Call InitLegendLabel
    
    m_AxesTextColor = vbWhite               ' 坐标轴文字颜色
    m_ShowAxesText = True                   ' 是否显示坐标轴文字
    m_xBarNowTimeFormat = "hh:mm:ss"        ' X 轴时间格式！
    
    ' 调用函数 初始化变量？
    Call ClearAll
'    ' 私有变量 初始化（注意：要在 UserControl_Resize 中再写一遍？否则初始化可能不对！！！）
'    picGraphHeight = picGraph.ScaleHeight   ' 图片框高度
'    picGraphWidth = picGraph.ScaleWidth     ' 宽度
'    GridPosition = 0                        ' 竖直网格线位置
'    StartPosition = picGraphWidth - 1       ' 起始画曲线的位置！
'    VerticalGridIndex = 0                   ' 竖直网格线序号，X轴上文字位置序号。
'
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)

End Sub

' SSS 控件实例第二次及以后重新创建时，读保存的属性值
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
' SSS 控件设计时的实例被销毁时（进入运行时），保存控件在设计时设置的属性值
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
' SSS 控件实例在销毁之前执行
Private Sub UserControl_Terminate()
    Erase yV   ' 保存Y轴坐标值的数组。
    Erase xBarText      ' 保存X轴显示的文字的数组！
End Sub
' SSS 当用户控件大小改变时，当改变控件实例大小时发生
Private Sub UserControl_Resize()
    ' 图片框控件位置和大小，注意：宽和高减去一个数，以免最后一个网格线无法显示！
    picGraph.Move 0, 0, UserControl.Width - 52, UserControl.Height - 52
    ' 设置 图例文本标签 属性' 某个图例是否已经移动了？
    Dim i As Integer
    If Not IsControlMoved(1) Then lblLegend(1).Left = (picGraph.ScaleWidth - lblLegend(1).Width) / 2
    ' 初始化 Shape 图例移动效果框
    Call InitLegendShape(1)
    For i = 2 To m_CurveCount
        If Not IsControlMoved(i) Then lblLegend(i).Move lblLegend(1).Left, lblLegend(i - 1).Top + lblLegend(i).Height
        ' 初始化 Shape 图例移动效果框
        Call InitLegendShape(i)
    Next i
    
    ' 调用函数 初始化变量？
    Call ClearAll
'    picGraphHeight = picGraph.ScaleHeight        ' 图片框高度
'    picGraphWidth = picGraph.ScaleWidth          ' 宽度
'    GridPosition = 0                             ' 竖直网格线位置
'    StartPosition = picGraphWidth - 1            ' 起始画曲线的位置！
'    VerticalGridIndex = 0                        ' 竖直网格线序号，X轴上文字位置序号。
'
'    ' Allocate array to hold all diagram values (value per pixel)
'    ReDim yValueArray(picGraphWidth - 1)
'    ReDim xBarText(m_HorizontalSplits)

End Sub
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' --- 用户控件自身事件 ---------------------------------------------------------------------------------
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$


' --- 控件内部的方法（友邻方法和私有方法）--------------------------------------------------------------
' 初始化 图例文本标签 属性
Private Sub InitLegendLabel()
    ' 图例文本标签第一个的位置等属性先写好。
    lblLegend(1).Caption = "— 1"
    lblLegend(1).Move (picGraph.ScaleWidth - lblLegend(1).Width) / 2, 6
    lblLegend(1).Visible = m_ShowLegend(1) And blnShouldDrawCurve(1)
    ' 初始化 Shape 图例移动效果框
    Call InitLegendShape(1)
        
    ' 其他文本标签属性
    Dim i As Integer
    For i = 2 To m_CurveCount
        ' 加载图例文本标签
        On Error Resume Next
        Load lblLegend(i)
        Load ShapeLegend(i)
        lblLegend(i).ForeColor = m_CurveLineColor(i)
        lblLegend(i).Caption = strLegendLine & CStr(i)
        lblLegend(i).Move lblLegend(1).Left, lblLegend(i - 1).Top + lblLegend(i).Height
        lblLegend(i).Visible = m_ShowLegend(i) And blnShouldDrawCurve(i)
        
        ' 初始化 Shape 图例移动效果框
        Call InitLegendShape(i)
    Next i
End Sub
' 初始化 Shape 图例移动效果框
Private Sub InitLegendShape(Optional ByVal Index As Integer)
    ' 初始化，左上角点、右下角点 两个变量！注意：ShapeLegend(index).Move
    ShapeLegend(Index).Move lblLegend(Index).Left, lblLegend(Index).Top, lblLegend(Index).Width, lblLegend(Index).Height
    ptTopLeft(Index).x = ShapeLegend(Index).Left
    ptTopLeft(Index).y = ShapeLegend(Index).Top
    ptBottomRight(Index).x = ShapeLegend(Index).Left + ShapeLegend(Index).Width
    ptBottomRight(Index).y = ShapeLegend(Index).Top + ShapeLegend(Index).Height
    ' 隐藏虚线框
    ShapeLegend(Index).Visible = False
    ' 虚线框颜色，背景色取反色！！！
    ShapeLegend(Index).BorderColor = ColorInverted(picGraph.BackColor)
End Sub

' 由一个 Long 型颜色值得的 R G B值分别是多少！
Private Sub ColorToRGB(ByVal lngColor As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    R = lngColor Mod 256
    G = lngColor \ 256 Mod 256
    B = lngColor \ 65536
End Sub

' 取反色，颜色 a = rgb(x1,y2,z3) a 的反色 = RGB(255 - x1, 255 - y2, 255 - z3)
Private Function ColorInverted(ByVal lngOldColor As Long) As Long
    Dim R As Integer, G As Integer, B As Integer
    ' 1、由一个 Long 型颜色值得的 R G B值分别是多少！
    Call ColorToRGB(lngOldColor, R, G, B)
    ' 2、取反色
    ColorInverted = RGB(255 - R, 255 - G, 255 - B)
End Function
' --- 控件内部的方法（友邻方法和私有方法）--------------------------------------------------------------


' ######################################################################################################
' --- 控件 公共事件 ---------------------------------------------------------------------------------
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
' --- 控件 公共事件 ---------------------------------------------------------------------------------
' ######################################################################################################


' === 移动图例说明 =============================
' 鼠标在控件上按下时，
' 特别注意：坐标单位不同！！！！图片框是　vbPixels，这里的 X Y 是 vbTwips
' 因此，y / Screen.TwipsPerPixelY
Private Sub lblLegend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    IsMovingControl(Index) = False
    If Button = vbLeftButton Then   ' 仅当鼠标左键按下时，
        IsMovingControl(Index) = True
        ' 鼠标在控件上按下时，当前鼠标点与控件左上角点的差！
        ptOffset(Index).y = (y / Screen.TwipsPerPixelY - ptTopLeft(Index).y)
        ptOffset(Index).x = (x / Screen.TwipsPerPixelX - ptTopLeft(Index).x)
        ' 显示虚线框
        ShapeLegend(Index).Visible = True
    End If
End Sub
' 鼠标在控件上按下，并拖动时，
Private Sub lblLegend_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsMovingControl(Index) Then
        ' 移动虚线框
        ShapeLegend(Index).Top = (y / Screen.TwipsPerPixelY - ptOffset(Index).y)
        ShapeLegend(Index).Left = (x / Screen.TwipsPerPixelX - ptOffset(Index).x)
    End If
End Sub
' 鼠标在控件上弹起时，
Private Sub lblLegend_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsMovingControl(Index) Then
        IsMovingControl(Index) = False
        ' 隐藏虚线框
        ShapeLegend(Index).Visible = False
        ' 重新初始化，左上角点、右下角点 两个变量！
        ptTopLeft(Index).y = ShapeLegend(Index).Top
        ptTopLeft(Index).x = ShapeLegend(Index).Left
        ptBottomRight(Index).y = ShapeLegend(Index).Top + ShapeLegend(Index).Height
        ptBottomRight(Index).x = ShapeLegend(Index).Left + ShapeLegend(Index).Width
        ' 移动控件
        lblLegend(Index).Move ptTopLeft(Index).x, ptTopLeft(Index).y, ptBottomRight(Index).x - ptTopLeft(Index).x, ptBottomRight(Index).y - ptTopLeft(Index).y
        
        ' 某个图例是否已经移动了？
        IsControlMoved(Index) = True
    End If
End Sub
' === 移动图例说明 =============================
