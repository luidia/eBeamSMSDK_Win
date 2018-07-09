﻿Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports Microsoft.Win32

Public Class frmPenEx
#Region "Member value"

    Dim m_DrawPen As Pen
    Dim bPenRunning = False
    Dim PacketCnt As Integer = 0
    Dim m_bDraw As Boolean = False

    Dim m_bDrag As Boolean
    Dim m_ptMouseMove As PointF
    Dim bPenConnecting As Boolean = False
    Dim bPenConnectExit As Boolean = False

    'for drawing .
    Dim m_Bmp As Bitmap              'bitmap for double buffering
    Dim m_BmpLayer As Bitmap          'bitmap for layer drawing in case of highlighter pen 
    Public m_GrBmp As Graphics         'virtual bitmap
    Public m_GrMain As Graphics     'graphics for picturebox control
    Public m_GrLayer As Graphics     'graphics for m_BmpLayer

    Dim intNumberOfPoints As Integer = 0

    Dim m_ReciveHanle As System.IntPtr
    Dim m_DrawHanle As System.IntPtr
    Dim m_pRec As New _pen_rec  'structure for Pen Data 

    Public bPenConnect As Boolean = False 'Pen is connect or not
    Public m_ConnectTime As System.Timers.Timer
    Public ConnectThread As Threading.Thread

    Dim m_PenStyle As PENSTYLE
    Dim m_Stroke As CPNFStroke
    Dim imgattr As ImageAttributes   'highlighter attribute
    Const Htransparency As Double = 0.3 ' transparency of highlighter 


    Dim m_frmMemory As frmMemoryImport

#End Region

#Region "Delegate"
    Delegate Sub ConnectDelegate()
    Delegate Sub connectPenTextDelegate(ByVal sTitle As String)
#End Region

#Region "Timer"
    Public Sub ConnectingTimerStart()
        m_ConnectTime = New System.Timers.Timer(10000)

        AddHandler m_ConnectTime.Elapsed, AddressOf ClickaliveTime_Elapsed

        m_ConnectTime.Start()
    End Sub

    Public Sub ConnectingTimerEnd()
        If m_ConnectTime Is Nothing Then Exit Sub

        m_ConnectTime.Stop()
        bPenConnectExit = True
        RemoveHandler m_ConnectTime.Elapsed, AddressOf ClickaliveTime_Elapsed
    End Sub

    Public Sub ClickaliveTime_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
        ConnectingTimerEnd()
    End Sub
#End Region

#Region "Initialize"
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub frmPenEx_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        cboPenStyle.Items.Add("Normal Pen")  ' bezier path
        cboPenStyle.Items.Add("Brush Pen")  ' with press
        cboPenStyle.Items.Add("Highlighter Pen")

        cboPenStyle.SelectedIndex = 0

        cboPenStyle.Visible = False
    End Sub
#End Region

#Region "Event"
#Region "Mouse Event"
    Sub picMainDoMouseDown(ByVal e As PointF, ByVal btn As Integer, ByVal press As Integer, ByVal bCtrlKey As Boolean)
        m_bDrag = True

        Select Case m_PenStyle
            Case PENSTYLE.PENSTYLE_BRUSH, PENSTYLE.PENSTYLE_HIGHLIGHT, PENSTYLE.PENSTYLE_NORMAL


                m_Stroke = New CPNFStroke
                m_Stroke.AddItem(e, press)
        End Select
    End Sub

    Sub picMainDoMouseMove(ByVal e As PointF, ByVal btn As Integer, ByVal Press As Integer, ByVal bCtrlKey As Boolean)
        Try
            If m_bDrag = True Then
                m_ptMouseMove.X = e.X
                m_ptMouseMove.Y = e.Y

                Select Case m_PenStyle

                    Case PENSTYLE.PENSTYLE_BRUSH
                        m_Stroke.AddItem(m_ptMouseMove, Press)
                        m_Stroke.DrawBrushByPenPressure(m_GrBmp, m_DrawPen, -1)
                        m_Stroke.DrawBrushByPenPressure(m_GrMain, m_DrawPen, -1)

                    Case PENSTYLE.PENSTYLE_HIGHLIGHT
                        m_Stroke.AddItem(m_ptMouseMove, Press)
                        Dim clipRect As RectangleF = m_Stroke.DrawLastQuard(m_GrLayer, m_DrawPen, -1)
                        picMain.Invalidate(Rectangle.Round(clipRect))



                    Case PENSTYLE.PENSTYLE_NORMAL
                        m_GrMain = picMain.CreateGraphics
                        m_GrMain.SmoothingMode = SmoothingMode.AntiAlias
                        m_Stroke.AddItem(m_ptMouseMove, Press)
                        m_Stroke.DrawLastQuard(m_GrBmp, m_DrawPen, -1)
                        m_Stroke.DrawLastQuard(m_GrMain, m_DrawPen, -1)

                End Select

            End If
        Catch ex As Exception
        Finally

        End Try
    End Sub

    Sub picMainDoMouseUp(ByVal e As PointF, ByVal btn As Integer, ByVal press As Integer, ByVal bCtrlKey As Boolean)
        If m_bDrag = True Then
            m_bDrag = False


            If m_PenStyle = PENSTYLE.PENSTYLE_HIGHLIGHT Then
                m_GrBmp.DrawImage(m_BmpLayer, New Rectangle(0, 0, m_Bmp.Width, m_Bmp.Height), _
                            0, 0, _
                            m_BmpLayer.Width, m_BmpLayer.Height, GraphicsUnit.Pixel, imgattr)
                m_GrLayer.Clear(Color.Transparent)
            End If
            picMain.Invalidate()

        End If
    End Sub
#End Region

#Region "Controls click event"
    Private Sub btnConnectPen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnectPen.Click



        If bPenConnect Then
            DisConnectPen()
            connectPenTextChange("Connect")
        Else
            InputLog("START PEN CONNECT...")
            m_ReciveHanle = Me.Handle
            m_DrawHanle = picMain.Handle

            If bPenConnecting = False Then
                ConnectThread = New Threading.Thread(AddressOf ConnectPen)
                ConnectThread.Start()

                bPenConnecting = True
            End If

        End If
    End Sub

    Private Sub btnDrawing_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDrawing.Click
        If m_bDraw Then
            lbLog.Visible = True
            picMain.Visible = False
            m_bDraw = False
            btnDrawing.Text = "Drawing"
            btnClearLog.Text = "Clear Log"
            cboPenStyle.Visible = False
        Else
            lbLog.Visible = False
            picMain.Visible = True
            m_bDraw = True
            btnDrawing.Text = "ShowLog"
            btnClearLog.Text = "Clear"
            cboPenStyle.Visible = True

            CreateBitmap(True)

        End If
    End Sub

    Private Sub btnCalibration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalibration.Click
        If bPenConnect Then
            If m_bDraw Then
                m_bDraw = False

                Dim frmCali As New frmCalibration
                SetReciveHandle(frmCali.Handle)
                frmCali.ShowDialog()

                SetReciveHandle(Handle)
                m_bDraw = True

            Else
                Dim frmCali As New frmCalibration
                SetReciveHandle(frmCali.Handle)
                frmCali.ShowDialog()

                SetReciveHandle(Handle)
            End If

        Else
            MessageBox.Show("Either device is off or not in good connection. Check device and try again. Or pair it again.")
        End If
    End Sub

    Private Sub btnReStartPen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReStartPen.Click
        If bPenConnect Then
            If bPenRunning Then
                bPenRunning = False
                btnReStartPen.Text = "ReStartPen"
            Else
                bPenRunning = True
                btnReStartPen.Text = "StopPen"
            End If
        End If
    End Sub

    Private Sub btnClearLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearLog.Click
        If m_bDraw Then
            CreateBitmap(True)

            picMain.Invalidate()
        Else
            lbLog.Items.Clear()
        End If
    End Sub

    Private Sub lbPacketCountClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbPacketCountClear.Click
        lbPacketCount.Text = String.Empty
        PacketCnt = 0
    End Sub
#End Region

#Region "Controls changed event"
    Private Sub TrackBar1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.ValueChanged
        SetPenDownThreshold(TrackBar1.Value)
    End Sub

    Private Sub cboPenStyle_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cboPenStyle.SelectedIndexChanged
        m_PenStyle = cboPenStyle.SelectedIndex

        If m_PenStyle = PENSTYLE.PENSTYLE_HIGHLIGHT Then
            m_DrawPen = New Pen(Color.Red, 3)
            m_DrawPen.SetLineCap(LineCap.Square, LineCap.Square, DashCap.Round)
            m_DrawPen.LineJoin = LineJoin.Round
            m_DrawPen.Color = Color.Red
        Else
            m_DrawPen = New Pen(Color.Black, 3)
            m_DrawPen.SetLineCap(LineCap.Round, LineCap.Round, DashCap.Flat)
            m_DrawPen.LineJoin = LineJoin.Round

        End If
    End Sub
#End Region

#Region "Paint event"
    Private Sub picMain_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles picMain.Paint
        If m_Bmp Is Nothing Then Exit Sub

        Dim gr As Graphics = e.Graphics
        Dim grState As GraphicsState = Nothing



        m_GrBmp.Dispose()
        m_GrBmp = Graphics.FromImage(m_Bmp)
        m_GrBmp.SmoothingMode = SmoothingMode.AntiAlias

        gr.DrawImage(m_Bmp, e.ClipRectangle, e.ClipRectangle, GraphicsUnit.Pixel)

        If m_bDrag Then
            Select Case m_PenStyle
                Case PENSTYLE.PENSTYLE_HIGHLIGHT
                    gr.DrawImage(m_BmpLayer, e.ClipRectangle, e.ClipRectangle.X, e.ClipRectangle.Y, e.ClipRectangle.Width, e.ClipRectangle.Height, GraphicsUnit.Pixel, imgattr)
            End Select
        End If
    End Sub
#End Region
#End Region

#Region "Connect / disconnect pen"
    Public Sub EquilReceiveHandle()
        'If you set this, USB connection message will be sent to the window handle.
        'without find connect pen
        SetReciveHandle(m_ReciveHanle)
    End Sub
    Public Sub ConnectPen()

        '        SetReciveHandle(m_ReciveHanle)
        Me.Invoke(New MethodInvoker(AddressOf Me.EquilReceiveHandle))
        SetDrawRect(picMain.Width, picMain.Height)

        If FindDevice() Then    'find usb connetion (only for over F/W version 6)
        Else
            PortSearch()        'find BT Connection

        End If


        ConnectingTimerStart() 'waiting for connection

        While bPenConnectExit = False
            Threading.Thread.Sleep(100)
        End While

        'check if connecting pen
        ConnectingTimerEnd()

        bPenConnectExit = False

        bPenConnect = CheckConnect()

        Dim del As ConnectDelegate = New ConnectDelegate(AddressOf DelegateUISet)
        Me.BeginInvoke(del)

        SetCalibration()
        bPenConnecting = False
        MouseFlagDisable()
    End Sub

    Public Sub DisConnectPen()
        Try
            OnDisconnect(0)
        Catch ex As Exception
        End Try
        Dim del As ConnectDelegate = New ConnectDelegate(AddressOf DelegateUISet)
        Me.BeginInvoke(del)

        bPenConnect = False

        MouseFlagDisable()
    End Sub
#End Region

#Region "General function"

    Public Sub MouseFlagDisable()
        If m_bDrag Then
            picMainDoMouseUp(m_ptMouseMove, MouseButtons.Left, 255, False)
        End If
    End Sub

    Public Sub InputLog(ByVal sLog As String)
        lbLog.Items.Add(sLog)
    End Sub

    Private Sub DelegateUISet()
        If bPenConnect Then
            bPenRunning = True
            InputLog("CONNECTED")
            btnConnectPen.Text = "DisConnect"
        Else
            InputLog("NOT_CONNECTED")
            btnConnectPen.Text = "Connect"
            'MessageBox.Show("Either device is off or not in good connection. Check device and try again. Or pair it again.")
        End If
    End Sub
    Public Function GetRegularRectFromPoint(ByVal ptStart As Point, ByVal ptEnd As Point) As Rectangle
        Dim rt2 As New Rectangle
        rt2.X = Math.Min(ptStart.X, ptEnd.X)
        rt2.Y = Math.Min(ptStart.Y, ptEnd.Y)

        rt2.Width = Math.Abs(ptEnd.X - ptStart.X)
        rt2.Height = Math.Abs(ptEnd.Y - ptStart.Y)
        Return rt2

    End Function
    Private Sub connectPenTextChange(ByVal sTitle As String)
        btnConnectPen.Text = sTitle
    End Sub

    Sub CreateBitmap(bCreate As Boolean)
        If m_Bmp IsNot Nothing Then m_Bmp.Dispose()
        If m_BmpLayer IsNot Nothing Then m_BmpLayer.Dispose()


        m_Bmp = New Bitmap(picMain.Width, picMain.Height)
        m_BmpLayer = New Bitmap(picMain.Width, picMain.Height)

        m_GrBmp = Graphics.FromImage(m_Bmp)
        m_GrLayer = Graphics.FromImage(m_BmpLayer)

        m_GrBmp.SmoothingMode = SmoothingMode.AntiAlias
        m_GrLayer.SmoothingMode = SmoothingMode.AntiAlias

        SetHPenAlpha()
    End Sub
    Public Sub SetHPenAlpha()
        If imgattr Is Nothing Then imgattr = New ImageAttributes() 'imgattr.Dispose()

        Dim values()() As Single = {New Single() {1, 0, 0, 0, 0}, _
                                 New Single() {0, 1, 0, 0, 0}, _
                                 New Single() {0, 0, 1, 0, 0}, _
                                 New Single() {0, 0, 0, Convert.ToSingle(Htransparency), 0}, _
                                 New Single() {0, 0, 0, 0, 1, 0}}
        Dim colMatrix As New ColorMatrix(values)
        imgattr.SetColorMatrix(colMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap)
    End Sub

    Public Sub SetCalibration()
        Dim ptStart, ptEnd As PointF
        Dim regWord As RegistryKey = Nothing
        If GetRegistryKey(regWord, True) Then
            If EquilModelCode = 3 Then
                'If key is not exist, each calibration value setting with letter
                ptStart.X = Single.Parse(regWord.GetValue("Equil2Left", "1737"))
                ptStart.Y = Single.Parse(regWord.GetValue("Equil2Top", "541"))
                ptEnd.X = Single.Parse(regWord.GetValue("Equil2Right", "3708"))
                ptEnd.Y = Single.Parse(regWord.GetValue("Equil2Bottm", "4277"))
            ElseIf EquilModelCode = 4 Then
                'If key is not exist, each calibration value setting with 6 fit X 4 fit Board
                ptStart.X = Single.Parse(regWord.GetValue("EquilMakerLeft", "1728"))
                ptStart.Y = Single.Parse(regWord.GetValue("EquilMakerTop", "45372"))
                ptEnd.X = Single.Parse(regWord.GetValue("EquilMakerRight", "13796"))
                ptEnd.Y = Single.Parse(regWord.GetValue("EquilMakerBottm", "54824"))
            ElseIf EquilModelCode = 5 Then 'eBeam Smart Marker

                ptStart.X = Single.Parse(regWord.GetValue("EquilMakerLeft", "1450"))
                ptStart.Y = Single.Parse(regWord.GetValue("EquilMakerTop", "44100"))
                ptEnd.X = Single.Parse(regWord.GetValue("EquilMakerRight", "20620"))
                ptEnd.Y = Single.Parse(regWord.GetValue("EquilMakerBottm", "56150"))

            Else
                ptStart.X = Single.Parse(regWord.GetValue("EquilLeft", "1737"))
                ptStart.Y = Single.Parse(regWord.GetValue("EquilTop", "541"))
                ptEnd.X = Single.Parse(regWord.GetValue("EquilRight", "3708"))
                ptEnd.Y = Single.Parse(regWord.GetValue("EquilBottm", "4277"))
            End If

            If bPenConnect Then
                If EquilModelCode = 5 Then
                    SetSMPosition(StationPosition.LEFT, True)
                    SetCalibration_Top(ptStart.X, ptStart.Y)
                    SetCalibration_Bottom(ptEnd.X, ptEnd.Y)
                    SetCalibration_End()
                    SaveCalibrationToDevice(ptStart.X, ptStart.Y, ptStart.X, ptEnd.Y, ptEnd.X, ptEnd.Y, ptEnd.X, ptStart.Y)

                Else
                    SetCalibration_Top(ptStart.X, ptStart.Y)
                    SetCalibration_Bottom(ptEnd.X, ptEnd.Y)
                    SetCalibration_End()

                End If
            End If
        Else
            If EquilModelCode = 3 Then
                'If key is not exist, each calibration value setting with letter
                ptStart.X = 1737
                ptStart.Y = 541
                ptEnd.X = 3708
                ptEnd.Y = 4277
            ElseIf EquilModelCode = 4 Then
                'If key is not exist, each calibration value setting with 6 fit X 4 fit Board
                ptStart.X = 1728
                ptStart.Y = 45372
                ptEnd.X = 13796
                ptEnd.Y = 54824
            ElseIf EquilModelCode = 5 Then
                'If key is not exist, each calibration value setting with 6 fit X 4 fit Board
                ptStart.X = 1728
                ptStart.Y = 44100
                ptEnd.X = 20620
                ptEnd.Y = 56150
            Else
                ptStart.X = 1737
                ptStart.Y = 541
                ptEnd.X = 3708
                ptEnd.Y = 4277
            End If

            If bPenConnect Then
                If EquilModelCode = 5 Then
                    SetSMPosition(StationPosition.LEFT, True)
                    SetCalibration_Top(ptStart.X, ptStart.Y)
                    SetCalibration_Bottom(ptEnd.X, ptEnd.Y)
                    SetCalibration_End()
                    SaveCalibrationToDevice(ptStart.X, ptStart.Y, ptStart.X, ptEnd.Y, ptEnd.X, ptEnd.Y, ptEnd.X, ptStart.Y)

                Else
                    SetCalibration_Top(ptStart.X, ptStart.Y)
                    SetCalibration_Bottom(ptEnd.X, ptEnd.Y)
                    SetCalibration_End()

                End If

                SaveRegisty(ptStart, ptEnd)
            End If
        End If
    End Sub

    Private Sub SaveRegisty(ByVal ptStart As PointF, ByVal ptEnd As PointF)
        Dim regWord As RegistryKey = Nothing
        If GetRegistryKey(regWord, False) Then
            If EquilModelCode = 3 Then
                regWord.SetValue("Equil2Left", ptStart.X.ToString)
                regWord.SetValue("Equil2Top", ptStart.Y.ToString)
                regWord.SetValue("Equil2Right", ptEnd.X.ToString)
                regWord.SetValue("Equil2Bottm", ptEnd.Y.ToString)
            ElseIf EquilModelCode = 4 Then
                regWord.SetValue("EquilMakerLeft", ptStart.X.ToString)
                regWord.SetValue("EquilMakerTop", ptStart.Y.ToString)
                regWord.SetValue("EquilMakerRight", ptEnd.X.ToString)
                regWord.SetValue("EquilMakerBottm", ptEnd.Y.ToString)
            Else
                regWord.SetValue("EquilLeft", ptStart.X.ToString)
                regWord.SetValue("EquilTop", ptStart.Y.ToString)
                regWord.SetValue("EquilRight", ptEnd.X.ToString)
                regWord.SetValue("EquilBottm", ptEnd.Y.ToString)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Return registry value
    ''' </summary>
    ''' <param name="regWord">return registry key</param>
    ''' <param name="bOpen">open or save</param>
    ''' <returns>True is exist key, false is not exist key.</returns>
    ''' <remarks></remarks>
    Public Function GetRegistryKey(ByRef regWord As Microsoft.Win32.RegistryKey, ByVal bOpen As Boolean) As Boolean
        Dim regClasses As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

        If bOpen Then
#If NET_4_OVER Then
            If Environment.Is64BitOperatingSystem Then
                regWord = regClasses.OpenSubKey("Software\Wow6432Node\PNF\Equilnote\Settings")
            Else
                regWord = regClasses.OpenSubKey("Software\PNF\Equilnote\Settings")
            End If
#Else
            regWord = regClasses.OpenSubKey("Software\PNF\EquilSimulator\Settings")
#End If
        Else
#If NET_4_OVER Then
            If Environment.Is64BitOperatingSystem Then
                regWord = regClasses.CreateSubKey("Software\Wow6432Node\PNF\Equilnote\Settings", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
            Else
                regWord = regClasses.CreateSubKey("Software\PNF\Equilnote\Settings", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
            End If
#Else
            regWord = regClasses.CreateSubKey("Software\PNF\EquilSimulator\Settings", Microsoft.Win32.RegistryKeyPermissionCheck.ReadWriteSubTree)
#End If
        End If

        regClasses.Close()
        If regWord Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region
    Public Event ShowDIList()
    Public Sub SetDIList() Handles Me.ShowDIList
        If m_frmMemory IsNot Nothing Then
        Else
            btnMemoryImport_Click(Nothing, Nothing)
        End If
        m_frmMemory.SetDIList()

    End Sub

#Region "Win Message"
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_DI_CHANGE_STATION
                MsgBox("Sensor position changed")

            Case WM_DISCONNECTUSB
                MsgBox("USB Disconnected")
            Case WM_DI_NEWPAGE
                MsgBox("press button of sensor(New Page)")
            Case WM_DI_DUPLICATE
                MsgBox("long press button of sensor(Duplicate Page)")
                ' Create New page and copy the data of current page to new page

            Case WM_RETURNMESSAGE
                If bPenConnect AndAlso bPenRunning Then
                    PacketCnt += 1
                    lbPacketCount.Text = PacketCnt.ToString

                    m_pRec = New _pen_rec

                    m_pRec = Marshal.PtrToStructure(m.WParam, New _pen_rec().GetType) 'get pen data from message

                    Dim press As Integer = m_pRec.P

                    Dim pt As New PointF

                    lbSensordis.Text = m_pRec.Sensor_dis
                    lbModelCode.Text = m_pRec.ModelCode
                    lbHWVer.Text = m_pRec.HWVer
                    lbMCU1Ver.Text = m_pRec.MCU1
                    lbMCU2Ver.Text = m_pRec.MCU2
                    lbRawX.Text = m_pRec.X
                    lbRawY.Text = m_pRec.Y
                    lbConvX.Text = m_pRec.TX.ToString
                    lbConvY.Text = m_pRec.TY.ToString
                    lbPressure.Text = press.ToString
                    lbTemperature.Text = m_pRec.T.ToString
                    lbStatus.Text = m_pRec.PenStatus.ToString
                    lbIRGAP.Text = m_pRec.IRGAP.ToString
                    lbStationFlag.Text = If(m_pRec.bRight = 1, "RIGHT", "LEFT")
                    If m_pRec.ModelCode = 4 Or m_pRec.ModelCode = 5 Then
                        Select Case m_pRec.PenTiming
                            '0x51 - Red marker
                            '0x52 - Green marker
                            '0x53 - Yellow marker
                            '0x54 - Blue marker
                            '0x0F - Pen up
                            '0x56 - Purple marker
                            '0x58 - Black marker
                            '0x59 - Eraser cap
                            '0x5A - Low battery
                            '0x5C - Big eraser
                            Case MakerState.RED
                                lbMakerState.Text = "Red"
                            Case MakerState.GREEN
                                lbMakerState.Text = "Green"
                            Case MakerState.YELLOW
                                lbMakerState.Text = "Yellow"
                            Case MakerState.BLUE
                                lbMakerState.Text = "Blue"
                            Case MakerState.PEN_UP
                                lbMakerState.Text = "Pen up"
                            Case MakerState.PURPLE
                                lbMakerState.Text = "Purple"
                            Case MakerState.BLACK
                                lbMakerState.Text = "Black"
                            Case MakerState.ERASER_CAP
                                lbMakerState.Text = "Eraser cap"
                            Case MakerState.LOW_BATTERY
                                lbMakerState.Text = "Low battery"
                            Case MakerState.BIG_ERASER
                                lbMakerState.Text = "Big eraser"
                        End Select
                    End If


                    If m_bDraw Then
                        Try
                            pt.X = m_pRec.TX ' Calibrated X
                            pt.Y = m_pRec.TY ' Calibrated Y
                            If EquilModelCode = 5 Then   'eBeam Smart Marker 
                                pt.Y *= picMain.Height / 65535.0
                                pt.X *= picMain.Width / 65535.0
                            Else

                            End If

                            Select Case m_pRec.PenStatus
                                Case PenStatus.PEN_DOWN
                                    picMainDoMouseDown(pt, MouseButtons.Left, press, False)
                                Case PenStatus.PEN_MOVE
                                    If m_bDrag = True Then
                                        If m_ptMouseMove.X * 1000 = pt.X * 1000 AndAlso m_ptMouseMove.Y * 1000 = pt.Y * 1000 Then 'Move 상태에서 중복된 데이터는 그릴필요가 없어서 제외함.
                                        Else
                                            picMainDoMouseMove(pt, MouseButtons.Left, press, False)
                                        End If
                                    End If
                                Case PenStatus.PEN_UP
                                    picMainDoMouseUp(pt, MouseButtons.Left, press, False)
                                Case PenStatus.PEN_HOVER
                            End Select

                        Catch ex As Exception
                            MouseFlagDisable()
                        End Try
                    End If
                End If
                Console.WriteLine("WM_RETURNMESSAGE")
            Case WM_DOWNLOAD_COMPLEATE
                Console.WriteLine("WM_DOWNLOAD_COMPLEATE")
                MsgBox("Import finished")
            Case WM_GESTUREMESSAGE
                MouseFlagDisable()

                Select Case m.LParam '
                    Case GestureMessage.GESTURE_CIRCLE_CLOCKWISE
                        InputLog("GESTURE CIRCLE CLOCKWISE")
                    Case GestureMessage.GESTURE_CIRCLE_COUNTERCLOCKWISE
                        InputLog("GESTURE CIRCLE COUNTERCLOCKWISE")
                    Case GestureMessage.GESTURE_CLICK
                        InputLog("GESTURE CLICK")
                End Select
            Case WM_PENCALIBRATION ' eBeam SM
                Console.WriteLine("WM_PENCALIBRATION")
                Dim pCondition As New PENConditionData
                pCondition = Marshal.PtrToStructure(m.WParam, New PENConditionData().GetType) 'PEN DATA STRUCTURE
                EquilModelCode = pCondition.modelCode                    'save equil model number
                bPenConnectExit = True
                bPenConnect = True
                If EquilModelCode = 4 Or EquilModelCode = 5 Then
                    cboAudioLang.Text = GetDeviceVoiceLanguage(GetAudioLang())
                    txtAudioMode.Text = GetAudioMode
                    txtAudioVolume.Text = GetAudioVolume


                End If

            Case WM_PENCONDITION
                'Console.WriteLine("WM_PENCONDITION")
                Dim pCondition As New PENConditionData
                pCondition = Marshal.PtrToStructure(m.WParam, New PENConditionData().GetType) 'PEN DATA STRUCTURE
                lbModelCode.Text = pCondition.modelCode
                lbBatteryStation.Text = pCondition.battery_station
                lbBatteryPen.Text = pCondition.battery_pen
                lbConnectionType.Text = If(pCondition.usbConnect = 1, "USB", "BT")
                Select Case pCondition.StationPosition
                    Case 1

                        lbStationPosition.Text = "TOP"
                    Case 2
                        lbStationPosition.Text = "LEFT"
                    Case 3
                        lbStationPosition.Text = "RIGHT"
                    Case 4
                        lbStationPosition.Text = "BOTTOM"

                End Select

                CURRENT_MARKER_DIRECT = pCondition.StationPosition

                PacketCnt += 1 '
                lbPacketCount.Text = PacketCnt.ToString

                lbAliveSec.Text = m.WParam.ToString

                If m.LParam = 1 Then 'power off of receiver
                    DisConnectPen()
                    MouseFlagDisable()
                    InputLog("SESSION CLOSED")
                ElseIf m.LParam = 2 Then '

                ElseIf m.LParam = 3 Then 'after connecting firt data receiving
                    If pCondition.modelCode <> 5 Then
                        EquilModelCode = pCondition.modelCode                    'save equil model number

                        bPenConnectExit = True
                        bPenConnect = True
                        If EquilModelCode = 4 Or EquilModelCode = 5 Then
                            cboAudioLang.Text = GetDeviceVoiceLanguage(GetAudioLang())
                            txtAudioMode.Text = GetAudioMode
                            txtAudioVolume.Text = GetAudioVolume


                        End If

                    End If

                    InputLog("RECV FIRST DATA")
                ElseIf m.LParam = 0 Then
                    If Not bPenConnect Then 'USB auto recognition.
                        bPenConnect = True
                        EquilModelCode = pCondition.modelCode                    'save equil model number
                    End If

                End If


                If bPenConnect AndAlso bPenRunning Then
                    MouseFlagDisable()
                End If
            Case WM_MCU_VERSION
                InputLog("RECV WM_MCU_VERSION ")
                If EquilModelCode = 4 Or EquilModelCode = 5 Then
                    Dim nVersion1 As Integer = m.WParam.ToInt32
                    Dim nVersion2 As Integer = m.LParam.ToInt32
                    lbMCU1Ver.Text = nVersion1
                    lbMCU2Ver.Text = nVersion2




                End If

            Case WM_DISCONNECTPEN
                If Me.InvokeRequired Then
                    Me.BeginInvoke(New MethodInvoker(AddressOf DisConnectPen))
                Else
                    DisConnectPen()
                End If

                If btnConnectPen.InvokeRequired Then
                    Dim cpd As New connectPenTextDelegate(AddressOf connectPenTextChange)
                    btnConnectPen.BeginInvoke(cpd, "Connect")
                End If

                ''''''' For Memory Import
            Case WM_SHOWLIST
                Console.WriteLine(String.Format("Memory Import WM_SHOWLIST - {0}", m.LParam))
                RaiseEvent ShowDIList()

            Case WM_DI_START
                ' When data in memory exists, 
                Console.WriteLine("Memory Import WM_DI_START =======================")
            Case WM_DI_OK

                Console.WriteLine("Memory Import WM_DI_OK =======================")

            Case WM_DI_FAIL

                Console.WriteLine("Memory Import WM_DI_FAIL =======================")

            Case WM_DI_REMOVEOK

                Console.WriteLine(String.Format("Memory ImportWM_DI_REMOVEOK - folder : {0}, file : {1}", m.WParam, m.LParam))

            Case WM_DI_ACC

                Console.WriteLine("Memory Import  Duplicated =======================")
                ' Should create New page and copied data will come

            Case WM_DI_CHANGE_STATION
                Console.WriteLine("Memory Import WM_DI_CHANGE_STATION =======================")
                ' for smart marker
                ' station is changed

                'PenStateTimerStop()
            Case WM_DI_PAGE_NUM
                Dim nDIPaperSize As Integer = m.LParam 'Paper Size
                Console.WriteLine("WM_DI_PAGE_NUM ==>{0}", nDIPaperSize)
                Dim rect As Rectangle = GetWorkAreaRect(nDIPaperSize)
                SetCalibration_Top(rect.Left, rect.Top)
                SetCalibration_Bottom(rect.Right, rect.Bottom)
                SetCalibration_End()


                'nMakePage = 1
            Case Else
                MyBase.WndProc(m)
        End Select
    End Sub
#End Region

    Public Function GetWorkAreaRect(ByVal nWorkAreaType As Integer) As System.Drawing.Rectangle

        Select Case nWorkAreaType
            Case WorkAreaType.LETTER
                Return Cali_LETTER
            Case WorkAreaType.A4
                Return Cali_A4
            Case WorkAreaType.A5
                Return Cali_A5
            Case WorkAreaType.B5
                Return Cali_B5
            Case WorkAreaType.B6
                Return Cali_B6
            Case WorkAreaType.FT_6X4
                Return Cali_FT_6X4
            Case WorkAreaType.FT_6X5
                Return Cali_FT_6X5
            Case WorkAreaType.FT_8X4
                Return Cali_FT_8X4
            Case WorkAreaType.FT_8X5
                Return Cali_FT_8X5
            Case WorkAreaType.FLIP_7X10
                Return Cali_Flip_7X10
            Case WorkAreaType.TFT_3X5
                Return Cali_TFT_3X5
            Case WorkAreaType.TFT_3X6
                Return Cali_TFT_3X6
            Case WorkAreaType.TFT_4X6
                Return Cali_TFT_4X6
            Case WorkAreaType.BFT_3X5
                Return Cali_BFT_3X5
            Case WorkAreaType.BFT_3X6
                Return Cali_BFT_3X6
            Case WorkAreaType.BFT_4X6
                Return Cali_BFT_4X6
            Case WorkAreaType.MANUAL, WorkAreaType.AUTO
                If EquilModelCode = 4 Or EquilModelCode = 5 Then

                    If CURRENT_MARKER_DIRECT = StationPosition.LEFT Or CURRENT_MARKER_DIRECT = StationPosition.BOTH Or CURRENT_MARKER_DIRECT = StationPosition.RIGHT Then
                        Return Cali_FT_8X5
                    Else
                        Return Cali_TFT_4X6
                    End If
                Else
                    Return Cali_A4
                End If

                
        End Select
    End Function

    Private Sub btnAudio_Click(sender As System.Object, e As System.EventArgs) Handles btnAudio.Click
        SetAudio(CByte(CInt(txtAudioMode.Text)), CByte(CInt(txtAudioVolume.Text)))
        SetDeviceVoiceLanguage(cboAudioLang.Text)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


        'Dim frmCali As New frmCalibration2


        'SetCalibMode(True)
        'SetReciveHandle(frmCali.Handle)
        'If frmCali.ShowDialog() = DialogResult.OK Then
        '    Dim ff2 As New frmDrawing2
        '    Dim tMarginX As Integer = ff2.Width - ff2.ClientSize.Width
        '    Dim tMarginY As Integer = ff2.Height - ff2.ClientSize.Height

        '    ff2.Width = frmCali.m_RectSource.Width + tMarginX
        '    ff2.Height = frmCali.m_RectSource.Height + tMarginY
        '    ff2.SetConfig(m_DrawPen, m_PenStyle)

        '    SetDrawHandle(ff2.picMain.Handle)
        '    SetReciveHandle(ff2.Handle)


        '    SetCalibMode(False)

        '    Dim rect As Rectangle = frmCali.GetTargetRect
        '    Dim rectOrg As Rectangle = frmCali.picMain.ClientRectangle
        '    Dim rectCal As Rectangle = GetRegularRectFromPoint(Point.Round(frmCali.m_CaliTop), Point.Round(frmCali.m_CaliBottom))
        '    Dim tx, ty, bx, by As Integer
        '    Const ABSOLUTE_XY_SIZE As Integer = 65535
        '    tx = ABSOLUTE_XY_SIZE * (rect.X / rectOrg.Width)
        '    ty = ABSOLUTE_XY_SIZE * (rect.Y / rectOrg.Height)

        '    bx = ABSOLUTE_XY_SIZE * (rect.Right / rectOrg.Width)
        '    by = ABSOLUTE_XY_SIZE * (rect.Bottom / rectOrg.Height)
        '    Console.WriteLine("rectOrg={0} rect={1} tx={2}, ty={3}, bx={4}, by={5}", rectOrg, rect, tx, ty, bx, by)
        '    'WORKAREA_RECT_MANUAL_LAST = CUtil.GetRegularRectFromPoint(Point.Round(frmCali.m_CaliTop), Point.Round(frmCali.m_CaliBottom))

        '    'frmGr.SetCaliWorkSheet(frmCali.m_CaliTop.X, frmCali.m_CaliTop.Y, frmCali.m_CaliBottom.X, frmCali.m_CaliBottom.Y)

        '    SetCalibration_Top(rectCal.Left, rectCal.Top)
        '    SetCalibration_Bottom(rectCal.Right, rectCal.Bottom)
        '    SetCalibration_EndWithDest(tx, ty, bx, by)

        '    ff2.ShowDialog()
        '    SetCalibMode(False)
        '    SetReciveHandle(Me.Handle)
        '    SetDrawHandle(picMain.Handle)


        'Else
        '    SetCalibMode(False)
        '    SetReciveHandle(Me.Handle)
        '    SetDrawHandle(picMain.Handle)
        'End If



    End Sub

    Private Sub btnMemoryImport_Click(sender As Object, e As EventArgs) Handles btnMemoryImport.Click
        If bPenConnect Then
            m_frmMemory = Nothing
            m_frmMemory = New frmMemoryImport
            m_frmMemory.Show()
        Else
            MessageBox.Show("Either device is off or not in good connection. Check device and try again. Or pair it again.")
        End If
    End Sub


    Private Sub cboAudioLang_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboAudioLang.SelectedIndexChanged

    End Sub
End Class
