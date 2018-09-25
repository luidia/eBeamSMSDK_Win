Imports System.Runtime.InteropServices

Module Module1
    Enum PenStatus
        PEN_DOWN = 1
        PEN_MOVE
        PEN_UP
        PEN_HOVER
    End Enum

    Enum PenCalibrationStatus
        CAL_TOP = 0
        CAL_BOTTOM
        CAL_END
    End Enum

    Public Enum StationPosition
        NONE = 0
        TOP = 1
        LEFT
        RIGHT
        BOTTOM
        BOTH
    End Enum

    Public Enum MakerState
        RED = 81
        GREEN = 82
        YELLOW = 83
        BLUE = 84
        PEN_UP = 15
        PURPLE = 86
        BLACK = 88
        ERASER_CAP = 89
        LOW_BATTERY = 90
        BIG_ERASER = 92
    End Enum

    Enum PENSTYLE
        PENSTYLE_NORMAL = 0
        PENSTYLE_BRUSH
        PENSTYLE_HIGHLIGHT
    End Enum

    Enum GestureMessage
        GESTURE_RIGHT_LEFT = 100
        GESTURE_LEFT_RIGHT
        GESTURE_UP_DOWN
        GESTURE_DOWN_UP
        GESTURE_CIRCLE_CLOCKWISE
        GESTURE_DOUBLECIRCLE_CLOCKWISE
        GESTURE_CIRCLE_COUNTERCLOCKWISE
        GESTURE_CROSS_DOWN
        GESTURE_CROSS_UP
        GESTURE_CROSS_LEFT
        GESTURE_CROSS_RIGHT
        GESTURE_ZIGZAG
        GESTURE_CLICK
        GESTURE_DOUBLECLICK
        GESTURE_TRIPPLECLICK
        GESTURE_LONGCLICK
        GESTURE_DOUBLECROSS_DOWN
        GESTURE_DOUBLECROSS_UP
        GESTURE_DOUBLECROSS_LEFT
        GESTURE_DOUBLECROSS_RIGHT
        GESTURE_DOUBLECIRCLE_COUNTERCLOCKWISE
    End Enum

    Public Const WM_RETURNMESSAGE As Integer = &H400 + 1        'pen message
    Public Const WM_GESTUREMESSAGE As Integer = &H400 + 2       'gesture message
    Public Const WM_PENCONDITION As Integer = &H400 + 3         'gesture message
    Public Const WM_PENCALIBRATION As Integer = &H400 + 21         'gesture message
    Public Const WM_ERASER_ON As Integer = &H400 + 4

    Public Const WM_SHOWLIST As Integer = &H400 + 5             'di show list
    Public Const WM_DI_START As Integer = &H400 + 6             'di start
    Public Const WM_DI_OK As Integer = &H400 + 7                'di ok
    Public Const WM_DI_FAIL As Integer = &H400 + 8              'di fail
    Public Const WM_DI_REMOVEOK As Integer = &H400 + 9          'di remove ok
    Public Const WM_DI_ACC As Integer = &H400 + 10              'di acc
    Public Const WM_DISCONNECTPEN As Integer = &H400 + 11       'Disconnect pen
    Public Const WM_DOWNLOAD_COMPLEATE As Integer = &H400 + 12  'di download complete
    Public Const WM_DI_NEWPAGE As Integer = &H400 + 13          'New Page
    Public Const WM_DI_PAGE_NUM As Integer = &H400 + 14        'di duplicate
    Public Const WM_DI_CHANGE_STATION As Integer = &H400 + 15        'di duplicate
    Public Const WM_DISCONNECTUSB As Integer = &H400 + 16        'Disconnect Usb
    Public Const WM_DI_DUPLICATE As Integer = &H400 + 17        'Duplicate Button event
    Public Const WM_MCU_VERSION As Integer = &H400 + 18        'Duplicate Button event
    Public Const WM_TIMERRESET As Integer = &H400 + 19        'Duplicate Button event
    Public Const WM_LOST_PEN As Integer = &H400 + 20        'Disconnect by timer cannot receive pendata during 30sec

    Public Enum DICOMMAND
        SHOWLIST = 1
        OPENFILE
        OPENFOLDER
        REMOVEFILE
        REMOVEFOLDER
        REMOVEALL
        SETDATETIME
        SHOWDATETIME
        SHOWDISKFREESPACE
        SHOWDEVICEID
        SHOWTEMPFILE
        CHANGEMODETOREAL
        CHANGEMODETOT2
    End Enum
    Public Enum WorkAreaType
        AUTO = 0
        LETTER
        A4
        A5
        B5
        B6
        FT_6X4 = 6
        FT_6X5
        FT_8X4
        FT_8X5
        MANUAL = 10
        TFT_3X5 = 11
        TFT_3X6
        TFT_4X6
        BFT_3X5 = 14
        BFT_3X6
        BFT_4X6

        FLIP_7X10
        MANUAL_FORM
    End Enum
    Public Cali_LETTER As New System.Drawing.Rectangle(1737, 541, (5445 - 1737), (4818 - 541))   '(1700, 500, 5470, 4800)
    Public Cali_A4 As New System.Drawing.Rectangle(1768, 563, (5392 - 1768), (5160 - 563))       '(1750, 450, 5450, 5120)
    Public Cali_A5 As New System.Drawing.Rectangle(2341, 542, (4865 - 2341), (3631 - 542))       '(2300, 540, 4880, 3625)
    Public Cali_B5 As New System.Drawing.Rectangle(2027, 561, (5183 - 2027), (4462 - 561))       '(2000, 500, 5200, 4430)
    Public Cali_B6 As New System.Drawing.Rectangle(2500, 544, (4704 - 2500), (3154 - 544))       '(2460, 530, 4660, 3170)

    Public Cali_FT_6X4 As New System.Drawing.Rectangle(1728, 45372, (15524 - 1728), (54824 - 45372))
    Public Cali_FT_6X5 As New System.Drawing.Rectangle(1830, 44156, (15506 - 1830), (56034 - 44156))
    Public Cali_FT_8X4 As New System.Drawing.Rectangle(1868, 45377, (20153 - 1868), (54735 - 45377))
    Public Cali_FT_8X5 As New System.Drawing.Rectangle(1450, 44100, (20620 - 1450), (56150 - 44100)) 'ebeam sm

    Public Cali_Flip_7X10 As New System.Drawing.Rectangle(1505, 46453, 5054, 7839)

    Public Cali_TFT_3X5 As New System.Drawing.Rectangle(12790, 1547, (19966 - 12790), (13248 - 1547))
    Public Cali_TFT_3X6 As New System.Drawing.Rectangle(12790, 1532, (19966 - 12790), (15298 - 1532))
    Public Cali_TFT_4X6 As New System.Drawing.Rectangle(11590, 1450, (21230 - 11590), (15827 - 1450)) 'ebeam sm

    Public Cali_BFT_3X5 As New System.Drawing.Rectangle(46612, 53961, (53788 - 46612), (65662 - 53961))
    Public Cali_BFT_3X6 As New System.Drawing.Rectangle(46612, 51800, (53788 - 46612), (65566 - 51800))
    Public Cali_BFT_4X6 As New System.Drawing.Rectangle(45305, 50708, (54945 - 45305), (65086 - 50708)) 'ebeam sm

    Public MINI_MANUAL_CAL As Integer = Cali_FT_8X5.Width / 8 '1 ft

    '' Structure for receiving pen data 
    Friend Structure _pen_rec
        Public X As Integer ' X (before calibration)
        Public Y As Integer ' Y (before calibration)
        Public T As Short ' temperature (celcious)
        Public P As Integer ' pressure ( 0 ~ 900 )
        Public TX As Single '  X (after calibration)
        Public TY As Single ' Y (after calibration)
        Public FUNC As Integer 'pen button clicked
        Public ModelCode As Integer 'Device Model Code 2: Equil Smart Pen 3:Equil Smart Pen 2 4:Equil Smart Marker   
        '        Public Sensor_dis As Integer 'distance between sensors (need not to know) 
        '        Public HWVer As Integer 'HWVer 
        '        Public MCU1 As Integer 'MCU1 (need not to know) 
        '        Public MCU2 As Integer 'MCU2 (need not to know)
        Public PenStatus As Integer 'Pen tip (1: Down 2:Move 3:Up 4:Hover)
        '        Public IRGAP As Integer 'sensor property (need not to know) 
        'SmartMaker Variable Add
        Public PenTiming As Integer 'Maker State Data
        Public bRight As Integer
        Public Station_Position As Integer
        '       Public drawRectX As Integer
        '        Public drawRectY As Integer
    End Structure
    Public Structure PENConditionData
        Public modelCode As Integer
        Public pen_alive As Integer 'Pen alive
        Public battery_station As Integer
        Public battery_pen As Integer
        Public StationPosition As Integer
        Public usbConnect As Integer
    End Structure
    '    Public Declare Function FindPort Lib "PNFPenLib.dll" () As Boolean 'scan port from device
    Public Declare Sub OnDisconnect Lib "PNFPenLib.dll" (Terminate As Integer)
    Public Declare Function FindDevice Lib "PNFPenLib.dll" () As Boolean   'Search for USB connection
    Public Declare Sub PortSearch Lib "PNFPenLib.dll" ()
    Public Declare Sub SetReciveHandle Lib "PNFPenLib.dll" (ByVal hWnd As System.IntPtr)
    Public Declare Sub SetDrawRect Lib "PNFPenLib.dll" (ByVal Width As Double, ByVal Height As Double)
    Public Declare Sub SetDrawHandle Lib "PNFPenLib.dll" (ByVal hWnd As System.IntPtr)
    Public Declare Function CheckConnect Lib "PNFPenLib.dll" () As Boolean
    Public Declare Function SetSMPosition Lib "PNFPenLib.dll" (ByVal position As StationPosition, bSaveDevice As Boolean) As Integer
    Public Declare Sub SetUserLang Lib "PNFPenLib.dll" (ByVal langCode As Integer)
    Public Declare Sub SetCalibMode Lib "PNFPenLib.dll" (ByVal bCalibMode As Boolean, bSaveDevice As Boolean)
    Public Declare Sub SetPenDownThreshold Lib "PNFPenLib.dll" (ByVal iDown As Integer)
    Public Declare Sub SetCalibration_Top Lib "PNFPenLib.dll" (ByVal x As Double, ByVal y As Double)
    Public Declare Sub SetCalibration_Bottom Lib "PNFPenLib.dll" (ByVal x As Double, ByVal y As Double)
    Public Declare Sub SetCalibration_End Lib "PNFPenLib.dll" ()
    Public Declare Sub SaveCalibrationToDevice Lib "PNFPenLib.dll" (x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, x3 As Integer, y3 As Integer)

    'for calibration by form
    Public Declare Sub SetCalibration_EndWithDest Lib "PNFPenLib.dll" (tx As Integer, ty As Integer, bx As Integer, by As Integer)

    Public Declare Sub SetAudio Lib "PNFPenLib.dll" (ByVal _AudioMode As Byte, ByVal _AudioVolume As Byte)
    Public Declare Function GetAudioMode Lib "PNFPenLib.dll" () As Byte
    Public Declare Function GetAudioVolume Lib "PNFPenLib.dll" () As Byte
    Public Declare Function GetAudioLang Lib "PNFPenLib.dll" () As Byte


    ''' <summary>
    ''' Memory Import Apis '''''
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Public Declare Function StartDISetup Lib "PNFPenLib.dll" (ByVal nState As Integer) As Integer '' Intialize nState =0:BT, nSate=1 USB
    Public Declare Sub SetDI Lib "PNFPenLib.dll" (ByVal nState As Integer, ByVal nFolder As Integer, ByVal nFile As Integer)
    Public Declare Sub SetDIByte Lib "PNFPenLib.dll" (ByVal commandBytes As Byte())
    Public Declare Function GetDIList Lib "PNFPenLib.dll" (ByVal byteArray As IntPtr, ByRef _size As Integer) As Integer 'get file list
    Public Declare Function DIOpenFileStop Lib "PNFPenLib.dll" () As Integer


    Public EquilModelCode As Integer = 2
    Public CURRENT_MARKER_DIRECT As Integer


    Public Sub SetDeviceVoiceLanguage(langStr As String)
        Dim langCode As Integer
        Select Case langStr
            Case "Korean"
                langCode = 1
            Case "Japanese"
                langCode = 6
            Case "Spanish"
                langCode = 2
            Case "German"
                langCode = 3
            Case "French"
                langCode = 4
            Case "Italian"
                langCode = 5
            Case "Portuguese"
                langCode = 7
            Case "Russian"
                langCode = 8
            Case "Chinese1", "Chinese2"
                langCode = 9
            Case "Arabic"
                langCode = 11
            Case "English"
                langCode = 0
            Case Else
                langCode = 0
        End Select
        SetUserLang(langCode)

    End Sub
    Public Function GetDeviceVoiceLanguage(langCode As Integer) As String

        Select Case langCode
            Case 1 : Return "Korean"
            Case 6 : Return "Japanese"
            Case  2 : Return "Spanish"
            Case 3 : Return "German"
            Case 4 : Return "French"
            Case 5 : Return "Italian"
            Case 7 : Return "Portuguese"
            Case 8 : Return "Russian"
            Case 9 : Return "Chinese"
            Case 11 : Return "Arabic"
            Case 10 : Return "Spanish"
            Case 11 : Return "Spanish"
            Case 0 : Return "English"



        End Select

        Return ""
    End Function

End Module
