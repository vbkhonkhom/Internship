Imports System.Data.SqlClient
Imports System.Net
Imports System.Text
Imports System.IO

' [Form1] ศูนย์ควบคุมหลักของโปรแกรม
Public Class Form1
    ' --- ส่วนประกาศตัวแปร Global (ใช้ได้ทั่วทั้งโปรแกรม) ---
    
    ' StrServerConnection_RIST: เก็บ Connection String (ที่อยู่, user, password) สำหรับต่อฐานข้อมูลต้นทาง (RIST)
    Public StrServerConnection_RIST As String = ""
    
    ' StrServerConnection_Yuku: เก็บ Connection String สำหรับต่อฐานข้อมูลปลายทาง (Yukuhashi)
    Public StrServerConnection_Yuku As String = ""
    
    ' StrStartDate: เก็บวันที่เริ่มต้น (Start Date) ที่จะเริ่มดึงข้อมูล (รูปแบบ String)
    Public StrStartDate As String = ""
    
    ' StrCDir: เก็บ Path ของโฟลเดอร์ปัจจุบันที่โปรแกรมรันอยู่ (เอาไว้หาไฟล์ Config หรือ Log)
    public StrCDir As String = System.IO.Directory.GetCurrentDirectory
    
    ' OKFlag: ตัวแปร Boolean (True/False) เอาไว้เช็คสถานะรวมๆ ว่าการทำงานราบรื่นหรือไม่
    Public OKFlag As Boolean

    ' Form Load: เริ่มต้นทำงานเมื่อเปิดโปรแกรม
    Private Sub Form1_Load(ByVal sender As System.Object, ByVale As System.EventArds) Handles MyBase.Load
        Button2.Enabled = False
        Button4.Enabled = False
        ReadConfigData() ' อ่านค่า Config เก่า
        SaveConfigData() ' บันทึกซ้ำ (Update ไฟล์)
        ' My.Settings.DisplayNotify: ค่า Setting ของโปรแกรมว่าให้แสดงปุ่มแจ้งเตือนไหม
        If My.Settings.DisplayNotify Then Button5.Visible = True
    End Sub

    ' อ่านไฟล์ Config.csv
    Private Sub ReadConfigData()
        ' strFileName: ตัวแปร Local เก็บชื่อไฟล์ Config ที่จะอ่าน
        Dim strFileName As String = strCDir & "\Config.csv"
        
        ' temp(): อาร์เรย์ String ใช้พักข้อมูลแต่ละบรรทัดที่อ่านได้ (แยกด้วย comma)
        Dim temp() As String
        
        ' sr: Object สำหรับอ่านไฟล์ Text
        Dim sr As New System.IO.StreamReader(strFileName, System.Text.Encoding.Default)
        
        Do Until sr.Peek() = -1
            temp = Split(sr.ReadLine(), ",")
            ' temp(0) คือชื่อตัวแปร (Key), temp(1) คือค่า (Value)
            If temp(0).Trim(Chr(34)) <> "" Then
                If temp(0) = "Server_RIST" Then
                    StrServerConnection_RIST = temp(1)
                ElseIf temp(e) = "Server_Yukuhashi" Then
                    StrServerConnection_Yuku = temp(1)
                ElseIf temp(0) = "StartDate" Then
                    StrStartDate = temp(1)
                End If
            End If
        Loop
        sr.Close()

        ' ถ้าไม่มีวันที่กำหนดไว้ ให้ Default เป็น 2020-01-01
        If StrStartDate = "" Then
            StrStartDate = "2020-01-01 00:00:00.000"
            Form2.Button2.Enabled = False
            Button3_Click(Nothing, Nothing)
        End If
    End Sub

    ' บันทึก Config ลงไฟล์
    Public Sub SaveConfigData()
        Dim str1 As String
        Dim strFileName As String = strCDir & "\Config-csv"
        ' sw1: Object สำหรับเขียนไฟล์ Text
        Dim sw1 As New System.IO.StreamWriter(strFileName, False, System.Text.Encoding.Default)
        str1 = Err.Description
        If str1 = "" Then
            sw1.WriteLine(",Connection string")
            sw1.WriteLine("Server_RIST," & StrServerConnection_RIST)
            sw1.WriteLine("Server_Yukuhashi," & StrServerConnection_Yuku)
            sw1.WriteLine("StartDate," & StrStartDate)
            sw1.Close()
        End If
    End Sub

    ' หา ID ล่าสุดในตาราง SPC_Master
    Public Function Get_NextMasterNo() As String
        ' Cn: ตัวจัดการการเชื่อมต่อ Database
        Dim Cn As New System.Data.SqlClient.SqlConnection
        ' Adapter: ตัวกลางดึงข้อมูลจาก DB มาใส่ DataTable
        Dim Adapter As New SqlDataAdapter
        ' strSQL: เก็บคำสั่ง SQL Query (SELECT ...)
        Dim strSQL As String = ""
        ' n: เก็บจำนวนแถวที่ได้
        Dim n As Integer
        ' table: ตารางจำลองในหน่วยความจำ (DataTable) เก็บผลลัพธ์ที่ดึงมา
        Dim table As New DataTable
        Try
            Get_NextMasterNo = ""
            Cn.ConnectionString = StrServerConnection_Yuku
            table.Clear()
            strSQL = "SELECT iID"
            strSQL &= "FROM SPC_Master"
            
            Adapter = New SqlDataAdapter()
            Adapter.SelectCommand = New SqlCommand(strSQL, Cn)
            Adapter.SelectCommand.CommandType = CommandType.Text
            Adapter.Fill(table)
            n = table.Rows.Count

            Adapter.Dispose()
            Cn.Dispose()

            If n = 0 Then
                table.Dispose()
                Get_NextMasterNo = 1
                Exit Function
            End If

            Get_NextMasterNo = table.Rows(n - 1)(0) + 1
        Catch ex As System.Exception
            Get_NextMasterNo = ""
            Adapter.Dispose()
            Cn.Dispose()
            MsgBox("ErrorCode " + eCode + " , " + ex.Message & ex.StackTrace)
            Exit Function
        End Try
    End Function

    ' ฟังก์ชันหา Max ID อีกเวอร์ชัน
    Public Function Get_NextMasterNo2(ByVal Tname As String, ByVal StrConnectionString As String) As String
        eCode = Tname & "Get_NextMasterNo2"
        Get_NextMasterNo2 = -1
        Dim Cn As New SqlConnection
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim strSQL As String = ""
        Cn.ConnectionString = StrConnectionString
        SQLCm.CommandTimeout = 600
        strSQL = "SELECT MAX(iID) FROM " & Tname
        Cn.Open()
        SQLCm.CommandText = strSQL
        ' ExecuteScalar: ดึงค่าเดียวโดดๆ (ในที่นี้คือค่า MAX)
        Get_NextMasterNo2 = SQLCm.ExecuteScalar
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Function

    ' ฟังก์ชัน Insert ข้อมูลจำนวนมากลง Server
    Public Sub To_Servere(ByVal Array0(,) As String, ByVal Array1(,) As String)
        Dim Cn As New SqlConnection
        Dim strSQL As String
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        ' trans: ตัวจัดการ Transaction (เพื่อให้แน่ใจว่าถ้า Insert พลาด ต้องไม่เข้าเลยสักแถว - Rollback)
        Dim trans As SqlTransaction = Nothing
        
        Try
            Cn.ConnectionString = StrServerConnection_Yuku
            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans

            For i As Integer = 0 To UBound(Array0, 1)
                strSQL = ""
                strSQL = "INSERT INTO " & "SPC_Master" & " VALUES ("
                
                For j As Integer = 0 To UBound(Array0, 2)
                    If Array0(i, j) = "Null" Then
                        strSQL &= "Null"
                    Else
                        strSQL &= "'" & Array0(i, j) & "'"
                    End If
                    strSQL &= ","
                Next
                For j As Integer = 0 To UBound(Array1, 2)
                    If Array1(i, j) = "Null" Then
                        strSQL &= "Null"
                    Else
                        strSQL &= "'" & Array1(i, j) & "'"
                    End If
                    If j = UBound(Array1, 2) Then
                        strSQL &= ")"
                    Else
                        strSQL &= ","
                    End If
                Next
                SQLCm.CommandText = strSQL
                SQLCm.ExecuteNonQuery()
            Next
            trans.Commit()
            Cn.Close()

        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback() ' ย้อนกลับข้อมูลถ้า error
            End If
            OKflag = False
            MsgBox("ErrorCode " + eCode + "," + ex.Message & ex.StackTrace)
        End Try
    End Sub

    ' eCode: ตัวแปร String เก็บข้อความ Error ล่าสุด (ใช้ Debug)
    Dim eCode As String
    ' Dflag: ตัวแปร Flag บอกสถานะว่ากำลังเขียนข้อมูลอยู่ไหม
    Public Dflag As Boolean = False
    
    ' ตัวแปรเกี่ยวกับเวลา (ใช้จับเวลาการทำงาน)
    Dim SDt_All As DateTime ' เวลาเริ่มงานทั้งหมด
    Dim eDt_All As DateTime ' เวลาจบงานทั้งหมด
    Dim SDt As DateTime     ' เวลาเริ่มงานย่อย
    Dim eDt As DateTime     ' เวลาจบงานย่อย
    Dim ts(6 - 1) As TimeSpan ' อาร์เรย์เก็บระยะเวลา (TimeSpan) ของแต่ละขั้นตอน
    Dim ts_All As TimeSpan    ' ระยะเวลารวมทั้งหมด
    
    ' StopFlag: ตัวแปรบอกให้ BackgroundWorker หยุดทำงาน
    Public StopFlag As Boolean
    
    ' eMsg: เก็บข้อความ Error Message
    Dim eMsg As String = ""
    ' Write_times: ตัวนับจำนวนรอบการเขียน (อาจเอาไว้รีเซ็ตค่าเมื่อครบจำนวน)
    Dim Write_times As Integer = 0

    ' [ปุ่ม Start] เริ่มการแปลงข้อมูล
    Private Sub Butto1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form4.Close()
        If StrServerConnection_RIST = "" Then
            MsgBox("Set the RIST server connection string.")
            Exit Sub
        End If
        If StrServerConnection_Yuku = "" Then
            MsgBox("Set the Yukuhashi server connection string.")
            Exit Sub
        End If

        Write_times = 0
        StopFlag = False
        OKflag = True
        Me.Hide()
        Form3.Show()
        Form3.TopMost = True

        ' สั่งเริ่ม BackgroundWorker (ทำงานแยก Thread ไม่ให้โปรแกรมค้าง)
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker1.WorkerSupportsCancellation = True
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    ' หา "วันที่ทำงาน" ล่าสุดที่มีในระบบปลายทาง (เพื่อดึงต่อจากวันนั้น)
    Public Function Get_Recently_Day() As String
        eCode = "a Get_Recently_Day"
        Dim Cn As New SqlConnection
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim Adapter As New SqlDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim strSQL As String = ""
        Cn.ConnectionString = StrServerConnection_Yuku
        SQLCm.CommandTimeout = 600
        strSQL = "SELECT MAX(dworkDate) FROM SPC_Master"
        strSQL &= " WHERE dworkDate < '2292-01-01 00:00:00.000'"
        SQLCm.CommandText = strSQL
        Adapter.Fill(Table)
        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        eCode = "b Get_Recently_Day"
        
        ' เช็คว่าได้ค่ามาไหม ถ้าได้ก็ Return กลับไป
        If Not IsDBNull(Table.Rows(0)(0)) Then
            Get_Recently_Day = Format(Table.Rows(0)(0), "yyyy-MM-dd HH:mm:ss.fff")
            eCode = "c Get_Recently_Day"
            If Get_Recently_Day < StrStartDate Then
                Get_Recently_Day = StrStartDate
            End If
        Else
            Get_Recently_Day = Nothing
            Get_Recently_Day = StrStartDate
        End If
        eCode = "d Get Recently_Day"
    End Function

    ' ดึงข้อมูล BATCH_TBL ใส่ _TableB
    Public Sub Get_Data_BATCH_TBL(ByRef _TableB As DataTable, ByVal _gRows As Integer, ByVal _recDay As String)
        eCode = "a Get_Date_BATCH_TBL"
        _TableB.Clear()
        Dim Cn As New SqlConnection
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim Adapter As New SqlDataAdapter(SQLCm)
        Dim Table As New DataTable
        Dim strSQL As String = ""

        Cn.ConnectionString = StrServerConnection_RIST
        SQLCm.CommandTimeout = 600
        strSQL = "SELECT TOP " & _gRows
        strSQL &= " BATCH_ID,"
        strSQL &= "DATA_ID, "
        strSQL &= "START_TIME, "
        strSQL &= "RECIPE, "
        strSQL &= "EQP_NO"
        strSQL &= " FROM BATCH_TBL"
        If Not _recDay = Nothing Then
            strSQL &= " WHERE START_TIME > " & "'" & _recDay & "'"
            strSQL &= " AND START_TIME < '2292-01-01 00:00:00.000'"
        Else
            strSQL &= " WHERE START_TIME < '2292-01-01 00:00:00.000'"
        End If
        strSQL &= " ORDER BY START_TIME ASC"
        eCode = "b Get_Data_BATCH_TBL" & Environment.NewLine & strSQL
        SQLCm.CommandText = strSQL
        Adapter.Fill(_TableB)
        _TableB.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()
        eCode = "c Get_Data_BATCH_TBL" & strSQL
    End Sub

    ' ดึงข้อมูล LOT_TBL และ DATA_TBL
    ' _TableB: ตาราง Batch ที่ดึงมาแล้ว
    ' _TableL: ตาราง Lot ที่จะดึงมาใส่
    ' _TableD: ตาราง Data (Measurement) ที่จะดึงมาใส่
    ' _BidList: List เก็บ Batch ID ทั้งหมด
    ' _DidList: List เก็บ Data ID ทั้งหมด
    Public Sub Get_Data_LOT_DATA_TBL(ByVal _TableB As DataTable, ByRef _TableL As DataTable, ByRef _TableD As DataTable, ByRef _Bidlist As List(Of String), ByRef _DidList As List(Of String), ByRef _F_BATCHID As String, ByRef _E_BATCHID As String)
        eCode = "a Get_Data_LOT_DATA_TBL"
        _Bidlist.Clear()
        _DidList.Clear()
        sDt = DateTime.Now
        
        ' วนลูปแยก Batch ID และ Data ID จาก _TableB ลง List
        For i As Integer = 0 To _TableB.Rows.Count - 1
            If Not IsDBNull(_TableB.Rows(i)("BATCH_ID")) Then
                _Bidlist.Add(_TableB.Rows(i)("BATCH_ID"))
            End If
            If Not IsDBNull(_TableB.Rows(i)("DATA_ID")) Then
                _DidList.Add(_TableB.Rows(i)("DATA_ID"))
            End If
        Next
        
        ' ตรวจสอบความถูกต้องของจำนวนข้อมูล
        If Not _Bidlist.Count = _DidList.Count Then
            eCode = "b Get_Data_LOT_DATA_TBL Not _BidList.Count = _DidList.Count " & _Bidlist.Count & "  " & _DidList.Count
            Dim g As Integer = "w"
        End If
        ' ... (ตรวจสอบ Integrity ของข้อมูล) ...

        eCode = "d Get_Data_LOT_DATA_TBL"
        _F_BATCHID = _Bidlist(0) ' Batch ID แรก
        _E_BATCHID = _Bidlist(_Bidlist.Count - 1) ' Batch ID สุดท้าย
        _TableL.Clear()
        
        ' --- Query LOT_TBL โดยใช้ WHERE IN (หรือ OR หลายๆตัว) จาก Batch ID ---
        Dim CnL As New SqlConnection
        Dim SQLCmL As SqlCommand = CnL.CreateCommand
        Dim AdapterL As New SqlDataAdapter(SQLCmL)
        Dim strSQLL As String = ""
        CnL.ConnectionString = StrServerConnection_RIST
        SQLCmL.CommandTimeout = 600

        strSQLL = "SELECT "
        strSQLL &= "BATCH_ID,"
        strSQLL &= "LOT_NO,"
        strSQLL &= "PRODUCT, "
        strSQLL &= "PROCESS, "
        strSQLL &= "OPE_NAME,"
        strSQLL &= "L_PARAM2"
        strSQLL &= " FROM LOT_TBL"
        strSQLL &= " WHERE "
        sDt = DateTime.Now
        For i As Integer = 0 To _Bidlist.Count - 1
            strSQLL &= "BATCH_ID = " & " '" & _Bidlist(i) & "'"
            If Not i = _Bidlist.Count - 1 Then
                strSQLL &= " OR "
            End If
        Next
        strSQLL &= " ORDER BY BATCH_ID ASC"
        eDt = DateTime.Now
        eCode = "e Get_Data_LOT_DATA_TBL" & Environment.NewLine & strSQLL
        SQLCmL.CommandText = strSQLL
        AdapterL.Fill(_TableL)
        _TableL.Dispose()
        AdapterL.Dispose()
        SQLCmL.Dispose()
        CnL.Dispose()
        eDt = DateTime.Now

        ' --- Query DATA_TBL (ข้อมูลวัดผล) โดยใช้ Data ID ---
        eCode = "f Get_Data_LOT_DATA_TBL " & Environment.NewLine & strSQLL
        _TableD.Clear()
        Dim CnD As New SqlConnection
        Dim SQLCmD As SqlCommand = CnD.CreateCommand
        Dim AdapterD As New SqlDataAdapter(SQLCmD)
        Dim strSQLD As String = ""
        CnD.ConnectionString = StrServerConnection_RIST
        SQLCmD.CommandTimeout = 600

        strSQLD = "SELECT "
        strSQLD &= "DATA_ID,"
        strSQLD &= "MEASURE_ITEM,"
        strSQLD &= "MEASURE_POINT, "
        strSQLD &= "DATA "
        strSQLD &= " FROM DATA_TBL"
        strSQLD &= " WHERE "
        For i As Integer = 0 To _DidList.Count - 1
            strSQLD &= "DATA_ID = " & " '" & _DidList(i) & "'"
            If Not i = _DidList.Count - 1 Then
                strSQLD &= " OR "
            End If
        Next
        strSQLD &= " ORDER BY DATA_ID ASC"
        eCode = "g Get_Data_LOT_DATA_TBL " & Environment.NewLine & strSQLD
        SQLCmD.CommandText = strSQLD
        AdapterD.Fill(_TableD)
        _TableD.Dispose()
        AdapterD.Dispose()
        SQLCmD.Dispose()
        CnD.Dispose()
        eCode = "h Get_Data_LOT_DATA_TBL " & Environment.NewLine & strSQLD

        ' (ส่วน Logic ตรวจสอบว่า Batch ID ใน List กับใน Table ตรงกันไหม ถ้าไม่ตรงให้ลบออก)
        Dim buf(_TableL.Rows.Count - 1)
        Dim flag(_Bidlist.Count - 1) As Boolean
        If _Bidlist.Count = _TableL.Rows.Count Then
        ElseIf _Bidlist.Count > _TableL.Rows.Count Then
            ' ... (Clean ข้อมูลที่ไม่สมบูรณ์) ...
            eCode = "cc Get_Data_LOT_DATA_TBL_BidList.Count > _TableL.Rows.Count " & _Bidlist.Count & " & " & _TableL.Rows.Count & " " & _F_BATCHID & " " & _E_BATCHID
            For i As Integer = 0 To UBound(buf, 1)
                buf(i) = _TableL.Rows(i)("BATCH_ID")
            Next

            For i As Integer = 0 To _Bidlist.Count - 1
                flag(i) = False
                For j As Integer = 0 To UBound(buf, 1)
                    If _Bidlist(i) = buf(j) Then
                        flag(i) = True
                        Exit For
                    End If
                Next
            Next
            For i As Integer = _Bidlist.Count - 1 To 0 Step -1
                If flag(i) = False Then
                    _Bidlist.RemoveAt(i)
                    _DidList.RemoveAt(i)
                End If
            Next
        ElseIf _Bidlist.Count < _TableL.Rows.Count Then
            eCode = "cc Get_Data_LOT_DATA_TBL_BidList.Count < _Tablel.Rows.Count " & _Bidlist.Count & " " & _TableL.Rows.Count & " " & _F_BATCHID & " " & _E_BATCHID
            Dim g As Integer = "W"
        End If
    End Sub

    ' [Logic หลัก] เตรียมข้อมูลลง Buffer (_DataBuf) และคำนวณ
    ' _DataBuf: อาร์เรย์ 2 มิติที่สำคัญที่สุด เก็บข้อมูลที่พร้อมจะบันทึกลง DB
    Public Sub Get_Data_Buf(ByRef _TableB As DataTable, ByRef _TableL As DataTable, ByRef _TableD As DataTable, ByRef _DataBuf(,) As String, ByVal _MaxID As Integer, ByVal _i As Integer, ByVal _BidList As List(Of String), ByVal _DidList As List(Of String))
        BackgroundWorker1.ReportProgress(6)
        
        ' ReDim _DataBuf: ปรับขนาดอาร์เรย์ให้พอดีกับจำนวน Batch (แถว) และ 126 คอลัมน์
        ReDim _DataBuf(_BidList.Count - 1, 126 - 1)
        
        ' เคลียร์ค่าเริ่มต้นเป็น "Null"
        For i As Integer = 0 To UBound(_DataBuf, 1)
            For j As Integer = 0 To UBound(_DataBuf, 2)
                _DataBuf(i, j) = "Null"
            Next
        Next
        
        ' Count_Data: ตัวนับว่าแต่ละ Batch มีข้อมูล Measurement กี่ตัวแล้ว
        Dim Count_Data(_DidList.Count - 1) As Integer
        sDt = DateTime.Now
        
        ' buf0: พักข้อมูลจาก TableB (Batch)
        Dim buf0(_TableB.Rows.Count - 1, 4 - 1) As String
        For j As Integer = 0 To UBound(buf0, 1)
            eCode = "a Get Data_Buf " & j & " " & UBound(buf0, 1)
            buf0(j, 0) = "Null"
            buf0(j, 1) = "Null"
            buf0(j, 2) = "Null"
            buf0(j, 3) = "Null"
            ' ดึง BatchID, EqpNo, Recipe, StartTime
            If Not IsDBNull(_TableB.Rows(j)("BATCH_ID")) Then buf0(j, 0) = _TableB.Rows(j)("BATCH_ID")
            If Not IsDBNull(_TableB.Rows(j)("EQP_NO")) Then buf0(j, 1) = _TableB.Rows(j)("EQP_NO")
            If Not IsDBNull(_TableB.Rows(j)("RECIPE")) Then buf0(j, 2) = _TableB.Rows(j)("RECIPE")
            If Not IsDBNull(_TableB.Rows(j)("START_TIME")) Then buf0(j, 3) = Format(_TableB.Rows(j)("START_TIME"), "yyyy-MM-dd HH:mm:ss.fff")
        Next
        
        ' buf1: พักข้อมูลจาก TableL (Lot)
        Dim buf1(_TableL.Rows.Count - 1, 5 - 1) As String
        For j As Integer = 0 To UBound(buf1, 1)
            eCode = "b Get Data_Buf " & j & " " & UBound(buf1, 1)
            ' ... (กำหนดค่า Null) ...
            ' ดึง Process, LotNo, Param2, OpeName
            If Not IsDBNull(_TableL.Rows(j)("BATCH_ID")) Then buf1(j, 0) = _TableL.Rows(j)("BATCH_ID")
            If Not IsDBNull(_TableL.Rows(j)("PROCESS")) Then buf1(j, 1) = _TableL.Rows(j)("PROCESS")
            If Not IsDBNull(_TableL.Rows(j)("LOT_NO")) Then buf1(j, 2) = _TableL.Rows(j)("LOT_NO")
            If Not IsDBNull(_TableL.Rows(j)("L_PARAM2")) Then buf1(j, 3) = _TableL.Rows(j)("L_PARAM2")
            If Not IsDBNull(_TableL.Rows(j)("OPE_NAME")) Then buf1(j, 4) = _TableL.Rows(j)("OPE_NAME")
        Next
        
        ' เคลียร์ DataTable เพื่อคืน Memory
        _TableB.Clear()
        _TableL.Clear()
        
        ' --- วนลูปจับคู่ข้อมูล (Matching) ลง _DataBuf ---
        For i As Integer = 0 To _BidList.Count - 1
            ' จับคู่ Batch Info
            For j As Integer = 0 To UBound(buf0, 1)
                eCode = "c Get_Data_Buf " & i & " " & _BidList.Count - 1 & " " & UBound(buf0, 1)
                If Not buf0(j, 0) = "Null" Then
                    If _BidList(i) = buf0(j, 0) Then
                        If Not buf0(j, 1) = "Null" Then _DataBuf(i, 2) = buf0(j, 1) ' เก็บ EQP_NO
                        If Not buf0(j, 2) = "Null" Then _DataBuf(i, 4) = buf0(j, 2) ' เก็บ RECIPE
                        If Not buf0(j, 3) = "Null" Then _DataBuf(i, 18) = buf0(j, 3) ' เก็บ START_TIME
                        Exit For
                    End If
                End If
            Next
            ' จับคู่ Lot Info
            For j As Integer = 0 To UBound(buf1, 1)
                eCode = "d Get_Data_Buf " & i & " " & j & " " & _BidList.Count - 1 & " " & UBound(buf1, 1)
                If Not buf1(j, 0) = "Null" Then
                    If _BidList(i) = buf1(j, 0) Then
                        If Not buf1(j, 1) = "Null" Then _DataBuf(i, 1) = buf1(j, 1) ' เก็บ Process
                        If Not buf1(j, 2) = "Null" Then _DataBuf(i, 14) = buf1(j, 2) ' เก็บ LotNo
                        If Not buf1(j, 3) = "Null" Then _DataBuf(i, 17) = buf1(j, 3) ' เก็บ InCharge
                        If Not buf1(j, 4) = "Null" Then _DataBuf(i, 5) = buf1(j, 4) ' เก็บ OpeName
                        Exit For
                    End If
                End If
            Next
        Next
        
        eDt = DateTime.Now
        ts(2) = eDt - sDt
        BackgroundWorker1.ReportProgress(7)
        BackgroundWorker1.ReportProgress(_H)
        sDt = DateTime.Now
        
        ' buf2: พักข้อมูลจาก TableD (Data)
        Dim buf2(_TableD.Rows.Count - 1, 3 - 1) As String
        For i As Integer = 0 To UBound(buf2, 1)
            eCode = "e Get Data_Buf " & i & " " & UBound(buf2, 1)
            ' ... (กำหนดค่า Null) ...
            If Not IsDBNull(_TableD.Rows(i)("DATA_ID")) Then buf2(i, 0) = _TableD.Rows(i)("DATA_ID")
            If Not IsDBNull(_TableD.Rows(i)("MEASURE_ITEM")) Then buf2(i, 1) = _TableD.Rows(i)("MEASURE_ITEM")
            If Not IsDBNull(_TableD.Rows(i)("DATA")) Then buf2(i, 2) = _TableD.Rows(i)("DATA")
        Next
        _TableD.Clear()
        
        ' --- วนลูปเอาข้อมูลวัดผล (Data) มาเรียงใส่ _DataBuf ตั้งแต่คอลัมน์ 26 เป็นต้นไป ---
        For i As Integer = 0 To _DidList.Count - 1
            Count_Data(i) = 0
            For j As Integer = 0 To UBound(buf2, 1)
                eCode = "f Get Data_Buf " & i & " " & j & " " & _DidList.Count - 1 & " " & UBound(buf2, 1)
                If Not buf2(j, 0) = "Null" Then
                    If _DidList(i) = buf2(j, 0) Then
                        If Count_Data(i) = 0 Then
                            If Not buf2(j, 1) = "Null" Then _DataBuf(i, 3) = buf2(j, 1) ' เก็บ Measure Item (เก็บครั้งเดียว)
                        End If
                        If Not buf2(j, 2) = "Null" Then
                            ' เก็บข้อมูลดิบลงคอลัมน์ 26, 27, 28...
                            _DataBuf(i, 26 + Count_Data(i)) = buf2(j, 2)
                            Count_Data(i) += 1
                        End If
                    End If
                End If
            Next
        Next

        ' ตัวแปรสำหรับคำนวณสถิติ
        Dim max As Double = -100000 ' ค่าสูงสุด (ตั้งเริ่มต้นต่ำมากๆ)
        Dim min As Double = 100000  ' ค่าต่ำสุด (ตั้งเริ่มต้นสูงมากๆ)
        Dim sum As Double = 0       ' ผลรวม
        Dim c As Integer = 0        ' ตัวนับจำนวน

        ' --- วนลูปคำนวณค่าสถิติ (Max, Min, Avg) ---
        For i As Integer = 0 To UBound(_DataBuf, 1)
            _DataBuf(i, 0) = _MaxID + i ' กำหนด iID (Running Number)

            If Not _DataBuf(i, 26) = "Null" Then
                ' ลูปอ่านข้อมูลดิบ 100 ตัว (สมมติว่าไม่เกิน 100) จากคอลัมน์ 26
                For j As Integer = 0 To 100 - 1
                    eCode = "g Get_Data_Buf " & i & " " & j & " " & UBound(_DataBuf, 1) & " " & 100 - 1
                    If _DataBuf(i, 26 + j) = "Null" Then Exit For ' ถ้าไม่มีข้อมูลแล้วก็ออกลูป

                    If j = 0 Then
                        max = _DataBuf(i, 26 + j)
                        min = _DataBuf(i, 26 + j)
                    Else
                        If max < _DataBuf(i, 26 + j) Then max = _DataBuf(i, 26 + j) ' หา Max
                        If min > _DataBuf(i, 26 + j) Then min = _DataBuf(i, 26 + j) ' หา Min
                    End If
                    sum += _DataBuf(i, 26 + j) ' บวกสะสม
                    c += 1
                Next
                eCode = "h Get_Data_Buf "
                ' คำนวณเสร็จ เก็บค่าลง Array ในตำแหน่งที่กำหนด
                _DataBuf(i, 22) = Mid(sum / c, 1, 10) ' ค่าเฉลี่ย (AVG)
                _DataBuf(i, 23) = Mid(max - min, 1, 10) ' ค่าพิสัย (Range)
                _DataBuf(i, 24) = max ' ค่าสูงสุด (MAX)
                _DataBuf(i, 25) = min ' ค่าต่ำสุด (MIN)
            Else
                ' กรณีไม่มีข้อมูล ให้ใส่ 0
                _DataBuf(i, 22) = 0
                _DataBuf(i, 23) = 0
                _DataBuf(i, 24) = 0
                _DataBuf(i, 25) = 0
            End If
            ' รีเซ็ตค่าเพื่อคำนวณแถวถัดไป
            sum = 0
            c = 0
            max = -100000
            min = 100000
        Next
        eDt = DateTime.Now
        ts(3) = eDt - sDt
    End Sub

    ' ฟังก์ชัน Insert ข้อมูลจาก _DataBuf ลง SPC_Master
    Public Function Write_Data(ByVal _Data(,) As String) As String
        Write_Data = ""
        Dim Cn As New SqlConnection
        Dim strSQL As String = ""
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction = Nothing
        Try
            Cn.ConnectionString = StrServerConnection_Yuku
            For i As Integer = 0 To UBound(_Data, 1)
                Cn.Open()
                trans = Cn.BeginTransaction
                SQLCm.Transaction = trans
                strSQL = ""
                strSQL = "INSERT INTO " & "SPC_Master" & " VALUES ("

                For j As Integer = 0 To UBound(_Data, 2)
                    eCode = "a Write_Data " & i & " " & j & " " & UBound(_Data, 1) & " " & UBound(_Data, 2)
                    If _Data(i, j) = "Null" Then
                        strSQL &= "Null"
                    Else
                        strSQL &= "'" & _Data(i, j) & "'"
                    End If
                    If j = UBound(_Data, 2) Then
                        strSQL &= ")"
                    Else
                        strSQL &= ","
                    End If
                Next
                eCode = "b Write_Data " & strSQL
                SQLCm.CommandText = strSQL
                SQLCm.ExecuteNonQuery()
                trans.Commit()
                Cn.Close()
            Next
        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            OKflag = False
            Cn.Close()
            Write_Error(">>" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
            Write_Data = "error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace
        End Try
    End Function

    ' บันทึก Log ไฟล์ Debug (ถ้ามี)
    Public Sub Write_History()
        Dim TextFile As IO.StreamWriter = New IO.StreamWriter(StrCDir & "\" & "DebugFile.txt", True, System.Text.Encoding.Default)
        Dim Str As String = ""
        Str = ""
        TextFile.Write(Str)
        TextFile.WriteLine()
        TextFile.Close()
    End Sub

    ' [ปุ่ม Create Table] สร้างตารางฐานข้อมูลใหม่จาก CSV
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If StrServerConnection_RIST = "" Then
            MsgBox("Set the RIST server connection string.")
            Exit Sub
        End If
        If StrServerConnection_Yuku = "" Then
            MsgBox("Set the Yukuhashi server connection string.")
            Exit Sub
        End If
        OKflag = True
        Form3.Show()
        Form3.Label1.Text = "Table creating..."
        Try
            Input_CSV_to_Server(StrCDir & " \SPC_Alarm.csv", 0)
            Input_CSV_to_Server(StrCDir & " \SPC_Master.csv", 0)
            Input_CSV_to_Server(StrCDir & " \SPC_Property.csv", 0)
            Input_CSV_to_Server(StrCDir & " \SPC_User.csv", 0)
        Catch ex As Exception
            OKflag = False
        End Try
        Form3.Close()
        If OKflag = True Then
            MsgBox("Table creation OK")
        Else
            MsgBox("Table creation NG")
        End If
    End Sub

    Public Function Input_CSV_to_Server(ByVal CsvName As String, ByVal Del As Integer) As String
        Input_CSV_to_Server = ""
        Dim CSVData(0, 0) As String
        Dim CSVData_Header(0, 0) As String
        Dim TableName As String = ""
        Input_CSV_to_Server = Input_CSV(CsvName, CSVData, CSVData_Header, TableName)
        If Not Input_CSV_to_Server = "" Then
            Exit Function
        End If
        Input_CSV_to_Server = To_Server(CSVData, CSVData_Header, TableName, Del)
        If Not Input_CSV_to_Server = "" Then
            Exit Function
        End If
    End Function

    Public Function Input_CSV(ByVal Filename As String, ByRef Array(,) As String, ByRef Array_Header(,) As String, ByRef TName As String) As String
        Input_CSV = ""
        Dim eCode As String = ""
        Try
            Dim Buf(,) As String
            Dim temp0() As String
            Dim sr0 As System.IO.StreamReader
            Dim CoMax As Integer = 0
            Dim Gyo As Integer = 0
            eCode = "a"
            sr0 = New System.IO.StreamReader(Filename, System.Text.Encoding.Default)
            Do Until sr0.Peek() = -1
                temp0 = Split(sr0.ReadLine(), ",")
                If CoMax < temp0.Length Then
                    CoMax = temp0.Length
                End If
                Gyo += 1
            Loop
            sr0.Close()
            eCode = "b"
            ReDim Buf(Gyo - 1, CoMax - 1)
            Gyo = 0
            sr0 = New System.IO.StreamReader(Filename, System.Text.Encoding.Default)
            Do Until sr0.Peek() = -1
                temp0 = Split(sr0.ReadLine(), ",")
                For i As Integer = 0 To CoMax - 1
                    Buf(Gyo, i) = temp0(i)
                Next
                Gyo += 1
            Loop
            sr0.Close()
            eCode = "c"
            ReDim Array_Header(3 - 1, UBound(Buf, 2) - 1)
            ReDim Array(UBound(Buf, 1) - (UBound(Array_Header, 1) + 1) - 1, UBound(Buf, 2) - 1)
            TName = Buf(0, 1)
            For i As Integer = 1 To 1 + UBound(Array_Header, 1)
                For j As Integer = 1 To UBound(Buf, 2)
                    Array_Header(i - 1, j - 1) = Buf(i, j)
                Next
            Next
            eCode = "d"
            For i As Integer = 1 + UBound(Array_Header, 1) + 1 To UBound(Buf, 1)
                For j As Integer = 1 To UBound(Buf, 2)
                    Array(i - (1 + UBound(Array_Header, 1) + 1), j - 1) = Buf(i, j)
                Next
            Next
            eCode = "e"
        Catch ex As Exception
            Input_CSV = "Input_CSV " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace
        End Try
    End Function

    Public Function To_Server(ByVal Array(,) As String, ByVal Array_Header(,) As String, ByVal TName As String, ByVal Del As Integer) As String
        To_Server = ""
        Dim eCode As String = ""
        Dim Cn As New SqlConnection
        Dim strSQL As String = ""
        Dim SQLCm As SqlCommand = Cn.CreateCommand
        Dim trans As SqlTransaction = Nothing
        Try
            eCode = "a"
            Cn.ConnectionString = StrServerConnection_Yuku
            Cn.Open()
            trans = Cn.BeginTransaction
            SQLCm.Transaction = trans
            If 1 Then
                eCode = "b"
                If Check_TableRist(TName) = True Then
                    If Del = 0 Then
                        MsgBox("Table:" & TName & " already exists.")
                        OKflag = False
                        Cn.Close()
                        Exit Function
                    End If
                    eCode = "c"
                    SQLCm.CommandText = "DROP TABLE " & TName
                    SQLCm.ExecuteNonQuery()
                End If
                eCode = "d"
                strSQL = "Create Table " & TName & "("
                For j As Integer = 0 To UBound(Array_Header, 2)
                    strSQL &= Array_Header(0, j) & " " & Array_Header(1, j)
                    If Array_Header(2, j) = "NO" Then
                        strSQL &= " Not Null"
                    End If
                    If j = UBound(Array_Header, 2) Then
                        strSQL &= ",PRIMARY KEY (" & Array_Header(0, 0) & ")"
                        strSQL &= ")"
                    Else
                        strSQL &= ","
                    End If
                Next
                eCode = "e"
                SQLCm.CommandText = strSQL
                SQLCm.ExecuteNonQuery()
                eCode = "f"
                SQLCm.CommandText = "DELETE FROM " & TName
                SQLCm.ExecuteNonQuery()
                eCode = "g"
                For i As Integer = 0 To UBound(Array, 1)
                    strSQL = ""
                    strSQL = "INSERT INTO " & TName & " VALUES ("
                    For j As Integer = 0 To UBound(Array, 2)
                        If Array(i, j) = "Null" Then
                            strSQL &= "Null"
                        Else
                            strSQL &= "'" & Array(i, j) & "'"
                        End If
                        If j = UBound(Array, 2) Then
                            strSQL &= ")"
                        Else
                            strSQL &= ","
                        End If
                        eCode = "h" & i & " " & j
                    Next
                    SQLCm.CommandText = strSQL
                    SQLCm.ExecuteNonQuery()
                Next
            End If
            trans.Commit()
            Cn.Close()
        Catch ex As Exception
            If IsNothing(trans) = False Then
                trans.Rollback()
            End If
            OKflag = False
            To_Server = "To_Server " & eCode & Environment.NewLine & strSQL & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace
        End Try
    End Function

    Dim TableRist As String()
    Public Function Make_TableRist() As Integer
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim strSQL As String = ""
        Dim dtColumns As DataTable
        Dim Buf() As String
        Try
            Make_TableRist = 0
            Cn.ConnectionString = StrServerConnection_Yuku
            Cn.Open()
            dtColumns = Cn.GetSchema("Columns")
            Cn.Close()
            Cn.Dispose()
            Make_TableRist = dtColumns.Rows.Count
            If Make_TableRist = 0 Then
                Exit Function
            End If
            ReDim Buf(dtColumns.Rows.Count - 1)
            For i As Integer = 0 To dtColumns.Rows.Count - 1
                Buf(i) = dtColumns.Rows(i)("TABLE_NAME")
            Next
            Dim al2 As New System.Collections.ArrayList(UBound(Buf, 1))
            Dim j2 As String
            For Each j2 In Buf
                If Not al2.Contains(j2) Then
                    al2.Add(j2)
                End If
            Next
            TableRist = DirectCast(al2.ToArray(GetType(String)), String())
        Catch ex As System.Exception
            Cn.Dispose()
            MsgBox("Error getting table list" + "," + ex.Message & ex.StackTrace)
            Exit Function
        End Try
    End Function

    Public Function Check_TableRist(ByVal TName As String) As Boolean
        Check_TableRist = False
        If Make_TableRist() = 0 Then
            Exit Function
        End If
        For i As Integer = 0 To UBound(TableRist, 1)
            If TableRist(i) = TName Then
                Check_TableRist = True
                Exit Function
            End If
        Next
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DialogResult
        Dim frm As New Form2
        dr = frm.ShowDialog
    End Sub

    Dim MaxID As Integer = 0 ' MaxID: เก็บค่า ID สูงสุดปัจจุบัน เพื่อ Run Number ต่อ
    Dim gRows As Integer = 1000 ' gRows: จำนวนแถวสูงสุดที่จะดึงต่อรอบ (Limit ไว้ไม่ให้หนักเกิน)

    ' บันทึก Error ลงไฟล์ Text และส่ง Line Notify
    Private Sub Write_Error(ByVal strE As String)
        Dim TextFile As IO.StreamWriter = New IO.StreamWriter(StrCDir & "\Emsg.txt", True, System.Text.Encoding.Default)
        TextFile.Write(strE)
        TextFile.WriteLine()
        TextFile.Close()

        LineNotify("แจ้งเตือนสถานะ SPC Converter ERROR" + vbNewLine + "ประจำวันที่: " + Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    Dim recDay As String = "" ' recDay: วันที่ของข้อมูลล่าสุดที่ดึงไปแล้ว
    Dim F_BATCHID As String = "" ' F_BATCHID: Batch ID แรกในรอบการทำงานนี้
    Dim E_BATCHID As String = "" ' E_BATCHID: Batch ID สุดท้ายในรอบการทำงานนี้
    Dim Bidlist As New List(Of String) ' List เก็บ Batch ID

    ' ==========================================
    ' MAIN LOOP: การทำงานเบื้องหลัง (Background Worker)
    ' ==========================================
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim TableB As New DataTable ' TableB: ตารางพักข้อมูล Batch
        Dim TableL As New DataTable ' TableL: ตารางพักข้อมูล Lot
        Dim TableD As New DataTable ' TableD: ตารางพักข้อมูล Data
        Dim DataBuf(0, 126 - 1) As String ' DataBuf: อาร์เรย์หลักที่รวมข้อมูลทุกอย่างไว้เตรียมบันทึก
        Dim buf As String = ""
        Dim DidList As New List(Of String) ' DidList: List เก็บ Data ID
        
        BackgroundWorker1.ReportProgress(Display_Clear)
        BackgroundWorker1.ReportProgress(_A)
        Try
            ' หาข้อมูลวันที่ล่าสุดที่มี
            recDay = Get_Recently_Day()
        Catch ex As Exception
            Write_Error(">>" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
            BackgroundWorker1.ReportProgress(9)
            Exit Sub
        End Try
        
        ' ถ้า user กด Cancel
        If BackgroundWorker1.CancellationPending Then
            Exit Sub
        End If
        
        BackgroundWorker1.ReportProgress(_B)
        Try
            ' หา ID ล่าสุด
            buf = Get_NextMasterNo2("SPC_Master", StrServerConnection_Yuku)
        Catch ex As Exception
            Write_Error(">" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
            BackgroundWorker1.ReportProgress(9)
            Exit Sub
        End Try
        If Not buf = -1 Then
            MaxID = CInt(buf) + 1 ' กำหนด ID เริ่มต้นสำหรับรอบนี้
        Else
            Write_Error(">>" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " buf=-1 ")
            BackgroundWorker1.ReportProgress(9)
            Exit Sub
        End If
        If BackgroundWorker1.CancellationPending Then
            Exit Sub
        End If
        eCode = "d"
        BackgroundWorker1.ReportProgress(Set_Rows)
        BackgroundWorker1.ReportProgress(_D)
        
        ' Do Loop: ทำงานวนไปเรื่อยๆ จนกว่าจะกด Stop หรือ Error
        Do
            Write_times += 1
            If Write_times > 100 Then
                Write_times = 0
                BackgroundWorker1.ReportProgress(_K)
            End If
            If BackgroundWorker1.CancellationPending Then
                Exit Do
            End If
            SDt_All = DateTime.Now
            BackgroundWorker1.ReportProgress(4)
            BackgroundWorker1.ReportProgress(_E)
            sDt = DateTime.Now
            eCode = "e"
            Try
                ' 1. ดึงข้อมูล Batch ใหม่ๆ (หลังจาก recDay)
                Get_Data_BATCH_TBL(TableB, gRows, recDay)
            Catch ex As Exception
                Write_Error(">" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
                BackgroundWorker1.ReportProgress(9)
                Exit Sub
            End Try
            eDt = DateTime.Now
            ts(0) = eDt - sDt
            
            ' ถ้าไม่มีข้อมูลใหม่ ให้รอ 5 นาที
            If TableB.Rows.Count = 0 Then
                BackgroundWorker1.ReportProgress(_C)
                For i As Integer = 0 To 60 * 5 - 1
                    System.Threading.Thread.Sleep(1000) ' หลับ 1 วินาที (วน 300 รอบ = 5 นาที)
                    If BackgroundWorker1.CancellationPending Then
                        Exit Do
                    End If
                Next
                Continue Do ' กลับไปเริ่มลูปใหม่
            End If
            
            If BackgroundWorker1.CancellationPending Then
                Exit Do
            End If
            BackgroundWorker1.ReportProgress(5)
            BackgroundWorker1.ReportProgress(_F)
            sDt = DateTime.Now
            eCode = "f"
            Try
                ' 2. ดึงข้อมูล Lot และ Data
                Get_Data_LOT_DATA_TBL(TableB, TableL, TableD, Bidlist, DidList, F_BATCHID, E_BATCHID)
            Catch ex As Exception
                Write_Error(">>" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
                BackgroundWorker1.ReportProgress(9)
                Exit Sub
            End Try
            eDt = DateTime.Now
            ts(1) = eDt - sDt
            If BackgroundWorker1.CancellationPending Then
                Exit Do
            End If
            BackgroundWorker1.ReportProgress(_G)
            Try
                ' 3. คำนวณและจัดเตรียมข้อมูล (Calculate Stats: Avg, Min, Max)
                Get_Data_Buf(TableB, TableL, TableD, DataBuf, MaxID, 0, Bidlist, DidList)
            Catch ex As Exception
                Write_Error(">" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
                BackgroundWorker1.ReportProgress(9)
                Exit Sub
            End Try
            If BackgroundWorker1.CancellationPending Then
                Exit Do
            End If
            SDt = DateTime.Now
            Dflag = True
            BackgroundWorker1.ReportProgress(8)
            BackgroundWorker1.ReportProgress(_I)
            sDt = DateTime.Now
            
            ' 4. บันทึกลงฐานข้อมูลปลายทาง
            If Not Write_Data(DataBuf) = "" Then
                BackgroundWorker1.ReportProgress(9)
                Exit Sub
            End If
            eCode = "s"
            Dflag = False
            eDt = DateTime.Now
            ts(4) = eDt - sDt
            BackgroundWorker1.ReportProgress(_J)
            eDt_All = DateTime.Now
            ts_All = eDt_All - SDt_All
            
            ' อัปเดต MaxID สำหรับรอบถัดไป
            MaxID += Bidlist.Count
            BackgroundWorker1.ReportProgress(Display_Seconds)
            
            ' อัปเดตวันที่ล่าสุด (recDay) จากข้อมูลตัวสุดท้ายที่เพิ่งบันทึกไป
            recDay = DataBuf(UBound(DataBuf, 1), 18)
            If BackgroundWorker1.CancellationPending Then
                Exit Do
            End If
        Loop
    End Sub

    ' --- ตัวแปร Constant สำหรับ State การทำงาน (ใช้สื่อสารกับ Progress Bar) ---
    Public Display_Clear As Integer = 0
    Public Set_Rows As Integer = 1
    Public Display_Seconds As Integer = 2
    Public Stop_Convert As Integer = 3
    Public _A As Integer = 10
    Public _B As Integer = 11
    Public _C As Integer = 12
    Public _D As Integer = 13
    Public _E As Integer = 14
    Public _F As Integer = 15
    Public _G As Integer = 16
    Public _H As Integer = 17
    Public _I As Integer = 18
    Public _J As Integer = 19
    Public _K As Integer = 20

    ' แสดงความคืบหน้าบน Form3
    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ' e.ProgressPercentage: รับค่า State ที่ส่งมาจาก BackgroundWorker แล้วแสดงข้อความตามนั้น
        If e.ProgressPercentage = Display_Clear Then
            ForeColorBlack()
            Form3.Label1.Text = "Preparing..."
            Form3.TextBox1.Text = ""
        ElseIf e.ProgressPercentage = Set_Rows Then
            Form3.Label1.Text = "Processing..."
        ElseIf e.ProgressPercentage = Display_Seconds Then
        ElseIf e.ProgressPercentage = Stop_Convert Then
            Form3.Label1.Text = "Stopping..."
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Stopping..." & Environment.NewLine
            TextScroll()
        ElseIf e.ProgressPercentage = 9 Then
            Form4.TopMost = True
            Form4.Show()
        ElseIf e.ProgressPercentage = _A Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Getting latest date of data from the SPC_Master." & Environment.NewLine
        ElseIf e.ProgressPercentage = _B Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Getting number of data in the SPC_Master." & Environment.NewLine
        ElseIf e.ProgressPercentage = _C Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > No latest data." & Environment.NewLine
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Wait for 5 minutes." & Environment.NewLine
        ElseIf e.ProgressPercentage = _D Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Start conversion." & Environment.NewLine
        ElseIf e.ProgressPercentage = _E Then
            Form3.TextBox1.Text &= "--------------------------------------------------------" & Environment.NewLine
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Getting data in the BATCH_TBL. ( After " & recDay & ")" & Environment.NewLine
            Form3.TextBox1.SelectionStart = Form3.TextBox1.Text.Length
            TextScroll()
        ElseIf e.ProgressPercentage = _F Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Getting data in the DATA_TBL and LOT_TBL." & Environment.NewLine
            TextScroll()
        ElseIf e.ProgressPercentage = _G Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Organizing data in the BATCH_TBL and LOT_TBL." & Environment.NewLine
            TextScroll()
        ElseIf e.ProgressPercentage = _H Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Organizing data in the DATA_TBL." & Environment.NewLine
            TextScroll()
        ElseIf e.ProgressPercentage = _I Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Writing data." & Environment.NewLine
            TextScroll()
        ElseIf e.ProgressPercentage = _J Then
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > Conversion successful. ( BATCH_ID [" & F_BATCHID & "] to [" & E_BATCHID & "], " & Bidlist.Count & "rows )" & Environment.NewLine
            TextScroll()
            Form3.TextBox1.Text &= Format(DateTime.Now, "MM/dd HH:mm:ss") & " > (Elapsed time: " & Math.Round(ts(0).TotalSeconds, 2, MidpointRounding.AwayFromZero) & "s, " & Math.Round(ts(1).TotalSeconds, 2, MidpointRounding.AwayFromZero) & "s, " & Math.Round(ts(2).TotalSeconds, 2, MidpointRounding.AwayFromZero) & "s, " & Math.Round(ts(3).TotalSeconds, 2, MidpointRounding.AwayFromZero) & "s, " & Math.Round(ts(4).TotalSeconds, 2, MidpointRounding.AwayFromZero) & "s / Total: " & Math.Round(ts_All.TotalSeconds, 2, MidpointRounding.AwayFromZero) & "s " & Math.Round(Bidlist.Count / ts_All.TotalMinutes, 0, MidpointRounding.AwayFromZero) & "row/min" & " )" & Environment.NewLine
            TextScroll()
        ElseIf e.ProgressPercentage = _K Then
            Form3.TextBox1.Text = ""
        End If
    End Sub

    Public Sub ForeColorBlack()
        Form3.Label5.ForeColor = Color.Black
        Form3.Label6.ForeColor = Color.Black
        Form3.Label7.ForeColor = Color.Black
        Form3.Label8.ForeColor = Color.Black
        Form3.Label9.ForeColor = Color.Black
    End Sub

    Public Sub TextScroll()
        Form3.TextBox1.SelectionStart = Form3.TextBox1.Text.Length
        Form3.TextBox1.Focus()
        Form3.TextBox1.ScrollToCaret()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Form3.Close()
    End Sub

    ' [ปุ่ม Delete Data] ลบข้อมูลทั้งหมดใน SPC_Master
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim result As DialogResult
        result = MessageBox.Show("Do you want to delete the data in the SPC_Master ?", "Delete Data", MessageBoxButtons.OKCancel)
        If result = Windows.Forms.DialogResult.OK Then
            eMsg = Input_CSV_to_Server(StrCDir & "\SPC_Master.csv", 1)
            If Not eMsg = "" Then
                Write_Error("error SPC_Master_Recreated" & eMsg)
                Form4.Show()
                Exit Sub
            End If
            MsgBox("Deleted the data in the SPC_Master.")
        End If
    End Sub

    ' ฟังก์ชันส่งแจ้งเตือนเข้า Line Notify
    Private Sub LineNotify(ByVal msg As String)
        Dim token As String = My.Settings.LineToken ' LineToken: รหัส Token ที่ได้จาก Line Developer
        Try
            Dim request As HttpWebRequest = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notify"), HttpWebRequest)
            Dim postData = String.Format("message={0}", msg)
            Dim data = Encoding.UTF8.GetBytes(postData)
            request.Method = "POST"
            request.ContentType = "application/x-www-form-urlencoded"
            request.ContentLength = data.Length
            request.Headers.Add("Authorization", "Bearer " + token)
            request.AllowWriteStreamBuffering = True
            request.KeepAlive = False
            request.Credentials = CredentialCache.DefaultCredentials
            Using stream = request.GetRequestStream()
                stream.Write(data, 0, data.Length)
            End Using
            Dim response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Dim responseString = New StreamReader(response.GetResponseStream()).ReadToEnd()
        Catch ex As Exception
            Dim LogFilePath As String = My.Application.Info.DirectoryPath + "/" + My.Settings.LineLog
            Using FileLog As New FileStream(LogFilePath, FileMode.Append, FileAccess.Write)
                Using WriteFile As New StreamWriter(FileLog)
                    WriteFile.WriteLine(DateTime.Now)
                    WriteFile.WriteLine(ex.Message)
                    WriteFile.WriteLine(" {0} Exception caught.", ex)
                    WriteFile.WriteLine("------------------------------------------------------")
                End Using
            End Using
        End Try
    End Sub

    ' [ปุ่ม Test Line] ทดสอบส่ง Line
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        LineNotify("แจ้งเตือนสถานะ SPC Converter ERROR" + vbNewLine + "ประจำวันที่: " + Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub
End Class