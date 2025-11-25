Imports System.Data.SqlClient
Imports System.Net
Imports System.Text
Imports System.IO

Public Class Form1
	Public StrServerConnection_RIST As String = “”
	Public StrServerConnection_Yuku As String = “”
	Public StrStartDate As String = “”
	public StrDir As String = “”
	Public OKFlag As Boolean
	Private Sub Form1_Load(ByVal sender As System.Object, ByVale As System.EventArds) Handles MyBase.Load
		Button2.Enablzed = False
		Button4.Enabled = False
		ReadConfigData()
		SaveConfigData()
		If My.Settings.DisplayNotify Then Button5.Visible = True
	End Sub
	Private Sub ReadConfigData()
		Dim strFileName As String = strCDir & "\Config. csv"
		Dim temp() As String
		Dim sr As New System.IO.StreamReader (strFileName, System.Text.Encoding.Default)
		Do Until sr. Peek() = -1
	
			temp = Split(sr. ReadLine(), “,")
			If temp(0).Trim(Chr(34)) <> “” Then
				If temp(0) = "Server_RIST" Then
					StrServerConnection_RIST = temp (1)
				ElseIf temp(e) = "Server_Yukuhashi" Then
					StrServerConnection_Yuku = temp(1)
				ElseIf temp(0) = "StartDate" Then
					StrStarDate = temp(1)
				End If
			End If
		Loop 
		sr. Close()

		If StrStartDate = "" Then
			StrStartDate = "2020-01-01 00:00:00.000"
			Form2.Button2.Enabled = False
			Button3_Click(Nothing, Nothing)
		End If
	End Sub
	Public Sub SaveConfigData()
		Dim str1 As String
		Dim strFileName As String = strCDir & "\Config-csv"
		Dim sw1 As New System. IO.StreamWriter(strFileName, False, System. Text.Encoding.Default)
		str1 = Err.Description
		If str1 - “” Then 
			sw1.WriteLine(",Connection string")
			sw1.WriteLine("Server_RIST," & StrServerConnection_RIST)
			sw1.WriteLine("Server_Yukuhashi," & StrServerConnection_Yuku) 
			sw1.WriteLine("StartDate," & StrStartDate)
			sw1.Close()
		End If
	End Sub

	Public Function Get_NextMasterNo() As String
		Dim Cn As New System.Data.SqlClient.SqlConnection
		Dim Adapter As New SqlDataAdapter
		Dim strSQL As String = "”
		Dim n As Integer
		Dim table As New DataTable
		Try
			Get_NextMasterNo = "”
			Cn.ConnectionString = StrServerConnection_Yuku
			table.Clear()
			strSQL = "SELECT İID"
			strQL &= "FROM SPC_Master"
I
			Adapter = New SqlDataAdapter()
			Adapter.SelectCommand = New SqlCommand (strSQL, Cn)
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

			Get_NextMasterNo = table. Rows(n - 1) (0) + 1
		
		Catch ex As System.Exception
			Get_NextMasterNo = “”
			Adapter.Dispose()
			Cn.Dispose()
			MsgBox("ErrorCode " + eCode + " , " + ex.Message & ex.StackTrace)
			Exit Function
		End Try
	End Function

	Public Function Get_NextMasterNo2(ByVal Tname As String, ByVal StrConnectionString As String) As String
		
		eCode = Tname & “Get_NextMasterNo2”
 		Get_NextMasterNo2 = -1
		Dim Cn As New SqlConnection
		Dim SQLCm As SqlCommand = Cn.CreateCommand
		Dim strSQL As String = “”
		Cn.ConnextionString = StrConnectionString 
		SQLCm.CommandTimeout = 600
		strSQL = “SELECT MAX(iID) FROM ” & Tname
		Cn.Open()
		SQLCm.CommandText = strSQL
		Get_NestMasteterNo2 = SQLCm.ExecuteScalar
		Cn.Close()
		SQLCm.Dispose()
		Cn.Dispose()
	End Function

	Public Sub To_Servere(ByVal Array0(,) As String, ByVal Array1(,) As String)
		Dim Cn As New SqlConnection
		Dim strSQL As String 
		Dim SQLCm As SqlCommand = Cn.CreateCommand
		Dim trans As SqlTransaction = Nothing
		
		Try
			Cn.ConnectionString = StrServerConnection_Yuku
			Cn.Open()
			trans = Cn.BeginTransaction
			SQLCm.Transaction = trans

			For i As Integer = 0 To UBound (Array0, 1)
				strSQL = “”
				strSQL - "INSERT INTO " & "SPC_Master" & " VALUES ("
				
				For j As Integer = 0 To UBound (Arraye, 2)
					If Array0(i, j) = "Null" Then
						strSQL &= "Null"
					Else
						strSQL &= "’" & Array0(i, j) & "’"
					End If 
					strSQL &= ","
				Next
				For j As Integer = 0 To UBound(Array1, 2)
					If Array1(i, j) ‎ =  "Null" Then 
						strSQL &= "Null"
					Else
						strSQL &= "'" & Array1（i, j） & "’”
					End if
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
				trans.Rollback()
			End If
			OKflag = False
			MSgBox("ErrorCode " + eCode + "," + ex.Message & ex.StackTrace)
			End Sub
		End Try

	End Sub

	Dim eCode As String
	Public Dflag As Boolean = False
	Dim SDt_AlI As DateTime
	Dim eDt_All As DateTime
	Dim SDt As DateTime
	Dim eDt As DateTime
	Dim ts(6 - 1) As TimeSpan
	Dim ts_All As TimeSpan
	Public StopFlag As Boolean
	Dim eMsg As String="”
	Dim Write_times As Integer = 0
	Private Sub Butto1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

		Form4. Close()
		If StrServerConnection_RIST = “" Then
			MsgBox("Set the RIST server connection string.")
			Exit Sub
		End If
		If StrServerConnection_Yuku = "" Then
			MsgBox("Set the Yukuhashi server connection string.")
			Exit Sub
		End If

		Write_times = 0
		StopFlag - False
		OKflag = True
		Me.Hide ()
		Form3.Show()
		Form3.TopMost = True

		Backgroundworker1.WorkerReportsProgress = True
		Backgroundworker1.WorkerSupportsCancellation = True
		Backgroundworker1.RunworkerAsync()
	End Sub

	Public Function Get_Recently_Day() As String
		eCode = "a Get_Recently_Day"
		Dim Cn As New SqlConnection
		Dim SQLCm As SqlCommand = Cn.CreateCommand
		Dim Adapter As New SqlDataAdapter(SQLCm)
		Dim Table As New DataTable
		Dim strSQL As String = ""
		Cn. ConnectionString = StrServerConnection_Yuku
		SQLCm.CommandTimeout = 600
		strsQL = "SELECT MAX (dworkDate) FROM SPC_Master"
		strSQL &= " WHERE dworkDate < ‘2292-01-01 00:00:00.000’"
		SQLCm.CommandText = strSQL
		Adapter.Fill(Table)
		Table.Dispose()
		Adapter.Dispose()
		SQLCm.Dispose()
		Cn.Dispose()
		eCode = "b Get_Recently_Day"
		If Not IsDBNull(Table.Rows(0)(0)) Then
			Get_Recently_Day = Format(Table.Rows(0)(0), "yyyy-MM-dd HH: mm:ss.fff")
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
	Public Sub Get_Data_BATCH_TBL(ByRef _TableB As DataTable, ByVal _gRows As Integer,
ByVal_recDay As String)
		eCode = a Get_Date_BATCH_TBL"
		_TableB.Clear()
		Dim Cn As New SqlConnection
		Dim SQLCm As SqlCommand = Cn.CreateCommand
		Dim Adapter As New SqlDataAdapter(SQLCm)
		Dim Table As New DataTable
		Dim strSQL As String = ""

		Cn. ConnectionString = StrServerConnection_RIST
		SQLCm.CommandTimeout = 600
		strSQL = "SELECT TOP " & _gRows
		strSQL &= " BATCH_ID,”
		strSQL &= "DATA_ID, " 
		strSQL &= "START_TIME, "
		strSQL &= "RECIPE, "
		strSQL &= "EQP_NO"
		strSQL &= "FROM BATCH_TBL"
		If Not _recDay = Nothing Then
			StrSQL &= “WHERE START_TIME > ”&"’” &_recDay &"’"
			strSQL &= " AND START_TIME < ‘2292-01-01 00:00:00.000’"
		Else
			strSQL &= " WHERE START_TIME < '2292-81-01 00:08:00.000'"
		End IF
		strSQL & = " ORDER BY START_TIME ASC"
		eCode = "b Get_Data_BATCH_TBL" & Environment.NewLine & strSQL
		SQLCm.CommandText = strSQL
		Adapter. Fill(_TableB)
		_TableB.Dispose()
		Adapter.Dispose()
		SQLCm.Dispose()
		Cn.Dispose()
		eCode = "c Get_Data_BATCH_TBL" & strSQL
	End Sub
	Public Sub Get_Data_LOT_DATA_TBL(ByVal _TableB As Datatable, ByRef _TablelL As Datatable, ByRef_TableD As DataTable, ByRef _Bidlist As List(Of String), ByRef_DidList As List(Of String),ByRef _F_BATCHID As String, ByRef_E_BATCHID As String)
		eCode = "a Get_Data_LOT_DATA_TBL"
		_BidList.Clear()
		_DidList.Clear()
		sDt = DateTime.Now
		For i As Integer = 0 To_TableB.Rows.Count - 1
			If Not IsDBNuIl(_TableB.Rows(i) ("BATCH_ID"')) Then
				_BidList.Add(_TableB.Rows(i) ("BATCH_ID"))
			End If
			If Not IsDBNul1(_TableB.Rows(i) ("DATA_ID")) Then
				_DidList.Add(_TableB.Rows(i) ("DATA_ID"))
			End If
		Next
		If Not _BidList.Count = _DidList.Count Then
			eCode = "b Get_Data_LOT_DATA_TBL Not _BidList.Count = _DidList.Count "& _BidList.Count & "  "& _DidList.Count
			Dim g As Integer = "w"
		End If
		For i As Integer = 0 To_Bidlist.Count - 1
			If InStr(_DidList(i), _BidList(i)) = 0 Then
				eCode = "c Get_Data_LOT_DATA_ȚBL InStr(_DidList(i), _BidList(i)) = 0 "& _BidList(i) &" "& _Didlist(i)
				Dim g As Integer = "w"
			End If
		Next
		eCode = "d Get_Data_LOT_DATA_TBL"
		_F_BATCHID =_BidList(0)
		_E_BATCHID =_BidList(_BidList.Count - 1)
		_Tablel.Clear()	
		Dim CL As New SqlConnection
		Dim SQLCmL As SqlCommand = CnL.CreateCommand
		Dim AdapterL As New SqlDataAdapter(SQLCmL)
		Dim strQLL As String = "”
		CnL.ConnectionString - StrServerConnection_RIST
		SQLCmL. CommandTimeout = 600

		strSQLL = "SELECT "
		strSQLL &= "BATCH_ID,"
		strSQLL &= "LOT NO,"
		strSQLL &= "PRODUCT, "
		strSQLL &= "PROCESS, "
		strSQLL &= "OPE_NAME,"
		strSQLL &= "L_PARAM2"
		strSQLL &= " FROM LOT_TBL”
		strSQLL &= " WHERE "
		sDt = DateTime.Now
		For i As Integer = 0 to _BidList.Count -1
			StrSQLL &= “BATCH_ID = ” & “ ‘ “ & _BidList(i) & “‘“
			If Not i = _BidList.Count - 1 Then 
				strSQLL &= " OR "
			End If
		Next
		strSQLL &= " ORDER BY BATCH_ID ASC”
		eDt = DateTime.Now
		eCode - "e Get_Data_LOT_DATA_TBL" & Environment.NewLine & strSQLL
		SQLCmL.CommandText = strSQLL
		AdapterL.Fill(_TableL)
		_TableL. Dispose() 
		AdapterL.Dispose()
		SQLCmL.Dispose ()
		CnL.Dispose ()
		eDt = DateTime. Now

		eCode = "f Get_Data_LOT_DATA_TBL " & Environment.NewLine & strSQLL
		_TableD. CLear()
		Dim CnD As New SqlConnection
		Dim SQLCD As SqlCommand = CnD. CreateCommand
		Dim AdapterD As New SqlDataAdapter (SQLCmD)
		Dim strSQLD As String = "”
		CnD.ConnectionString = StrServerConnection_RIST
		SQLCmD.CommandTimeout = 600

		strSQLD = "SELECT "
		strSQLD &= "DATA_ID,"
		strSQLD &= "MEASURE_ITEM,"
		strSQLD &= "MEASURE_POINT, "
		strSQLD &= "DATA, "
		strSQLD &= " FROM DATA_TBL”
		strSQLD &= " WHERE "
		For i As Integer = 0 To _DidList.Count - 1
			strSQLD &= "DATA_ID = "&" '"& _DidList(i) & "’"
			If Not i = _DidList.Count - 1 Then
				strSQLD &= " OR "
			End If
		Next
		strSQLD &= " ORDER BY DATA_ID ASC"
		eCode = "g Get_Data_LOT_DATA_TBL " & Environment.NewLine & strSQLD
		SQLCmD.CommandText = strSQLD
		AdapterD.Fill(_TableD)
		TableD.Dispose()
		AdapterD.Dispose()
		SQLCmD.Dispose()
		CnD.Dispose()
		eCode = "h Get_Data_LOT_DATA_TBL " & Environment.NewLine & strSQLD

		Dim buf(_TableL.Rows.Count - 1)
		Dim flag(_BidList.Count - 1) As Boolean
		If _BidList.Count = _TableL.Rows.Count Then 
		ElseIf _BidList.Count > _TableL.Rows.Count Then
			eCode = "cc Get_Data_LOT_DATA_TBL_BidList.Count > _TableL.Rows.Count
" & _BidList.Count & “ & _TableL.Rows.Count & " “ _F_BATCHID & “ " & _E_BATCHID
			For i As Integer = 0 To UBound (buf, 1)
				buf (i) = _TableL.Rows(i) ("BATCH_ID")
			Next

			For i As Integer = 0 To _BidList.Count - 1
				flag (i) = False
				For j As Integer = 0 To UBound (buf, 1)
					If_BidList (i) = buf(j) Then
						flag (i) = True
						Exit For
					End If
				Next
			Next
			For i As Integer = _BidList.Count - 1 To 0 Step -1
				If flag(i) = False Then
					_BidList. RemoveAt (i) 
					_DidList. RemoveAt (i)
				End If
			Next
			ElseIf_BidList.Count < _TableL.Rows. Count Then
				eCode = "cc Get_Data_LOT_DATA_TBL_BidList.Count < _Tablel.Rows.Count " & _BidList.Count &" "& _TableL.Rows.Count & " "& _F_BATCHID & " " & _E_BATCHID
				Dim g As Integer = "W"
			End If
		End Sub
		Public Sub Get_Data_Buf(ByRef _TableB As DataTable, ByRef _TableL As DataTable, ByRef_TableD As DataTable, ByRef_DataBuf(,) As String, ByVal _MaxID As Integer, ByVal _i As Integer, ByVal _BidList As List(Of String), ByVal _DidList As List(Of String))
			Backgroundworker1.ReportProgress(6)
			ReDim _DataBuf(_BidList.Count - 1, 126 - 1)
			For i As Integer = 0 To UBound(_DataBuf, 1)
				For j As Integer = 0 To UBound(_DataBuf, 2)
					_DataBuf (1, j) = "Null"
				Next
			Next
			Dim Count_Data(_DidList.Count - 1) As Integer
			sDt = DateTime.Now
			Dim buf0(_TableB.Rows.Count - 1, 4 - 1) As String
			For j As Integer = e To UBound (bufe, 1)
				eCode = "a Get Data_Buf " & j & " " & UBound(buf0, 1)
				buf0(j, 0) = "Null"
				buf0(j, 1) = "Null"
				buf0(j, 2) = "Null"
				buf0(j, 3) = "Null"
				If Not IsDBNull(_TableB(j) ("BATCH_ID")) Then
					buf0(j, 0) = _TableB(j)("BATCH_ID")
				End If
				If Not IsDBNull( TableB(j) ("EQP_NO")) Then 
					buf0(J, 1) =_TableB(j) ("EQP_NO")
				End if
				If Not IsDBNull(_TableB(j) ("RECIPE")) Then
					buf0 (j, 2) = _TableB (j) ("RECIPE")
				End If
				If Not IsDBNulI(TableB.Rows (j) ("START_TIME")) Then
					buf0(j, 3) = Format (_TableB.Rows(j)("START_TIME"), "yyyy-MM-dd HH: mm:ss. fff")
				End If
			Next
			Dim buf1(_Table.Rows. Count - 1, 5 - 1) As String
			For j As Integer = 0 To UBound (buf1, 1)
				eCode = "b Get Data_Buf " & j & " "& UBound (buf1, 1)
				buf1(j, 0) = "Null"
				buf1(j, 1) = "Null"
				buf1j, 2) = "Null"
				buf1(j, 3) = "Null"
				buf1(j, 4) = "Null"
				If Not IsDBNulI(_TableL(j) ("BATCH_ID"')) Then
					buf1(j, 0) = _TableL(İ) ("BATCH_ID"')
				End If
				If Not IsDBNulI(Table.Rows(j) ("PROCESS" )) Then
					buf1(j, 1) =_Table. Rows (j) ("PROCESS")
				End If
				If Not IsDBNull_Table. Rows (j) ("LOT_NO")) Then.
					buf1(j, 2) =_Table.Rows(j) ("LOT_NO")
				End If
				If Not IsDBNuII(_Tablel.Rows(j) ("L_PARAM2")) Then
					buf1(j, 3) = _Table.Rows(j) ("L_PARAM2")
				End If
				If Not IsDBNull_Table.Rows(j) ("OPE_NAME")) Then
					buf1(j, 4) - _Table. Rows(J) ("OPE_NAME")
				End If
			Next
			_TableB.Clear()
			_TableL.Clear()
			For i As Integer = 0 To _BidList.Count - 1
				For j As Integer = 0 To UBound (buf0, 1)
					eCode = "c Get_Data_Buf " & i & " “ & _BidList.Count -1 & “ “ & UBound(buf0, 1)
					If Not buf0(j, 0) = "Null" Then
						If _BidList(i) = buf0(j, 0) Then
							If Not buf0(j, 1) = "Null" Then
								_DataBuf(1, 2) = buf0(j, 1)
							End If
							If Not buf0(j, 2) = "Null" Then
								_DataBuf (i, 4) = buf0(j, 2)
							End If
							If Not buf0(J, 3) = "Null" Then
								_DataBuf(i, 18) = buf0(j, 3)
							End If
							Exit For
						End If
					End If
				Next
				For j As Integer = 0 To UBound (buf1, 1)
					eCode = "d Get_Data_Buf " & i & " " & j & " " & _BidList.Count - 1 & " " & UBound(buf1, 1)
					If Not buf1(j, 0) = "Null" Then
						If _Bidlist(1) = buf1(j, 0) Then•
							If Not buf1(j, 1) = "Null" Then
								_DataBuf (i, 1) = buf1(j, 1) '[cProcessName] ‹= PROCESS
							End If
							If Not buf1(j, 2) = "Null" Then
								_DataBuf(i, 14) = buf1(j, 2) [cLotNo] <= LOT_NO"
							End If
							If Not buf1(j, 3) = "Null" Then
								_DataBuf (i, 17) = buf1(j, 3) [cInCharge] ‹= L_PARAM2
							End If
							If Not buf1(j, 4) = "Null" Then
								_DataBuf (1, 5) = buf1(j, 4) '[cFilter_2] <= OPE_NAME
							End If
							Exit For
						End If
					End If
				Next
			Next
			eDt = DateTime.Now
			ts(2) = eDt - sDt
			Backgroundworker1.ReportProgress(7) 
			Backgroundworker1.ReportProgress(_H)
			sDt = DateTime.Now
			Dim buf2(_TableD.Rows.Count - 1, 3 - 1) As String
			For i As Integer = 0 To UBound(buf2, 1)
				eCode = "e Get Data_Buf " & i & " " & UBound(buf2, 1)
				buf2(i, 0) = "Null"
				buf2(i, 1) = "Null"
				buf2(i, 2) = "Null"
				If Not IsDBNull(_TableD(i) ("DATA_ID")) Then
					buf2(1, 0) =_TableD(i) ("DATA_ID")
				End If
				If Not IsDBNull(_TableD(1) ("MESURE_ITEM")) Then
					buf2(i, 1) =_TableD(i) ("MESURE_ITEM")
				End If
				If Not IsDBNull(_TableD(1) ("DATA")) Then
					buf2(1, 2) = TableD(i) ("DATA")
				End If
			Next
			_TableD. Clear()
			For i As Integer = 0 To DidList. Count - 1
				Count_Data (i) = 0
				For j As Integer = 0 To UBound (buf2, 1)
					eCode = "f Get Data_Buf " & i &" "& j & " " &_DidList.Count - 1 & " " & UBound(buf2, 1)
					If Not buf2 (j, 0) = "Null" Then
						If _Didlist(i) = buf2(j, 0) Then 
							If Count_Data(i) = 0 Then
								If Not buf2 (j, 1) = "Null" Then
									_DataBuf (i, 3) = buf2(j, 1)
								End If
							End If
							If Not buf2(j, 2) = "Null" Then
								_DataBuf (i, 26 + Count_Data(i)) = buf2(j, 2)
								Count_Data (i) += 1
							End If
						End If
					End If
				Next
			Next
			Dim max As Double = -100000
			Dim min As Double = 100000
			Dim sum As Double = 0
			Dim c As Integer = 0
	
			For i As Integer = 0 To UBound(_DataBuf, 1)
				_DataBuf (i, 0) = _MaxID + i [iID]
				
				If Not _DataBuf (i, 26) - "Null" Then
					For j As Integer = 0 To 100 - 1
						eCode = "g Get_Data_Buf " & i & " " & j & " “ & UBound(_DataBuf, 1) & “ “ & 100 - 1
						If _DataBuf(i, 26 + j) = "Null" Then
							Exit For
						End If
						If j = 0 Then
							max = _DataBuf(i, 26 + j)
							min = _DataBuf(i, 26 + j)
						Else
							If max<_DataBuf(i, 26+ j) Then
								max =_DataBuf(1, 26 + j)
							End If
							If min >DataBuf(i, 26+ J) Then
								min= DataBuf (1, 26 + j)
							End If
						End If
						sum += _DataBuf（1, 26 +j）
						c += 1
					Next
					eCode = "h Get_Data_Buf "
					_DataBuf (i, 22) = Mid (sum / c, 1, 10)
					_DataBuf (1, 23) = Mid(max - min, 1, 10)
					_DataBuf (i, 24) = max
					_DataBuf (i, 25) = min
				Else
					_DataBuf (i, 22) = 0
					_DataBuf (i, 23) = 0
					_DataBuf (1, 24) = 0
					_DataBuf (i, 25) = 0
				End If
				sum = 0
				c = 0
				max = -100000
				min = 100000
			Next
			eDt = DateTime.Now
			ts (3) = eDt - sDt
		End Sub
		Public Function Write_Data(ByVal _Data(,) As String) As String
			Write_Data - ""
			Dim Cn As New SqlConnection
			Dim strQL As String ="”
			Dim SQLCm As SqlCommand = Cn.CreateCommand
			Dim trans As SqlTransaction = Nothing
			Try
			Cn.ConnectionString = StrServerConnection_Yuku

			For i As Integer = 0 To UBound (Data, 1)

				Cn.Open()
				trans = Cn.BeginTransaction
				SQLCm. Transaction - trans
				strSQL = ""
				strSQL = "INSERT INTO " & "SPC_Master" & " VALUES ("

				For j As Integer = 0 To UBound(_Data, 2)
					eCode = "a Write_Data " & i & " " & j& " " & UBound(_Data, 1) & " " & UBound(_Data, 2)
					If _Data (i, j) = "Null" Then
						strSQL &= "Null"
					Else 
						strSQL &= "'" & _Data (i, j) & "'"
					End IF
					If j = UBound (_Data, 2) Then
						strSQL &= ")"
					Else 
						strSQL &= ","
					End If
				Next
				eCode = "b Write_Data " & strSQL
				SQLCm. CommandText = strSQL
				SQLCm. ExecuteNonQuery () 
				trans. Commit()
				Cn. Close ()
			Next
		Catch ex As Exception
			If IsNothing(trans) = False Then
				trans. Rollback()
			End If
			OKflag = False
			Cn. Close()
			Write_Error(">›" & Format(DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error " & eCode & Environment.NewLine & ex.Message & Environment. NewLine & ex.StackTrace)
			Write_Data = "error " & eCode & Environment. NewLine & ex.Message & Environment.NewLine & ex.StackTrace
		End Try

	End Function
	Public Sub Write_History() 
		Dim TextFile As IO.Streamwriter - New IO. StreamWriter(StrCDir & "\" & "DebugFile.txt", True, System. Text.Encoding.Default)
		Dim Str As String = ""
		Str="”
		TextFile.Write(Str)
		TextFile. WriteLine()
		TextFile. Close()
	End Sub
	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System. EventArgs)
Handles Button2.Click
		If StrServerConnection_RIST = “” Then
			MsgBox("Set the RIST server connection string.")
			Exit Sub
		End If
		If StrServerConnection_Yuku = “” Then
			MsgBox("Set the Yukuhashi server connection string.")
			Exit Sub
		End If
		OKflag = True
		Form3.Show()
		Form3.Label1.Text - "Table creating..."
		Try
			Input_CSV_to_Server(StrCDir & " \SPC_Alarm.csv", 0)
			Input_CSV_to_Server(StrCDir & " \SPC_Master.csv", 0)
			Input_CSV_to_Server (StrDir & " (SPC_Property. csv",0)
			Input_CSV_to_Server (StrDir & " \SPC_User. csv"; 0)
		Catch ex As Exception
			OKflag = False
		End Try
		Form3. Close()
		If OKflag = True Then
			MsgBox ("Table creation OK")
		Else
			MsgBox("Table creation  NG")
		End If
	End Sub
	Public Function Input_CSV_to_Server(ByVal CsvName As String, ByVal Del As Integer) As String
		Input_SV_to_Server = ""
		Dim CSVData(0, 0) As String 
		Dim CSVData_Header(0, 0) As String
		Dim TableName As String = " “
		Input_SV_to_Server = Input_CSV(CsvName, CSVData, CSVData_Header, TableName)
		If Not Input_CSV_to_Server = " “Then 
			Exit Function
		End If
		Input_CSV_to_Server = To_Server(CSVData, CSVData_Header, TableName, Del)
		If Not Input_CSV_to_Server = "" Then
			Exit Function
		End If
	End Function
	Public Function Input_CSV(ByVal Filename As String, ByRef Array(,) As String,
ByRef Array_Header(,) As String, ByRef TName As String) As String
		Input_SV = “”
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
				temp0 - Split(sr0.ReadLine(), ",")
				If CoMax < temp0.Length Then 
					CoMax = temp0.Length
				End If
				Gyo += 1
			Loop
			sr0. Close()
			eCode = "b"
			ReDim Buf (Gyo - 1, CoMax - 1)
			Gyo = e
			sr0 = New System.IO. StreamReader (Filename, System.Text.Encoding.Default)
			Do Until sr0.Peek() = -1
				temp0 = Split (sre.ReadLine(), ",”)
				For i As Integer = 0 To CoMax - 1
					Buf (Gyo, i) = temp0(i)
				Next
				Gyo += 1
			Loop
			sr0. Close ()
			eCode = "c”
			ReDim Array_Header(3 - 1, UBound (Buf, 2) - 1)
			ReDim ArrayUBound (Buf, 1) - (UBound (Array_Header, 1) + 1) - 1, UBound (Buf, 2) - 1)
			TName = Buf (0, 1)
			For i As Integer = 1 To 1 + UBound(Array_Header, 1)
				For j As Integer = 1 To UBound (Buf, 2)
					Array_Header(i - 1, j - 1) = Buf (i, j)
				Next
			Next
			eCode = "d"
			For i As Integer = 1 + UBound(Array_Header, 1) + 1 To UBound (Buf, 1)
				For j As Integer = 1 To UBound (Buf, 2)
					Array(i - (1 + UBound (Array_Header, 1) + 1), j - 1) = Buf (i, j)
				Next
			Next	
			eCode = "e"
		Catch ex As Exception
			Input_CSV = "Input_CSV " & eCode & Environment.NewLine & ex.Message & Environment.NewLine & ex.StackTrace
		End Try
	End Function
	Public Function To_Server(ByVal Array(,) As String, ByVal Array_Header(,) As String, ByVal TName As Strings, ByVal Del As Integer) As String
	To_Server = “”
	Dim eCode As String = “”
	Dim Cn As New SqlConnection
	Dim strSQL As String = “”
	Dim SQLCm As SqlCommand = Cn.CreateCommand
	Dim trans As SqlTransaction = Nothing
	Try
		eCode = "a"
		Cn. ConnectionString = StrServerConnection_Yuku
		Cn. Open ()
		trans = Cn. BeginTransaction
		SQLCm. Transaction = trans
		If 1 Then
			eCode = "b"
			If Check_TableRist（TName）= True Then
				If Del = 0 Then
					MsgBox ("Table:" & TName & " already exists.")
					OKflag = False
					Cn. Close()
					Exit Function
				End If
				eCode = "c"
				SQL Cm. CommandText = "DROP TABLE " & TName
				SQLCm. ExecuteNonQuery ()
			End If
			eCode = "d"
			strSQL = "Create Table " & TName & "("
			For j As Integer = 0 To UBound (Array_Header, 2)
				strSQL &= Array_Header(0, j) & " " & Array_Header(1, j)
				If Array_Header(2, j) = "NO" Then
					strSQL &= " Not Null"
				End If
				If j= UBound (Array_Header, 2) Then
					strSQL &= ",PRIMARY KEY (" & Array_Header(0, 0)& ")"
					strSQL &= ")"
				Else
					strSQL &- ","
				End If
			Next
			eCode = "e"
			SQLCm. CommandText = strSQL
			SQLCm. ExecuteNonQuery)
			eCode = "f"
			SQLCm. CommandText = "DELETE FROM " & TName
			SQLCm. ExecuteNonQuery ()
			eCode = "g"
			For i As Integer = 0 To UBound (Array, 1)
				strSQL = "”
				strSQL - "INSERT INTO " & TName & " VALUES ("
				For j As Integer = 0 To UBound (Array,2)
					If Array(i, i) = "Null" Then
						strSQL &= "Null"
					Else
						strSQL &= "'" & Array(i, j) &"’”
					End If
					If j = UBound (Array, 2) Then 
						strSQL &=")”
					Else
						strSQL &= ","
					End If
					eCode = "h" & i&" " & j
				Next
				SQLCm. CommandText = strSQL
				SQLCm. ExecuteNonQuery ()
			Next
		End If
		trans. Commit()
		Cn. Close()
	Catch ex As Exception
		If IsNothing(trans) = False Then
			trans. Rollback ()
		End If
		OKflag = False
		To_Server = "To_Server " & eCode & Environment.NewLine & strSQL & Environment.NewLine & ex.Message & Environment.NewLine & ex. StackTrace
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
		Cn. ConnectionString = StrServer Connection_Yuku
		Cn.Open()
		dtColumns = Cn.GetSchema("Columns")
		Cn. close()
		Cn.Dispose ()
		Make_TableRist = dtColumns.Rows. Count
		If Make_TableRist = 0 Then
			Exit Function
		End If
		ReDim Buf (dtColumns. Rows. Count - 1)
		For i As Integer = 0 To dtColumns. Rows. Count - 1
			Buf(i) = dtColumns.Rows (i) ("TABLE_NAME")
		Next
		Dim al2 As New System. Collections. ArrayList(UBound (Buf, 1))
		Dim j2 As String
		For Each j2 In Buf
			If Not al2. Contains (j2) Then
				al2. Add (j2)
			End If
		Next
		TableRist = DirectCast(al2.ToArray(GetType(String)), String())
	Catch ex As System.Exception
		Cn. Dispose()
		MsgBox（"グラフ設定データ取得エラー”+”,”+ ex. Message & ex.StackTrace)
		Exit Function
	End Try
End Function
Public Function Check_TableRist(ByVal TName As String) As Boolean
	Check_TableRist - False
	If Make_TableRist() = 0 Then
		Exit Function
	End If
	For i As Integer = 0 To UBound (TableRist, 1)
		If TableRist(i) = TName Then
			Check TableRist = True
			Exit Function
		End If
	Next
End Function
Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
Handles Button3.Click
	Dim dr As DialogResult
	Dim frm As New Form2
	dr = frm.ShowDialog
End Sub

Dim MaxID As Integer = 0
Dim gRows As Integer=1000
Private Sub Write_Error(ByVal strE As String)
	Dim TextFile As IO.Streamwriter = New IO. Streamwriter (StrCDir & "\Emsg.txt", True, System. Text.Encoding.Default)
	TextFile. Write(strE)
	TextFile. WriteLine()
	TextFile. Close()

	LineNotify ("แจ้งเตือนสถานะ SPC Converter ERROR"+ vbNewLine + "ประจำวันที่: " + Now.ToString("yyyy-MM-dd HH:mm:ss"))
End Sub
Dim recDay As String = ""
Dim F_BATCHID As String = "”
Dim E_BATCHID As String =""
Dim Bidlist, As New List(Of String)
Private Sub Backgroundworker1_Dowork(ByVal sender As System.Object,
ByVal e As System.ComponentModel.DoworkEventArgs), Handles Backgroundworker1.Dowork
	Dim TableB As New DataTable 
	Dim TableL As New DataTable 
	Dim TableD As New DataTable
	Dim DataBuf(0, 126 - 1) As String
	Dim buf As String = "”
	Dim DidList As New List(Of String)
	Backgroundworker1.ReportProgress(Display_Clear)
	Backgroundworker1.ReportProgress (_A)
	Try
		recDay = Get_Recently_Day()
	Catch As Exception
		Write_Error(">>" & Format (DateTime.Now, "yyyy/MM/dd HH:mm:ss") & " error "
& eCode & Environment .NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
		Backgroundworker1.ReportProgress(9)
		Exit Sub
	End Try
	If Backgroundworker1.CancellationPending Then
		Exit Sub
	End If
	Backgroundworker1. ReportProgress(_B)
	Try
		buf = Get_NextMasterNo2（“SPC_Master"， Strservenconnection_Yuku）
	Catch ex As Exception
		Write_Error (">" & Format(DateTime. Now, "yyyy/MM/dd HH: mm: ss") & " error " & ecode & Environment. NewLine & ex. Message & Environment. NewLine & ex. StackTrace)
		Backgroundworker1. ReportProgress (9)
		Exit Sub
	End Try
	If Not buf = -1 Then
			MaxID = CInt (buf) + 1 
	Else
		Write_Error(">>"& Format(DateTime.Now,"yyyy/MM/dd HH: mm: ss") & " buf=-1 ")
		Backgroundworker1. ReportProgress (9)
		Exit Sub
	End If
	If Backgroundworker1. CancellationPending Then
		Exit Sub
	End If
	eCode. = "d"
	Backgroundworker1. ReportProgress (Set_Rows)
	Backgroundworker1. ReportProgress(_D)
	Do
		write_times += 1
		If write_times ＞100 Then
			Write_times = 0
			Backgroundworker1.ReportProgress(_K)
		End If
		If Backgroundworker1. CancellationPending Then
			Exit Do
		End If
		SDt_All - DateTime.Now
		Backgroundworker1. ReportProgress(4)
		Backgroundworker1. ReportProgress(_E)
		sDt = DateTime.Now
		eCode = "e"
		Try
			Get_Data_BATCH_TBL（TableB， gRows, recDay）
		Catch ex As Exception
			Write_Error(">" & Format(DateTime. Now, "yyyy/MM/dd HH: mm: ss") & " error " & ecode & Environment. NewLine & ex. Message & Environment. NewLine & ex. StackTrace)
			Backgroundworker1. ReportProgress (9)
			Exit Sub
		End Try
		eDt = DateTime. Now
		ts (0) = eDt - sDt
		If TableB. Rows .Count = 0 Then
			Backgroundworker1. ReportProgress (_C)
			For i As Integer = 0 To 60 * 5 - 1
				System. Threading. Thread. Sleep (1000)
				If Backgroundworker1. CancellationPending Then
					Exit Do
				End If
			Next
			Continue Do
		End If
		If Backgroundworker1. CancellationPending Then
			Exit Do
		End If
		Backgroundworker1. ReportProgress (5)
		Backgroundworker1. ReportProgress(_F)
		SDt = DateTime.Now
		eCode = "f"
		Try
			Get_Data_LOT_DATA_TBL(TabbleB, TableL, TableD, BidList, F_BATCH, E_BATCH)
		Catch ex As Exception
			Write_Error("»›" & Format (DateTime. Now, "yyyy/MM/dd HH: mm: ss") & " error " & ecode & Environment. NewLine & ex.Message & Environment.NewLine & ex.StackTrace)
			Backgroundworker1. ReportProgress(9)
			Exit Sub
		End Try
		eDt = DateTime.Now
		ts (1) = eDt - sDt
		If Backgroundworker1. CancellationPending Then
			Exit Do
		End If
		Backgroundworker1. ReportProgress (_G)
		Try
			Get_Data_Buf(TableB, TableL, TableD, DataBuf, MaxID, 0, BidList, DidList)
		Catch ex As Exception
			Write_Error(">" & Format (DateTime. Now, "yyyy/MM/dd HH: mm: ss"). & " error " & ecode & Environment. NewLine & ex.Message & Environment. NewLine & ex.StackTrace)
			Backgroundworker1. ReportProgress (9)
			Exit Sub
		End Try
		If Backgroundworker1. CancellationPending Then
			Exit Do
		End If
		SDt = DateTime.Now
		Dflag = True
		Backgroundworker1. ReportProgress (8) 
		Backgroundworker1. ReportProgress (_I)
		sDt = DateTime.Now
		If Not Write_Data(DataBuf) = "” Then
			Backgroundworker1. ReportProgress (9)
			Exit Sub
		End If
		eCode = "s"
		Dflag = False
		eDt = DateTime.Now
		ts (4) = eDt - sDt	
		Backgroundworker1. ReportProgress(_J)
		eDt_All = DateTime.Now
		ts_AIl = eDt_AIl - SDt_AIl
		MaxID += BidList. Count
		Backgroundworker1. ReportProgress(Display_Seconds)
		recDay = DataBuf (UBound (DataBuf, 1), 18)
		If Backgroundworker1. CancellationPending Then
			Exit Do
		End If
	Loop
End Sub
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

Private Sub Backgroundworker1_ProgressChanged (ByVal sender As Object, ByVale As System. ComponentModel. ProgressChangedEventArgs) Handles Backgroundworker1. ProgressChanged
	If e. Progresspercentage = Display_Clear Then
		ForeColorBlack()
		Form3. Label1. Text = "Preparing..."
		Form3. TextBox1. Text = ""
	ElseIf e. ProgressPercentage = Set_Rows Then
		Form3. Label1. Text = “Processing...”
	ElseIf e. ProgressPercentage = Display_Seconds Then
	ElseIf e. ProgressPercentage = Stop_Convert Then
		Form3. Label1. Text = "Stopping...”
		Form3. TextBox1. Text &- Format(DateTime. Now, "MM/dd HH: mm: ss") & " > Stopping..." & Environment.NewLine
		TextScroll()
	ElseIf e.ProgressPercentage = 9 Then
		Form4. TopMost = True
		Form4. Show()
	ElseIf e.ProgressPercentage =_A Then
		Form3.TextBox1. Text &= Format (DateTime. Now,"MM/dd HH: mm: ss") & " > Getting latest date of data from the SPC_Master." & Environment.NewLine
	ElseIf e.ProgressPercentage =_B Then
		Form3. TextBox1. Text &= Format(DateTime. Now,"MM/dd HH: mm: ss") & " > Getting number of data in the SPC_Master." & Environment.NewLine
	ElseIf e.ProgressPercentage = _C Then
		Form3. TextBox1. Text &= Format (DateTime.Now,"MM/dd HH: mm: ss") & " > No latest data." & Environment.NewLine
		Form3. TextBox1. Text &= Format (DateTime. Now,"MM/dd HH: mm: ss") & " > Wait for 5 minutes." & Environment.NewLine
	ElseIf e.ProgressPercentage =_D Then
		Form3. TextBox1. Text &= Format(DateTime. Now, "MM/dd HH: mm: ss") & " > Start conversion." & Environment.NewLine
	ElseIf e.ProgressPercentage = _E Then
		Form3. TextBox1.Text &= "--———————————————————————————————————-" & Environment.NewLine
		Form3. TextBox1. Text &= Format (DateTime.Now, "MM/dd HH: mm:ss") & " › Getting data in the BATCH_TBL. ( After " & recDay & ")" & Environment.NewLine
		Form3.TextBox1.SelectionStart = Form3. TextBox1. Text. Length
		TextScroll()
	ElseIf e.ProgressPercentage =_F Then
		Form3. TextBox1. Text &= Format (DateTime.Now, "MM/dd HH: mm: ss") & " > Getting data in the DATA_TBL and LOT_TBL."& Environment. NewLine
		TextScroll()
	ElseIf e. ProgressPercentage = _G Then
		Form3. TextBox1. Text &= Format (DateTime. Now, "MM/dd HH: mm: ss") & " > organizing data in the BATCH_TBL and LOT_TBL." & Environment.NewLine
		TextScroll()
	ElseIf e. ProgressPercentage = _H Then
		Form3. TextBox1. Text &= Format (DateTime.Now, "MM/dd HH: mm: ss") & " › Organizing data in the DATA_TBL." & Environment.NewLine
		TextScro1l()
	ElseIf e.ProgressPercentage = _I Then
		Form3. TextBox1. Text &= Format(DateTime. Now, "MM/dd HH; mm:ss") & " > Writing data." & Environment.NewLine
		TextScroll()
	ElseIf e.ProgressPercentage =_J Then
		Form3. TextBox1.Text &- Format(DateTime. Now, "MM/dd HH: mm: ss") & " > Conversion successful. ( BATCH_ID [" & F_BATCHID & "] to [" & E_BATCHID & "], "& BidList.Count & "rows )" & Environment.NewLine
		TextScroll()
		Form3. TextBox1.Text &- Format(DateTime. Now,"MM/dd HH: mm: ss") & " > (Elapsed time: "& Math.Round(ts(0). TotalSeconds, 2, MidpointRounding.AwayFromZero)& "s, "& Math.Round(ts (1). TotalSeconds, 2,MidpointRounding.AwayFromZero)& "s, "& Math.Round(ts (2). TotalSeconds, 2,MidpointRounding.AwayFromZero)& "s, "& Math.Round(ts (3). TotalSeconds, 2,MidpointRounding.AwayFromZero)& "s, "& Math.Round(ts (4). TotalSeconds, 2,MidpointRounding.AwayFromZero)& "s / Total: " & Math.Round (ts_All. TotalSeconds, 2,
MidpointRounding.AwayFromZero)& "s " & Math.Round (BidList.Count / ts_All. TotalMinutes, 0,MidpointRounding.AwayFromZero) & "row/min" & " )" & Environment. NewLine
		TextScroll()
	Elself e. ProgressPercentage = _K Then
		Form3. TextBox1. Text =“”
	End If
End Sub
Public Sub ForeColorBlack()
	Form3. Label5. ForeColor= Color.Black
	Form3. Label6. ForeColor = Color.Black
	Form3. Label7. ForeColor = Color.Black
	Form3. Label8. ForeColor = Color.Black
	Form3. Label9. ForeColor = Color.Black
End Sub
Public Sub TextScroll()
	Form3.TextBox1.SelectionStart = Form3.TextBox1. Text. Length
	Form3. TextBox1. Focus ()
	Form3. TextBox1. ScrollToCaret()
EndSub
Private Sub Backgroundworker1_RunworkerCompleted(ByVal sender As Object, ByVal e As System. ComponentModel. RunworkerCompletedEventArgs) Handles Backgroundworker1.RunworkerCompleted
	Form3. Close()
End Sub
Private Sub Button4_Click(ByVal sender As System. Object, Byval e As System.EventArgs) Handles Button4.Click
	Dim result As DialogResult
	result = MessageBox. Show("Do you want to delete the data in the SPC_Master ?", "Delete Data", MessageBoxButtons.OKCancel)
	If result = Windows. Forms.DialogResult.OK Then
		eMsg = Input_CSV_to_Server(StrCDir & "\SPC_Master.csv", 1)
		If Not eMsg=“” Then
			Write_Error"error SPC_Master_Recreated"& eMsg)
			Form4. Show()
			Exit Sub
		End If
		MsgBox ("Deleted the data in the SPC_Master.")
	End If
End Sub
Private Sub LineNotify(ByVal msg As String)
	Dim token As String = My.Settings.LineToken
	Try
		Dim request As HttplebRequest = DirectCast(WebRequest.Create("https://notify-api.line.me/api/notifx"), HttpWebRequest)
		Dim postData - String.Format ("message-(0}", msg)
		Dim data = Encoding.UTF8.GetBytes(postData)
		request.Method = "POST"
		request. ContentType = "application/x-www-form-urlencoded"
		request. ContentLength = data. Length
		request. Headers.Add ("Authorization", "Bearer " + token) 					 
		request.AllowwriteStreamBuffering = True
		request.KeepAlive = False
		request.Credentials = CredentialCache.DefaultCredentials
		Using stream = request. GetRequestStream()
			stream. Write(data, 0, data. Length)
		End Using
		Dim response As HttpwebResponse = CType(request. GetResponse(), HttpWebResponse)
		Dim responseString = New StreamReader(response. GetResponseStream)) . ReadToEnd()
	Catch ex As Exception
		Dim LogFilePath As String = My.Application. Info.DirectoryPath + "/" + My.Settings.LineLog
		Using FileLog As New FileStream(LogFilePath, FileMode.Append, FileAccess. Write)
			Using WriteFile As New Streamwriter (FileLog)
				WriteFile.WriteLine(DateTime.Now)
				WriteFile.WriteLine(ex.Message)
				WriteFile WriteLine(" (0} Exception caught.", ex)
				WriteFile WriteLine("-——————————————————————————————————————“)
			End Using
		End Using
	End Try
End Sub
Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5. Click
	LineNotify("แจ้งเตือนสถานะ SPC Converter ERROR" + vbNewLine + _"ประจำวันที่: " + Now.ToString("yyyy-MM-dd HH: mm : ss"))
	End Sub
End Class

