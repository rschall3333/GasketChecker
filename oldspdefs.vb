Option Strict Off
Option Explicit On
Module oldspdefs
	
	'DSP functions
	Declare Sub olDspGetErrorString Lib "oldsp32.dll" (ByVal ecode As Integer, ByVal ErrMsg As String, ByVal maxsize As Short)
	Declare Function olDspRealFFT Lib "oldsp32.dll" (ByVal inVals As Integer, ByVal mag As Integer, ByVal phase As Integer, ByVal fft_size As Integer) As Integer
	Declare Function olDspWindow Lib "oldsp32.dll" (ByVal inVals As Integer, ByVal out As Integer, ByVal WindowSize As Short, ByVal WindowType As Short, ByVal DataType As Short, ByVal InputOffset As Integer) As Integer
	Declare Function olDspMagToDB Lib "oldsp32.dll" (ByVal inVals As Integer, ByVal out As Integer) As Integer
	Declare Function olDspInitFFT Lib "oldsp32.dll" () As Integer 'can be used to pre-init FFT coefficients, not required
	Declare Function olDspVoltsToOutput Lib "oldsp32.dll" (ByVal hBufVolts As Integer, ByVal hBufOut As Integer, ByVal DataType As Integer, ByVal MinRange As Single, ByVal MaxRange As Single, ByVal bits As Short) As Integer
	Declare Function olDspInputToVolts Lib "oldsp32.dll" (ByVal hBufIn As Integer, ByVal hBufVolts As Integer, ByVal DataType As Integer, ByVal MinRange As Single, ByVal MaxRange As Single, ByVal bits As Short) As Integer
	
	'OLDSP ERROR CODES
	
	Public Const OLDSP_CANT_ALLOC_MEMORY As Short = 300 'Can not allocate memory for coefficient tables
	Public Const OLDSP_FFT_SIZE_TOO_LARGE As Short = 301 'FFT size too large
	Public Const OLDSP_BAD_INPUT_DATA_WIDTH As Short = 302 'Illegal input data width
	Public Const OLDSP_BAD_FFT_SIZE As Short = 303 'Illegal FFT size
	Public Const OLDSP_NOT_ENOUGH_VALID_INPUT As Short = 304 'Insufficient input data
	Public Const OLDSP_WINDOW_SIZE_TOO_LARGE As Short = 305 'Window size too large
	Public Const OLDSP_BAD_DATA_TYPE As Short = 306 'Illegal data type
	Public Const OLDSP_BAD_WINDOW_TYPE As Short = 307 'Illegal window type
	Public Const OLDSP_NEGATIVE_INPUT_MAGNITUDE As Short = 308 'Illegal input: negative magnitude
	
	
	Sub DTOLDspTrap(ByRef FuncName As String, ByVal Errnum As Short)
		Dim ErrStg As New VB6.FixedLengthString(130)
		Const MB_ICONSTOP As Short = 16
		
		'Get DT-Open Layers Error string
		Call olDspGetErrorString(Errnum, ErrStg.Value, 129)
		'MsgBox ErrStg, MB_ICONSTOP, FuncName & " Error"
		Err.Raise(vbObjectError + Errnum, FuncName, ErrStg.Value)
		
		
	End Sub
	
	Sub VoltsToOutput(ByRef Subsys As AxDTAcq32Lib.AxDTAcq32, ByVal hBufVolts As Integer, ByVal hBufOut As Integer)
		Dim ecode As Integer
		
		ecode = olDspVoltsToOutput(hBufVolts, hBufOut, Subsys.Encoding, Subsys.MinRange, Subsys.MaxRange, Subsys.Resolution)
		If ecode <> 0 Then
			DTOLDspTrap("InputToVolts", ecode)
		End If
		
	End Sub
	
	Sub InputToVolts(ByRef Subsys As AxDTAcq32Lib.AxDTAcq32, ByVal hBufIn As Integer, ByVal hBufVolts As Integer)
		Dim ecode As Integer
		
		ecode = olDspInputToVolts(hBufIn, hBufVolts, Subsys.Encoding, Subsys.MinRange, Subsys.MaxRange, Subsys.Resolution)
		If ecode <> 0 Then
			DTOLDspTrap("InputToVolts", ecode)
		End If
		
	End Sub
	
	Sub Window(ByVal hBufIn As Integer, ByVal hBufOut As Integer, ByVal WindowSize As Short, ByVal WindowType As Short, ByVal DataType As Short, ByVal InputOffset As Integer)
		Dim ecode As Integer
		
		ecode = olDspWindow(hBufIn, hBufOut, WindowSize, WindowType, DataType, InputOffset)
		If ecode <> 0 Then
			DTOLDspTrap("Window", ecode)
		End If
		
	End Sub
	Sub RealFFT(ByVal hBufIn As Integer, ByVal hBufMag As Integer, ByVal hBufPhase As Integer, ByVal FFTSize As Short)
		Dim ecode As Integer
		
		ecode = olDspRealFFT(hBufIn, hBufMag, hBufPhase, FFTSize)
		If ecode <> 0 Then
			DTOLDspTrap("RealFFT", ecode)
		End If
		
	End Sub
	
	Sub MagToDB(ByVal hBufIn As Integer, ByVal hBufOut As Integer)
		Dim ecode As Integer
		
		ecode = olDspMagToDB(hBufIn, hBufOut)
		If ecode <> 0 Then
			DTOLDspTrap("MagToDB", ecode)
		End If
		
	End Sub
End Module