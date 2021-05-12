Option Strict Off
Option Explicit On
Module olmemdefs
	'time/date stamp
	Structure TDS
		Dim DtYear As Short
		Dim DtMonth As Short
		Dim DtDay As Short
		Dim DtHour As Short
		Dim min As Short
		Dim sec As Short
	End Structure
	'------------- ol Memory Management Function prototypes -----------'
	'get error string
	Declare Sub olDmGetErrorString Lib "OLMEM32.DLL" (ByVal ecode As Integer, ByVal ErrMsg As String, ByVal maxsize As Short)
	'get version info
	Declare Function olDmGetVersion Lib "OLMEM32.DLL" (ByVal Version As String, ByVal maxsize As Short) As Integer
	
	'Allocate buffers:
	Declare Function olDmAllocBuffer Lib "OLMEM32.DLL" (ByVal usWinFlags As Short, ByVal ulBufferSize As Integer, ByRef hBuffer As Integer) As Integer
	Declare Function olDmCallocBuffer Lib "OLMEM32.DLL" (ByVal usWinFlags As Short, ByVal ExFlags As Short, ByVal samples As Integer, ByVal SampSize As Short, ByRef hBuffer As Integer) As Integer
	
	'Free Buffer
	Declare Function olDmFreeBuffer Lib "OLMEM32.DLL" (ByVal hbuf As Integer) As Integer
	
	'Re-allocate (resize) buffer
	Declare Function olDmReAllocBuffer Lib "OLMEM32.DLL" (ByVal usWinFlags As Short, ByVal ulBufferSize As Integer, ByRef hBuffer As Integer) As Integer
	Declare Function olDmReCallocBuffer Lib "OLMEM32.DLL" (ByVal usWinFlags As Short, ByVal ExFlags As Short, ByVal samples As Integer, ByVal SampSize As Short, ByRef hBuffer As Integer) As Integer
	
	
	'Get physical size of Buffer
	Declare Function olDmGetBufferSize Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef lpUlBufferSize As Integer) As Integer
	
	'Copy VB array to buffer
	Declare Function olDmCopyToBuffer Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef vbarray As Short, ByVal Numsamples As Integer) As Integer
	Declare Function olDmCopyLongToBuffer Lib "OLMEM32.DLL"  Alias "olDmCopyToBuffer"(ByVal hbuf As Integer, ByRef vbarray As Integer, ByVal Numsamples As Integer) As Integer
	Declare Function olDmCopySingleToBuffer Lib "OLMEM32.DLL"  Alias "olDmCopyToBuffer"(ByVal hbuf As Integer, ByRef vbarray As Single, ByVal Numsamples As Integer) As Integer
	
	
	'Copy buffer data to VB array
	Declare Function olDmCopyFromBuffer Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef vbarray As Short, ByVal maxsamples As Integer) As Integer
	Declare Function olDmCopyLongFromBuffer Lib "OLMEM32.DLL"  Alias "olDmCopyFromBuffer"(ByVal hbuf As Integer, ByRef vbarray As Integer, ByVal maxsamples As Integer) As Integer
	Declare Function olDmCopySingleFromBuffer Lib "OLMEM32.DLL"  Alias "olDmCopyFromBuffer"(ByVal hbuf As Integer, ByRef vbarray As Single, ByVal maxsamples As Integer) As Integer
	
	
	'Get logical size of buffer
	Declare Function olDmGetValidSamples Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef size As Integer) As Integer
	Declare Function olDmGetMaxSamples Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef size As Integer) As Integer
	
	'Get size of sample in bytes
	Declare Function olDmGetDataWidth Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef size As Integer) As Integer
	
	'Get number of significant bits in each sample
	Declare Function olDmGetDataBits Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef size As Integer) As Integer
	
	'Set logical size of buffer
	Declare Function olDmSetValidSamples Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByVal size As Integer) As Integer
	
	'Get Time/Date Stamp of buffer
	'UPGRADE_WARNING: Structure TDS may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1050"'
	Declare Function olDmGetTimeDateStamp Lib "OLMEM32.DLL" (ByVal hbuf As Integer, ByRef stamp As TDS) As Integer
	
	'Lock / Unlock buffer in memory
	Declare Function olDmLockBuffer Lib "OLMEM32.DLL" (ByVal hbuf As Integer) As Integer
	Declare Function olDmUnlockBuffer Lib "OLMEM32.DLL" (ByVal hbuf As Integer) As Integer
	
	'Copy VB array to buffer
	Declare Function olDmCopyChannelToBuffer Lib "OLMEMSUP.DLL" (ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampInc As Short, ByRef vbarray As Short, ByRef Numsamples As Integer) As Integer
	Declare Function olDmCopyLongChannelToBuffer Lib "OLMEMSUP.DLL"  Alias "olDmCopyChannelToBuffer"(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampInc As Short, ByRef vbarray As Integer, ByRef Numsamples As Integer) As Integer
	Declare Function olDmCopySingleChannelToBuffer Lib "OLMEMSUP.DLL"  Alias "olDmCopyChannelToBuffer"(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampInc As Short, ByRef vbarray As Single, ByRef Numsamples As Integer) As Integer
	
	'Copy buffer data to VB array
	
	Declare Function olDmCopyChannelFromBuffer Lib "OLMEMSUP.DLL" (ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampInc As Short, ByRef vbarray As Short, ByRef Numsamples As Integer) As Integer
	Declare Function olDmCopyLongChannelFromBuffer Lib "OLMEMSUP.DLL"  Alias "olDmCopyChannelFromBuffer"(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampInc As Short, ByRef vbarray As Integer, ByRef Numsamples As Integer) As Integer
	Declare Function olDmCopySingleChannelFromBuffer Lib "OLMEMSUP.DLL"  Alias "olDmCopyChannelFromBuffer"(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampInc As Short, ByRef vbarray As Single, ByRef Numsamples As Integer) As Integer
	
	'OLMEM ERRORS
	Public Const OLNOERROR As Short = 0
	Public Const OLCANNOTALLOCBCB As Short = 200 '' Cannot allocate a buffer control block for the requested data buffer.
	Public Const OLCANNOTALLOCBUFFER As Short = 201 '' Cannot allocate the requested data buffer.
	Public Const OLBADBUFFERHANDLE As Short = 202 '' Invalid buffer handle (HBUF) passed to a library from an application.
	Public Const OLBUFFERLOCKFAILED As Short = 203 '' The data buffer cannot be put to a section because it cannot be properly locked.
	Public Const OLBUFFERLOCKED As Short = 204 '' Buffer locked
	Public Const OLBUFFERONLIST As Short = 205 '' Buffer on List
	Public Const OLCANNOTREALLOCBCB As Short = 206 '' Reallocation of a buffer control block was unsuccessful.
	Public Const OLCANNOTREALLOCBUFFER As Short = 207 '' Reallocation of the data buffer was unsuccessful.
	Public Const OLBADSAMPLESIZE As Short = 208 '' Bad Sample Size
	Public Const OLCANNOTALLOCLIST As Short = 209 '' Cannot Allocate List
	Public Const OLBADLISTHANDLE As Short = 210 '' Bad List Handle
	Public Const OLLISTNOTEMPTY As Short = 211 '' List not empty
	Public Const OLBUFFERNOTLOCKED As Short = 212 '' Buffer not locked
	Public Const OLBADDMACHANNEL As Short = 213 '' Invalid DMA channel specified
	Public Const OLDMACHANNELINUSE As Short = 214 '' Specified DMA channel in use
	Public Const OLBADIRQ As Short = 215 '' Invalid IRQ specified
	Public Const OLIRQINUSE As Short = 216 '' Specified IRQ in use
	Public Const OLNOSAMPLES As Short = 217 '' Buffer has no valid samples
	Public Const OLTOOMANYSAMPLES As Short = 218 '' Valid Samples cannot be larger than buffer
	Public Const OLBUFFERTOOSMALL As Short = 219 '' Specified buffer too small for requested copy operation
	
	' Global Memory Flags
	
	Public Const GMEM_FIXED As Short = 0
	Public Const GMEM_MOVEABLE As Short = 2
	Public Const GMEM_NOCOMPACT As Short = 16
	Public Const GMEM_NODISCARD As Short = 32
	Public Const GMEM_ZEROINIT As Short = 64
	Public Const GMEM_MODIFY As Short = 128
	Public Const GMEM_DISCARDABLE As Short = 256
	Public Const GMEM_NOT_BANKED As Short = 4096
	Public Const GMEM_SHARE As Short = 8192
	Public Const GMEM_DDESHARE As Short = 8192
	Public Const GMEM_NOTIFY As Short = 16384
	Sub CopyChannelFromBuffer(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampleInc As Short, ByRef vbarray As Short, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyChannelFromBuffer(hbuf, StartSamp, SampleInc, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyChannelFromBuffer", ecode)
		End If
		
		
	End Sub
	Sub CopyChannelToBuffer(ByRef hbuf As Integer, ByVal StartSamp As Short, ByVal SampleInc As Short, ByRef vbarray As Short, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyChannelToBuffer(hbuf, StartSamp, SampleInc, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyChannelToBuffer", ecode)
		End If
		
	End Sub
	Sub CopyLongChannelFromBuffer(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampleInc As Short, ByRef vbarray As Integer, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyLongChannelFromBuffer(hbuf, StartSamp, SampleInc, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyLongChannelFromBuffer", ecode)
		End If
		
		
	End Sub
	Sub CopyLongChannelToBuffer(ByRef hbuf As Integer, ByVal StartSamp As Short, ByVal SampleInc As Short, ByRef vbarray As Integer, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyLongChannelToBuffer(hbuf, StartSamp, SampleInc, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyLongChannelToBuffer", ecode)
		End If
		
	End Sub
	
	Sub CopyLongFromBuffer(ByVal hbuf As Integer, ByRef vbarray As Integer, ByVal maxsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyLongFromBuffer(hbuf, vbarray, maxsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyLongFromBuffer", ecode)
		End If
		
	End Sub
	Sub CopyLongToBuffer(ByVal hbuf As Integer, ByRef vbarray As Integer, ByVal Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyLongToBuffer(hbuf, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyLongToBuffer", ecode)
		End If
		
	End Sub
	Sub CopySingleChannelFromBuffer(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampleInc As Short, ByRef vbarray As Single, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopySingleChannelFromBuffer(hbuf, StartSamp, SampleInc, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopySingleChannelFromBuffer", ecode)
		End If
		
	End Sub
	
	Sub CopySingleChannelToBuffer(ByVal hbuf As Integer, ByVal StartSamp As Short, ByVal SampleInc As Short, ByRef vbarray As Single, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopySingleChannelToBuffer(hbuf, StartSamp, SampleInc, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopySingleChannelToBuffer", ecode)
		End If
		
	End Sub
	Sub CopySingleFromBuffer(ByVal hbuf As Integer, ByRef vbarray As Single, ByVal maxsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopySingleFromBuffer(hbuf, vbarray, maxsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopySingleFromBuffer", ecode)
		End If
	End Sub
	Sub CopySingleToBuffer(ByVal hbuf As Integer, ByRef vbarray As Single, ByVal Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopySingleToBuffer(hbuf, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopySingleToBuffer", ecode)
		End If
		
	End Sub
	
	
	
	
	
	
	
	
	Function AllocBuffer(ByVal WinFlags As Short, ByVal buffersize As Integer) As Integer
		Dim hbuf As Integer
		Dim ecode As Integer
		
		ecode = olDmAllocBuffer(WinFlags, buffersize, hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("AllocBuffer", ecode)
		Else
			AllocBuffer = hbuf
		End If
		
	End Function
	
	Sub OLMEMTrap(ByRef FuncName As String, ByVal Errnum As Short)
		Dim ErrStg As String = Space(255)
		Const MB_ICONSTOP As Short = 16
		
		'Get DT-Open Layers Error string
		ErrStg = GetOLMEMErrorString(Errnum, 80)
		'MsgBox ErrStg, MB_ICONSTOP, FuncName & " Error"
		Err.Raise(vbObjectError + Errnum, FuncName, ErrStg)
		
	End Sub
	Function GetOLMEMErrorString(ByVal ecode As Short, ByVal maxsize As Short) As String
		Dim ErrMsg As String = Space(255)
		
		If maxsize > 255 Then
			maxsize = 255
		End If
		
		olDmGetErrorString(ecode, ErrMsg, maxsize)
		
		GetOLMEMErrorString = ErrMsg
		
	End Function
	
	
	Function CallocBuffer(ByVal WinFlags As Short, ByVal ExFlags As Short, ByVal samples As Integer, ByVal Samplesize As Short) As Integer
		Dim hbuf As Integer
		Dim ecode As Integer
		
		ecode = olDmCallocBuffer(WinFlags, ExFlags, samples, Samplesize, hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CallocBuffer", ecode)
		Else
			CallocBuffer = hbuf
		End If
		
	End Function
	
	Sub CopyFromBuffer(ByVal hbuf As Integer, ByRef vbarray As Short, ByVal maxsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyFromBuffer(hbuf, vbarray, maxsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyFromBuffer", ecode)
		End If
		
	End Sub
	
	Sub CopyToBuffer(ByVal hbuf As Integer, ByRef vbarray As Short, ByVal Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDmCopyToBuffer(hbuf, vbarray, Numsamples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("CopyToBuffer", ecode)
		End If
		
	End Sub
	
	Sub FreeBuffer(ByVal hbuf As Integer)
		Dim ecode As Integer
		
		ecode = olDmFreeBuffer(hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("FreeBuffer", ecode)
		End If
		
	End Sub
	
	Function GetBufferSize(ByVal hbuf As Integer) As Integer
		Dim ecode As Integer
		Dim size As Integer
		
		ecode = olDmGetBufferSize(hbuf, size)
		If ecode <> OLNOERROR Then
			OLMEMTrap("GetBufferSize", ecode)
		Else
			GetBufferSize = size
		End If
		
	End Function
	Function GetDataBits(ByVal hbuf As Integer) As Short
		Dim ecode As Integer
		Dim size As Integer
		
		ecode = olDmGetDataBits(hbuf, size)
		If ecode <> OLNOERROR Then
			OLMEMTrap("GetDataBits", ecode)
		Else
			GetDataBits = size
		End If
		
	End Function
	Function GetDataWidth(ByVal hbuf As Integer) As Short
		Dim ecode As Integer
		Dim size As Integer
		
		ecode = olDmGetDataWidth(hbuf, size)
		If ecode <> OLNOERROR Then
			OLMEMTrap("GetDataWidth", ecode)
		Else
			GetDataWidth = size
		End If
		
	End Function
	Function GetMaxSamples(ByVal hbuf As Integer) As Integer
		Dim ecode As Integer
		Dim samples As Integer
		
		ecode = olDmGetMaxSamples(hbuf, samples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("GetMaxSamples", ecode)
		Else
			GetMaxSamples = samples
		End If
		
	End Function
	
	Function GetTimeDateStamp(ByVal hbuf As Integer) As String
		Dim ecode As Integer
		Dim TDSstamp As TDS
		Dim stampstg As String = Space(255)
		
		ecode = olDmGetTimeDateStamp(hbuf, TDSstamp)
		If ecode <> OLNOERROR Then
			OLMEMTrap("GetTimeDateStamp", ecode)
		Else
			stampstg = TDSstamp.DtMonth & "/" & TDSstamp.DtDay & "/" & TDSstamp.DtYear & "  " & TDSstamp.DtHour & ":" & TDSstamp.min & ":" & TDSstamp.sec
			GetTimeDateStamp = stampstg
		End If
		
	End Function
	
	Function GetValidSamples(ByVal hbuf As Integer) As Integer
		Dim ecode As Integer
		Dim samples As Integer
		
		ecode = olDmGetValidSamples(hbuf, samples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("GetValidSamples", ecode)
		Else
			GetValidSamples = samples
		End If
		
	End Function
	Sub LockBuffer(ByVal hbuf As Integer)
		Dim ecode As Integer
		
		ecode = olDmLockBuffer(hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("LockBuffer", ecode)
		End If
		
	End Sub
	
	Sub ReAllocBuffer(ByVal WinFlags As Short, ByVal buffersize As Integer, ByRef hbuf As Integer)
		Dim ecode As Integer
		
		ecode = olDmReAllocBuffer(WinFlags, buffersize, hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("ReAllocBuffer", ecode)
		End If
		
	End Sub
	
	Sub ReCallocBuffer(ByVal WinFlags As Short, ByVal ExFlags As Short, ByVal samples As Integer, ByVal Samplesize As Short, ByRef hbuf As Integer)
		Dim ecode As Integer
		
		ecode = olDmReCallocBuffer(WinFlags, ExFlags, samples, Samplesize, hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("ReCallocBuffer", ecode)
		End If
		
	End Sub
	
	Sub SetValidSamples(ByVal hbuf As Integer, ByVal samples As Integer)
		Dim ecode As Integer
		
		ecode = olDmSetValidSamples(hbuf, samples)
		If ecode <> OLNOERROR Then
			OLMEMTrap("SetValidSamples", ecode)
		End If
		
	End Sub
	
	Sub UnlockBuffer(ByVal hbuf As Integer)
		Dim ecode As Integer
		
		ecode = olDmUnlockBuffer(hbuf)
		If ecode <> OLNOERROR Then
			OLMEMTrap("UnLockBuffer", ecode)
		End If
		
	End Sub
End Module