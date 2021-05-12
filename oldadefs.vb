Option Strict Off
Option Explicit On
Module oldadefs
	
	'------------ ol Data Acq Function prototypes -----------'
	'Get corresponding DT-Open Layers error string
	Declare Sub olDaGetErrorString Lib "OLDAAPI32.DLL" (ByVal ecode As Integer, ByVal ErrMsg As String, ByVal maxsize As Short)
	
	'Fetch subsystem info
	Structure DassInfo
		Dim hDev As Integer 'handle to board   'same as dtol.hDev
		Dim SSType As Short 'same as dtol.SSType
		Dim SSElement As Short 'same as dtol.SSElement
	End Structure

	Declare Function olDaGetDASSInfo Lib "OLDAAPI32.DLL" (ByVal hDass As Integer, ByRef Info As DassInfo) As Integer
	
	
	'Fetch version info
	Declare Function olDaGetVersion Lib "OLDAAPI32.DLL" (ByVal Version As String, ByVal maxsize As Short) As Integer
	Declare Function olDaGetDriverVersion Lib "OLDAAPI32.DLL" (ByVal hDev As Integer, ByVal Version As String, ByVal maxsize As Short) As Integer
	
	'Set Range for specified channel
	Declare Function olDaSetChannelRange Lib "OLDAAPI32.DLL" (ByVal hDass As Integer, ByVal chan As Short, ByVal maxvoltage As Double, ByVal minvoltage As Double) As Integer
	
	'Get Range for specified channel
	Declare Function olDaGetChannelRange Lib "OLDAAPI32.DLL" (ByVal hDass As Integer, ByVal chan As Short, ByRef maxvoltage As Double, ByRef minvoltage As Double) As Integer
	
	'Simultaneous Starting of multiple subsystems
	'Create a Simultaneous Start List
	Declare Function olDaGetSSList Lib "OLDAAPI32.DLL" (ByRef hSSList As Integer) As Integer
	'Destroy Simultaneous Start List
	Declare Function olDaReleaseSSList Lib "OLDAAPI32.DLL" (ByVal hSSList As Integer) As Integer
	'Put subsystem onto Simultaneous Start List
	Declare Function olDaPutDassToSSList Lib "OLDAAPI32.DLL" (ByVal hSSList As Integer, ByVal hDass As Integer) As Integer
	'Prestart all subsystems on Simultaneous Start List
	Declare Function olDaSimultaneousPrestart Lib "OLDAAPI32.DLL" (ByVal hSSList As Integer) As Integer
	'Start all subsystems on Simultaneous Start List
	Declare Function olDaSimultaneousStart Lib "OLDAAPI32.DLL" (ByVal hSSList As Integer) As Integer
	
	'Copy data from inprocess buffer to specified buffer & put onto Done Queue
	Declare Function olDaFlushFromBufferInprocess Lib "OLDAAPI32.DLL" (ByVal hDass As Integer, ByVal hbuf As Integer, ByVal Numsamples As Integer) As Integer
	
	'------------------------------ system error codes-------------------------
	Public Const DTACQ32_ERRORBASE As Short = 22010
	Public Const OLNOERROR As Short = 0
	Public Const OLBADCAP As Short = 1
	Public Const OLBADELEMENT As Short = 2
	Public Const OLBADSUBSYSTEM As Short = 3
	Public Const OLNOTENOUGHDMACHANS As Short = 4
	Public Const OLBADLISTSIZE As Short = 5
	Public Const OLBADLISTENTRY As Short = 6
	Public Const OLBADCHANNEL As Short = 7
	Public Const OLBADCHANNELTYPE As Short = 8
	Public Const OLBADENCODING As Short = 9
	Public Const OLBADTRIGGER As Short = 10
	Public Const OLBADRESOLUTION As Short = 11
	Public Const OLBADCLOCKSOURCE As Short = 12
	Public Const OLBADFREQUENCY As Short = 13
	Public Const OLBADPULSETYPE As Short = 14
	Public Const OLBADPULSEWIDTH As Short = 15
	Public Const OLBADCOUNTERMODE As Short = 16
	Public Const OLBADCASCADEMODE As Short = 17
	Public Const OLBADDATAFLOW As Short = 18
	Public Const OLBADWINDOWHANDLE As Short = 19
	Public Const OLSUBSYSINUSE As Short = 20
	Public Const OLSUBSYSNOTINUSE As Short = 21
	Public Const OLALREADYRUNNING As Short = 22
	Public Const OLNOCHANNELLIST As Short = 23
	Public Const OLNOGAINLIST As Short = 24
	Public Const OLNOFILTERLIST As Short = 25
	Public Const OLNOTCONFIGURED As Short = 26
	Public Const OLDATAFLOWMISMATCH As Short = 27
	Public Const OLNOTRUNNING As Short = 28
	Public Const OLBADRANGE As Short = 29
	Public Const OLBADSSCAP As Short = 30
	Public Const OLBADDEVCAP As Short = 31
	Public Const OLBADRANGEINDEX As Short = 32
	Public Const OLBADFILTERINDEX As Short = 33
	Public Const OLBADGAININDEX As Short = 34
	Public Const OLBADWRAPMODE As Short = 35
	Public Const OLNOTSUPPORTED As Short = 36
	Public Const OLBADDIVIDER As Short = 37
	Public Const OLBADGATE As Short = 38
	
	Public Const OLBADDEVHANDLE As Short = 39
	Public Const OLBADSSHANDLE As Short = 40
	Public Const OLCANNOTALLOCDASS As Short = 41
	Public Const OLCANNOTDEALLOCDASS As Short = 42
	Public Const OLBUFFERSLEFT As Short = 43
	
	Public Const OLBOARDRUNNING As Short = 44 ' another subsystem on board is already running
	Public Const OLINVALIDCHANNELLIST As Short = 45 ' channel list has been filled incorrectly
	Public Const OLINVALIDCLKTRIGCOMBO As Short = 46 ' selected clock & trigger source may not be used together
	Public Const OLCANNOTALLOCLUSERDATA As Short = 47 ' driver could not allocate needed memory
	
	Public Const OLCANTSVTRIGSCAN As Short = 48
	Public Const OLCANTSVEXTCLOCK As Short = 49
	Public Const OLBADRESOLUTIONINDEX As Short = 50
	
	Public Const OLADTRIGERR As Short = 60
	Public Const OLADOVRRN As Short = 61
	Public Const OLDATRIGERR As Short = 62
	Public Const OLDAUNDRRN As Short = 63
	
	Public Const OLNOREADYBUFFERS As Short = 64
	
	Public Const OLBADCPU As Short = 65
	Public Const OLBADWINMODE As Short = 66
	Public Const OLCANNOTOPENDRIVER As Short = 67
	Public Const OLBADENUMCAP As Short = 68
	Public Const OLBADDASSPROC As Short = 69
	Public Const OLBADENUMPROC As Short = 70
	Public Const OLNOWINDOWHANDLE As Short = 71
	Public Const OLCANTCASCADE As Short = 72
	
	Public Const OLINVALIDCONFIGURATION As Short = 73
	Public Const OLCANNOTALLOCJUMPERS As Short = 74
	Public Const OLCANNOTALLOCCHANNELLIST As Short = 75
	Public Const OLCANNOTALLOCGAINLIST As Short = 76
	Public Const OLCANNOTALLOCFILTERLIST As Short = 77
	Public Const OLNOBOARDSINSTALLED As Short = 78
	Public Const OLINVALIDDMASCANCOMBO As Short = 79
	Public Const OLINVALIDPULSETYPE As Short = 80
	Public Const OLINVALIDGAINLIST As Short = 81
	Public Const OLWRONGCOUNTERMODE As Short = 82
	Public Const OLLPSTRNULL As Short = 83
	Public Const OLINVALIDPIODFCOMBO As Short = 84 ' Invalid Polled I/O combination
	Public Const OLINVALIDSCANTRIGCOMBO As Short = 85 ' Invalid Scan / Trigger combo
	Public Const OLBADGAIN As Short = 86 ' Invalid Gain
	Public Const OLNOMEASURENOTIFY As Short = 87 ' No window handle specified for frequency measurement
	Public Const OLBADCOUNTDURATION As Short = 88 ' Invalid count duration for frequency measurement
	Public Const OLBADQUEUE As Short = 89 ' Invalid queue type specified
	Public Const OLBADRETRIGGERRATE As Short = 90 ' Invalid Retrigger Rate for channel list size
	Public Const OLCOMMANDTIMEOUT As Short = 91 ' No Command Response from Hardware
	Public Const OLCOMMANDOVERRUN As Short = 92 ' Hardware Command Sequence Error
	Public Const OLDATAOVERRUN As Short = 93 ' Hardware Data Sequence Error
	
	Public Const OLCANNOTALLOCTIMERDATA As Short = 94 ' Cannot allocate timer data for driver
	Public Const OLBADHTIMER As Short = 95 ' Invalid Timer handle
	Public Const OLBADTIMERMODE As Short = 96 ' Invalid Timer mode
	Public Const OLBADTIMERFREQUENCY As Short = 97 ' Invalid Timer frequency
	Public Const OLBADTIMERPROC As Short = 98 ' Invalid Timer callback procedure
	Public Const OLBADDMABUFFERSIZE As Short = 99 ' Invalid Timer DMA buffer size
	
	Public Const OLBADDIGITALIOLISTVALUE As Short = 100 ' Illegal synchronous digital I/O value requested.
	Public Const OLCANNOTALLOCSSLIST As Short = 101 ' Cannot allocate simultaneous start list
	Public Const OLBADSSLISTHANDLE As Short = 102 ' Illegal simultaneous start list handle specified.
	Public Const OLBADSSHANDLEONLIST As Short = 103 ' Invalid subsystem handle on simultaneous start list.
	Public Const OLNOSSHANDLEONLIST As Short = 104 ' No subsystem handles on simultaneous start list.
	Public Const OLNOCHANNELINHIBITLIST As Short = 105 ' The subsystem has no channel inhibit list.
	Public Const OLNODIGITALIOLIST As Short = 106 ' The subsystem has no digital I/O list.
	Public Const OLNOTPRESTARTED As Short = 107 ' The subsystem has not been prestarted.
	Public Const OLBADNOTIFYPROC As Short = 108 ' Invalid notification procedure
	Public Const OLBADTRANSFERCOUNT As Short = 109 ' Invalid DT-Connect transfer count
	Public Const OLBADTRANSFERSIZE As Short = 110 ' Invalid DT-Connect transfer size
	Public Const OLCANNOTALLOCINHIBITLIST As Short = 111 ' Cannot allocate channel inhibit list
	Public Const OLCANNOTALLOCDIGITALIOLIST As Short = 112 ' Cannot allocate digital I/O list
	Public Const OLINVALIDINHIBITLIST As Short = 113 ' Channel inhibit list has been filled incorrectly
	Public Const OLSSHANDLEALREADYONLIST As Short = 114 ' Subsystem is already on simultaneous start list
	Public Const OLCANNOTALLOCRANGELIST As Short = 115 'Cannot allocate range list
	Public Const OLNORANGELIST As Short = 116 'No Range List
	Public Const OLNOBUFFERINPROCESS As Short = 117 'No buffers in process
	Public Const OLREQUIREDSUBSYSINUSE As Short = 118 'Additional required subsystem in use
	
	Public Const VBMAXERROR As Short = 32767
	'------------------------------- Device Capabilities -----------------------------'
	Public Const OLDC_ADELEMENTS As Short = 0
	Public Const OLDC_DAELEMENTS As Short = 1
	Public Const OLDC_DINELEMENTS As Short = 2
	Public Const OLDC_DOUTELEMENTS As Short = 3
	Public Const OLDC_SRLELEMENTS As Short = 4
	Public Const OLDC_CTELEMENTS As Short = 5
	Public Const OLDC_TOTALELEMENTS As Boolean = Not 0
	
	
	'------------------------------- SubSystem Capabilities -----------------------------'
	Public Const OLSSC_MAXSECHANS As Short = 0
	Public Const OLSSC_MAXDICHANS As Short = 1
	Public Const OLSSC_CGLDEPTH As Short = 2
	Public Const OLSSC_NUMFILTERS As Short = 3
	Public Const OLSSC_NUMGAINS As Short = 4
	Public Const OLSSC_NUMRANGES As Short = 5
	Public Const OLSSC_NUMDMACHANS As Short = 6
	Public Const OLSSC_NUMCHANNELS As Short = 7
	Public Const OLSSC_NUMEXTRACLOCKS As Short = 8
	Public Const OLSSC_NUMEXTRATRIGGERS As Short = 9
	Public Const OLSSC_NUMRESOLUTIONS As Short = 10
	Public Const OLSSC_SUP_INTERRUPT As Short = 11
	Public Const OLSSC_SUP_SINGLEENDED As Short = 12
	Public Const OLSSC_SUP_DIFFERENTIAL As Short = 13
	Public Const OLSSC_SUP_BINARY As Short = 14
	Public Const OLSSC_SUP_2SCOMP As Short = 15
	Public Const OLSSC_SUP_SOFTTRIG As Short = 16
	Public Const OLSSC_SUP_EXTERNTRIG As Short = 17
	Public Const OLSSC_SUP_THRESHTRIGPOS As Short = 18
	Public Const OLSSC_SUP_THRESHTRIGNEG As Short = 19
	Public Const OLSSC_SUP_ANALOGEVENTTRIG As Short = 20
	Public Const OLSSC_SUP_DIGITALEVENTTRIG As Short = 21
	Public Const OLSSC_SUP_TIMEREVENTTRIG As Short = 22
	Public Const OLSSC_SUP_TRIGSCAN As Short = 23
	Public Const OLSSC_SUP_INTCLOCK As Short = 24
	Public Const OLSSC_SUP_EXTCLOCK As Short = 25
	Public Const OLSSC_SUP_SWCAL As Short = 26
	Public Const OLSSC_SUP_EXP2896 As Short = 27
	Public Const OLSSC_SUP_EXP727 As Short = 28
	Public Const OLSSC_SUP_FILTERPERCHAN As Short = 29
	Public Const OLSSC_SUP_DTCONNECT As Short = 30
	Public Const OLSSC_SUP_FIFO As Short = 31
	Public Const OLSSC_SUP_PROGRAMGAIN As Short = 32
	Public Const OLSSC_SUP_PROCESSOR As Short = 33
	Public Const OLSSC_SUP_SWRESOLUTION As Short = 34
	Public Const OLSSC_SUP_CONTINUOUS As Short = 35
	Public Const OLSSC_SUP_SINGLEVALUE As Short = 36
	Public Const OLSSC_SUP_PAUSE As Short = 37
	Public Const OLSSC_SUP_WRPMULTIPLE As Short = 38
	Public Const OLSSC_SUP_WRPSINGLE As Short = 39
	Public Const OLSSC_SUP_POSTMESSAGE As Short = 40
	Public Const OLSSC_SUP_CASCADING As Short = 41
	Public Const OLSSC_SUP_CTMODE_COUNT As Short = 42
	Public Const OLSSC_SUP_CTMODE_RATE As Short = 43
	Public Const OLSSC_SUP_CTMODE_ONESHOT As Short = 44
	Public Const OLSSC_SUP_CTMODE_ONESHOT_RPT As Short = 45
	
	Public Const OLSSC_MAX_DIGITALIOLIST_VALUE As Short = 46
	Public Const OLSSC_SUP_DTCONNECT_CONTINUOUS As Short = 47
	Public Const OLSSC_SUP_DTCONNECT_BURST As Short = 48
	Public Const OLSSC_SUP_CHANNELLIST_INHIBIT As Short = 49
	Public Const OLSSC_SUP_SYNCHRONOUS_DIGITALIO As Short = 50
	Public Const OLSSC_SUP_SIMULTANEOUS_START As Short = 51
	Public Const OLSSC_SUP_INPROCESSFLUSH As Short = 52
	
	Public Const OLSSC_SUP_RANGEPERCHANNEL As Short = 53
	Public Const OLSSC_SUP_SIMULTANEOUS_SH As Short = 54
	Public Const OLSSC_SUP_RANDOM_CGL As Short = 55
	Public Const OLSSC_SUP_SEQUENTIAL_CGL As Short = 56
	Public Const OLSSC_SUP_ZEROSEQUENTIAL_CGL As Short = 57
	
	Public Const OLSSC_SUP_GAPFREE_NODMA As Short = 58
	Public Const OLSSC_SUP_GAPFREE_SINGLEDMA As Short = 59
	Public Const OLSSC_SUP_GAPFREE_DUALDMA As Short = 60
	
	Public Const OLSSCE_MAXTHROUGHPUT As Short = 61
	Public Const OLSSCE_MINTHROUGHPUT As Short = 62
	Public Const OLSSCE_MAXRETRIGGER As Short = 63
	Public Const OLSSCE_MINRETRIGGER As Short = 64
	Public Const OLSSCE_MAXCLOCKDIVIDER As Short = 65
	Public Const OLSSCE_MINCLOCKDIVIDER As Short = 66
	Public Const OLSSCE_BASECLOCK As Short = 67
	Public Const OLSSCE_RANGELOW As Short = 68
	Public Const OLSSCE_RANGEHIGH As Short = 69
	Public Const OLSSCE_FILTER As Short = 70
	Public Const OLSSCE_GAIN As Short = 71
	Public Const OLSSCE_RESOLUTION As Short = 72
	
	Public Const OLSSC_SUP_PLS_HIGH2LOW As Short = 73
	Public Const OLSSC_SUP_PLS_LOW2HIGH As Short = 74
	
	Public Const OLSSC_SUP_GATE_NONE As Short = 75
	Public Const OLSSC_SUP_GATE_HIGH_LEVEL As Short = 76
	Public Const OLSSC_SUP_GATE_LOW_LEVEL As Short = 77
	Public Const OLSSC_SUP_GATE_HIGH_EDGE As Short = 78
	Public Const OLSSC_SUP_GATE_LOW_EDGE As Short = 79
	Public Const OLSSC_SUP_GATE_LEVEL As Short = 80
	Public Const OLSSC_SUP_GATE_HIGH_LEVEL_DEBOUNCE As Short = 81
	Public Const OLSSC_SUP_GATE_LOW_LEVEL_DEBOUNCE As Short = 82
	Public Const OLSSC_SUP_GATE_HIGH_EDGE_DEBOUNCE As Short = 83
	Public Const OLSSC_SUP_GATE_LOW_EDGE_DEBOUNCE As Short = 84
	Public Const OLSSC_SUP_GATE_LEVEL_DEBOUNCE As Short = 85
	
	Public Const OLSS_SUP_RETRIGGER_INTERNAL As Short = 86
	Public Const OLSS_SUP_RETRIGGER_SCAN_PER_TRIGGER As Short = 87
	Public Const OLSSC_MAXMULTISCAN As Short = 88
	Public Const OLSSC_SUP_CONTINUOUS_PRETRIG As Short = 89
	Public Const OLSSC_SUP_CONTINUOUS_ABOUTTRIG As Short = 90
	Public Const OLSSC_SUP_BUFFERING As Short = 91
	Public Const OLSSC_SUP_RETRIGGER_EXTRA As Short = 92
	
	'------------------------------- Configuration Settings -----------------------------'
	' for SubSysType property
	Public Const OLSS_AD As Short = 0
	Public Const OLSS_DA As Short = 1
	Public Const OLSS_DIN As Short = 2
	Public Const OLSS_DOUT As Short = 3
	Public Const OLSS_SRL As Short = 4
	Public Const OLSS_CT As Short = 5
	
	' for ChannelType property
	Public Const OL_CHNT_SINGLEENDED As Short = 0
	Public Const OL_CHNT_DIFFERENTIAL As Short = 1
	
	' for Encoding property
	Public Const OL_ENC_BINARY As Short = 0
	Public Const OL_ENC_2SCOMP As Short = 1
	
	' for trigger property
	Public Const OL_TRG_SOFT As Short = 0
	Public Const OL_TRG_EXTERN As Short = 1
	Public Const OL_TRG_THRESHPOS As Short = 2
	Public Const OL_TRG_THRESHNEG As Short = 3
	Public Const OL_TRG_ANALOGEVENT As Short = 4
	Public Const OL_TRG_DIGITALEVENT As Short = 5
	Public Const OL_TRG_TIMEREVENT As Short = 6
	Public Const OL_TRG_EXTRA As Short = 7
	
	' for ClockSource property
	Public Const OL_CLK_INTERNAL As Short = 0
	Public Const OL_CLK_EXTERNAL As Short = 1
	Public Const OL_CLK_EXTRA As Short = 2
	
	' for GateType property
	Public Const OL_GATE_NONE As Short = 0
	Public Const OL_GATE_HIGH_LEVEL As Short = 1
	Public Const OL_GATE_LOW_LEVEL As Short = 2
	Public Const OL_GATE_HIGH_EDGE As Short = 3
	Public Const OL_GATE_LOW_EDGE As Short = 4
	Public Const OL_GATE_LEVEL As Short = 5
	Public Const OL_GATE_HIGH_LEVEL_DEBOUNCE As Short = 6
	Public Const OL_GATE_LOW_LEVEL_DEBOUNCE As Short = 7
	Public Const OL_GATE_HIGH_EDGE_DEBOUNCE As Short = 8
	Public Const OL_GATE_LOW_EDGE_DEBOUNCE As Short = 9
	Public Const OL_GATE_LEVEL_DEBOUNCE As Short = 10
	
	' for PulseType property
	Public Const OL_PLS_HIGH2LOW As Short = 0
	Public Const OL_PLS_LOW2HIGH As Short = 1
	
	' for CTMode property
	Public Const OL_CTMODE_COUNT As Short = 0
	Public Const OL_CTMODE_RATE As Short = 1
	Public Const OL_CTMODE_ONESHOT As Short = 2
	Public Const OL_CTMODE_ONESHOT_RPT As Short = 3
	Public Const OL_CTMODE_UP_DOWN As Short = 4
	Public Const OL_CTMODE_MEASURE As Short = 5
	Public Const OL_CTMODE_CONT_MEASURE As Short = 6
	
	' for MeasureEdge property
	Public Const OL_GATE_RISING As Short = 0
	Public Const OL_GATE_FALLING As Short = 1
	Public Const OL_CLOCK_RISING As Short = 2
	Public Const OL_CLOCK_FALLING As Short = 3
	
	' for DataFlow property
	Public Const OL_DF_CONTINUOUS As Short = 0
	Public Const OL_DF_SINGLEVALUE As Short = 1
	Public Const OL_DF_DTCONNECT_CONTINUOUS As Short = 2
	Public Const OL_DF_DTCONNECT_BURST As Short = 3
	Public Const OL_DF_CONTINUOUS_PRETRIG As Short = 4
	Public Const OL_DF_CONTINUOUS_ABOUTTRIG As Short = 5
	
	' for CascadeMode Property
	Public Const OL_CT_CASCADE As Short = 0
	Public Const OL_CT_SINGLE As Short = 1
	
	' for WrapMode Property
	Public Const OL_WRP_NONE As Short = 0
	Public Const OL_WRP_MULTIPLE As Short = 1
	Public Const OL_WRP_SINGLE As Short = 2
	
	' for QueueSize Property
	Public Const OL_QUE_READY As Short = 0
	Public Const OL_QUE_DONE As Short = 1
	Public Const OL_QUE_INPROCESS As Short = 2
	
	
	' for RetriggerMode Property
	Public Const OL_RETRIGGER_INTERNAL As Short = 0
	Public Const OL_RETRIGGER_SCAN_PER_TRIGGER As Short = 1
	Public Const OL_RETRIGGER_EXTRA As Short = 2
	
	
	
	
	
	
	
	
	
	Sub DTOLTrap(ByRef FuncName As String, ByVal Errnum As Short)
		Dim ErrStg As String = Space(255)
		Const MB_ICONSTOP As Short = 16
		
		'Get DT-Open Layers Error string
		ErrStg = GetErrorString(Errnum, 80)
		'MsgBox ErrStg, MB_ICONSTOP, FuncName & " Error"
		Err.Raise(vbObjectError + Errnum, FuncName, ErrStg)
		
	End Sub
	
	Sub FlushFromBufferInProcess(ByRef subsystem As AxDTAcq32Lib.AxDTAcq32, ByVal hbuf As Integer, ByRef Numsamples As Integer)
		Dim ecode As Integer
		
		ecode = olDaFlushFromBufferInprocess(subsystem.hDass, hbuf, Numsamples)
		If ecode <> OLNOERROR Then
			DTOLTrap("FlushFromBufferInProcess", ecode)
		End If
		
	End Sub
	
	
	
	Function GetChannelRange(ByRef subsystem As AxDTAcq32Lib.AxDTAcq32, ByVal chan As Short) As Short
		Dim ecode As Integer
		Dim MaxRange As Double
		Dim MinRange As Double
		Dim Range As Short
		
		ecode = olDaGetChannelRange(subsystem.hDass, chan, MaxRange, MinRange)
		If ecode <> OLNOERROR Then
			DTOLTrap("GetChannelRange", ecode)
		Else
			For Range = 0 To subsystem.numRanges - 1
				If (MaxRange = subsystem.get_MaxRangeValues(Range)) And (MinRange = subsystem.get_MinRangeValues(Range)) Then
					GetChannelRange = Range
				End If
			Next Range
		End If
		
		
	End Function
	
	
	Function GetDriverVersion(ByRef subsystem As AxDTAcq32Lib.AxDTAcq32, ByVal maxsize As Short) As String
		Dim ecode As Integer
		Dim ver As String = Space(255)
		
		If maxsize > 255 Then
			maxsize = 255
		End If
		
		ecode = olDaGetDriverVersion(subsystem.hDev, ver, maxsize)
		If ecode <> OLNOERROR Then
			DTOLTrap("GetDriverVersion", ecode)
		Else
			GetDriverVersion = ver
		End If
		
	End Function
	
	Function GetErrorString(ByVal ecode As Short, ByVal maxsize As Short) As String
		Dim ErrMsg As String = Space(255)
		
		If maxsize > 255 Then
			maxsize = 255
		End If
		
		olDaGetErrorString(ecode, ErrMsg, maxsize)
		
		GetErrorString = ErrMsg
		
	End Function
	
	Function GetSimultaneousStartList() As Integer
		Dim ecode As Integer
		Dim hSSList As Integer
		
		ecode = olDaGetSSList(hSSList)
		If ecode <> OLNOERROR Then
			DTOLTrap("GetSimultaneousStartList", ecode)
		Else
			GetSimultaneousStartList = hSSList
		End If
		
		
	End Function
	
	
	
	
	
	
	Sub PutSubSysOnSSList(ByVal hSSList As Integer, ByRef subsystem As AxDTAcq32Lib.AxDTAcq32)
		Dim ecode As Integer

		ecode = olDaPutDassToSSList(hSSList, subsystem.hDass)
		If ecode <> OLNOERROR Then
			DTOLTrap("PutSubSysOnSSList", ecode)
		End If

	End Sub

	Sub ReleaseSimultaneousStartList(ByVal hSSList As Integer)
		Dim ecode As Integer

		ecode = olDaReleaseSSList(hSSList)
		If ecode <> OLNOERROR Then
			DTOLTrap("ReleaseSimultaneousStartList", ecode)
		End If

	End Sub

	Sub SetChannelRange(ByRef subsystem As AxDTAcq32Lib.AxDTAcq32, ByVal chan As Short, ByVal Range As Short)
		Dim ecode As Integer

		ecode = olDaSetChannelRange(subsystem.hDass, chan, subsystem.get_MaxRangeValues(Range), subsystem.get_MinRangeValues(Range))
		If ecode <> OLNOERROR Then
			DTOLTrap("SetChannelRange", ecode)
		End If

	End Sub

	Sub SimultaneousPrestart(ByVal hSSList As Integer)
		Dim ecode As Integer

		ecode = olDaSimultaneousPrestart(hSSList)
		If ecode <> OLNOERROR Then
			DTOLTrap("SimultaneousPrestart", ecode)
		End If

	End Sub

	Sub SimultaneousStart(ByVal hSSList As Integer)
		Dim ecode As Integer

		ecode = olDaSimultaneousStart(hSSList)
		If ecode <> OLNOERROR Then
			DTOLTrap("SimultaneousStart", ecode)
		End If


	End Sub


	'This function converts a value from the specified subsystem into a voltage.
	'This function assumes a gain of 1.
	'
	Function ValueToVolts(ByRef subsystem As AxDTAcq32Lib.AxDTAcq32, ByVal Value As Integer) As Single
		Dim Res As Integer
		Dim min, Max As Single
		Dim Volts As Single

		Res = 2 ^ subsystem.Resolution
		min = subsystem.MinRange
		Max = subsystem.MaxRange

		'make sure value is sign extended if 2's comp
		If subsystem.Encoding = OL_ENC_2SCOMP Then
			Value = Value And (Res - 1)
			If Value >= Res / 2 Then Value = Value - Res
		End If

		'convert it to volts
		Volts = Value * (Max - min) / Res

		'adjust DC offset
		If subsystem.Encoding = OL_ENC_2SCOMP Then 'if 2's comp
			Volts = Volts + (Max + min) / 2
		Else 'if offset binary
			Volts = Volts + min
		End If

		ValueToVolts = Volts
	End Function


	'This function converts the specified voltage into units that
	'are appropriate to the specified subsystem.
	'This function assumes a gain of 1.
	'
	Function VoltsToValue(ByRef subsystem As AxDTAcq32Lib.AxDTAcq32, ByVal Volts As Single) As Integer
		Dim Res As Integer
		Dim min, Max As Single
		Dim Value As Integer

		Res = 2 ^ subsystem.Resolution
		min = subsystem.MinRange
		Max = subsystem.MaxRange

		'clip input to range
		If Volts > Max Then Volts = Max
		If Volts < min Then Volts = min

		If subsystem.Encoding = OL_ENC_2SCOMP Then 'if 2's comp encoding
			'convert to 2's comp value
			Value = (Volts - (min + Max) / 2) * Res / (Max - min)
			'adjust for binary wrap if any
			If Value = Res / 2 Then Value = Value - 1
		Else
			'convert to offset binary
			Value = (Volts - min) * Res / (Max - min)
			'adjust for binary wrap if any
			If Value = Res Then Value = Value - 1
		End If

		VoltsToValue = Value

	End Function
End Module