Option Strict Off
Option Explicit On 

Module globalVariables

    'Booleans
    Public partDataChanged As Boolean   ' Siganls user has changed part data
    Public AutoModeBool As Boolean      ' 0 - Manual, 1 - Auto
    Public ErrorStateBool As Boolean    ' 0 - No error 1 - Error

    ' Strings
    Public PartNumberStr As String      ' User entered part number
    Public OperatorIdStr As String      ' User entered operator ID
    Public MoldNumberStr As String      ' User entered mold number
    Public CavityNumberStr As String    ' User entered cavity number
    Public HeatNumberStr As String      ' User entered heat number
    Public PartDateStr As String        ' User entered part date
    Public RunNumberStr As String       ' User entered run number
    Public ShopOrderStr As String       ' User entered shop order
    Public PartTimeStr As String        ' User entered part time

    ' Double precision floats
    Public leftConeDbl As Double        ' Computed left cone value
    Public rightConeDbl As Double       ' Computed right cone value
    Public padDbl As Double             ' Computed pad value
    Public flangeDbl As Double          ' Converted flange value

    Public zeroLeftConeDbl As Double    ' Zero value for the left cone
    Public zeroRightConeDbl As Double   ' Zero value for the right cone
    Public zeroPadDbl As Double         ' Zero value for the pad
    Public zeroFlangeDbl As Double      ' Zero value for the flange

    Public rawLeftConeDbl As Double     ' raw value for left cone
    Public rawRightConeDbl As Double    ' raw value for right cone
    Public rawPadDbl As Double          ' raw value for pad
    Public rawFlangeDbl As Double       ' raw value for flange

    Public padDataArray(4) As Double       ' Report Data for Pad
    Public flangeDataArray(4) As Double      ' Report Data for Flange
    Public leftConeDataArray(4) As Double    ' Report Data for Left Cone
    Public rightConeDataArray(4) As Double   ' Report Data for Right Cone

    Public WithEvents keyenceLK5001 As Rs232   'Keyence LS-5041 Head01 and Head02
    Public WithEvents keyenceLKG3001 As Rs232  'Keyence LK-G87 Head01
    Public keyenceHead01Output02 As Double     'Data from LS-5041 Head01
    Public keyenceHead01Output03 As Double     'Data from LS-5041 Head01
    Public keyenceHead02Output04 As Double     'Data from LS-5041 Head01
    Public keyenceLKG3001Value As Double        'Data from LS-5041 Head02

    Public boolCycleComplete As Boolean            'Bit signaling cycle complete
    Public boolCycleRunning As Boolean             'Bit signaling cycle running
    Public boolVacuumFault As Boolean              'Bit signaling vacuum fault

    Public port02Head01Out02Err As Boolean         ' 0-okay , 1-port not available
    Public port02Head01Out03Err As Boolean         ' 0-okay , 1-port not available
    Public port02Head02Out04Err As Boolean         ' 0-okay , 1-port not available
    Public port05Head01Out01Err As Boolean         ' 0-okay , 1-port not available

    Public laserDisplacementERROR As Boolean       'Laser displacement alarm

    Public AutoPartCount As Integer = 0             ' Auto-Mode part counter

    Public CheckZeroPageUp As Boolean = False

    Public ValidationCollection As Boolean = False   ' validation data collection
    Public ProductionCollection As Boolean = True    ' production data collection

    Public writeStreamOneShot As Boolean

    Public disableManualButton As Boolean

    Public DataAcquisitionOneShot As Boolean               'one shot for data aquisition

    Public SuspendDOUT As Boolean

    Public AutoPartCountPrevious As Integer
    Public AutoPartCountRunning As Integer

    Public StreamingDataNote As String
    Public ProductionReportNote As String
    Public StreamDataNoteBool As Boolean
    Public ProductionReportBool As Boolean
End Module