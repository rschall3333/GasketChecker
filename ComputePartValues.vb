Option Strict Off
Option Explicit On 
Module ComputePartValues
    Public Sub leftConeCompute(ByVal keyenceHead01Output02 As Double)
        rawLeftConeDbl = keyenceHead01Output02
        ' Sin 30 degrees = 0.5
        leftConeDbl = (zeroLeftConeDbl - keyenceHead01Output02) * 0.5
    End Sub
    Public Sub rightConeCompute(ByVal keyenceHead01Output03 As Double)
        rawRightConeDbl = keyenceHead01Output03
        ' Sin 30 degrees = 0.5
        rightConeDbl = (zeroRightConeDbl - keyenceHead01Output03) * 0.5
    End Sub
    Public Sub PadCompute(ByVal keyenceHead03Output04 As Double)
        rawPadDbl = keyenceHead03Output04
        padDbl = zeroPadDbl - keyenceHead03Output04
    End Sub
    Public Sub FlangeCompute(ByVal keyenceLKOutput01 As Double)
        rawFlangeDbl = keyenceLKOutput01
        ' Divide by 1000 to convert from mils
        flangeDbl = (keyenceLKOutput01 / 1000) - (zeroFlangeDbl / 1000)
    End Sub
End Module