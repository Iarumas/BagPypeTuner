Attribute VB_Name = "dB_Meter"
Option Explicit

Private mobj_Level_Meter As LevelMeter
Private msng_dB_Value As Single
Private msng_dB_Limit As Single
Private msng_Max_Value As Single
Private msng_UpperLimit As Single
Private mbln_Level_Good As Boolean

Public Function LevelMeterObject() As LevelMeter
    Set LevelMeterObject = mobj_Level_Meter
End Function
Public Sub SetLevelMeterObject(ByRef Value As LevelMeter)
    Set mobj_Level_Meter = Value
End Sub
Public Sub Set_dB_Limit(ByVal Value As Single)
    msng_dB_Limit = Value
    mobj_Level_Meter.MaxLevel = -1 * msng_dB_Limit
    mobj_Level_Meter.HiLevel = 0.9 * mobj_Level_Meter.MaxLevel
    mobj_Level_Meter.MidLevel = 0.5 * mobj_Level_Meter.MaxLevel
End Sub
Public Function dB_Limit() As Single
    dB_Limit = msng_dB_Limit
End Function
Public Sub Set_UpperLimit(ByVal Value As Single)
    msng_UpperLimit = Value
End Sub
Public Function UpperLimit() As Single
    UpperLimit = msng_UpperLimit
End Function
Public Function Level_Good() As Boolean
    If msng_dB_Value < 0 And msng_dB_Value > msng_dB_Limit Then
        mbln_Level_Good = True
    Else
        mbln_Level_Good = False
    End If
    Level_Good = mbln_Level_Good
End Function
Public Function dB_Value() As Single
    dB_Value = msng_dB_Value
End Function
Private Function Value_To_dB(ByVal Value) As Single
    ' converts values to dB
    msng_Max_Value = Value
    If msng_Max_Value > 0 Then
        msng_dB_Value = CSng(10 * (Log(msng_Max_Value / msng_UpperLimit) / Log(10)))   ' dB
    Else
        msng_dB_Value = msng_dB_Limit
    End If
    
    Value_To_dB = msng_dB_Value
End Function

Private Function dB_To_Value(ByVal Value As Single)
    'converts dB (-x to 0) into linear values
    
    msng_dB_Value = Value
    If msng_dB_Value < 0 Then
        msng_Max_Value = msng_UpperLimit * (10 ^ (msng_dB_Value / 10))
    Else
        msng_Max_Value = msng_UpperLimit
    End If
     dB_To_Value = msng_Max_Value
End Function

Public Sub Update_dB_Level(Value As Single)
    ' update  dB meter with msg_db_Value
    
    msng_dB_Value = Value
    
    ' only if new value in visible area (e.g. max level = 15 -> new value > -15 dB)
    ' level meter: 0 to max level (e.g. 0 to 15)
    ' dB values are -x to 0
    If msng_dB_Value > -1 * mobj_Level_Meter.MaxLevel Then
        mobj_Level_Meter.Level = msng_dB_Value + mobj_Level_Meter.MaxLevel
    Else
        mobj_Level_Meter.Level = 0
    End If
    
End Sub



