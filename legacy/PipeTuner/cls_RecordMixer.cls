VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_RecordMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Mixer line to record from can be selected,
' its volume can be changed.

Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As GMEMFlags, _
    ByVal dwBytes As Long _
) As Long

Private Declare Function GlobalFree Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long _
) As Long

Private Declare Sub CpyMem Lib "kernel32" _
Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal cBytes As Long _
)

Private Declare Function waveInGetDevCaps Lib "winmm" _
Alias "waveInGetDevCapsA" ( _
    ByVal hwi As Long, _
    pwic As Any, _
    ByVal cbwic As Long _
) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function mixerClose Lib "winmm" ( _
    ByVal hmx As Long _
) As Long

Private Declare Function mixerOpen Lib "winmm" ( _
    phmx As Long, _
    ByVal uMxId As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long _
) As Long

Private Declare Function mixerGetControlDetails Lib "winmm" _
Alias "mixerGetControlDetailsA" ( _
    ByVal hmxobj As Long, _
    pmxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long _
) As Long

Private Declare Function mixerGetLineControls Lib "winmm" _
Alias "mixerGetLineControlsA" ( _
    ByVal hmxobj As Long, _
    pmxlc As MIXERLINECONTROLS, _
    ByVal fdwControls As Long _
) As Long

Private Declare Function mixerGetLineInfo Lib "winmm" _
Alias "mixerGetLineInfoA" ( _
    ByVal hmxobj As Long, _
    pmxl As MIXERLINE, _
    ByVal fdwInfo As Long _
) As Long

Private Declare Function mixerSetControlDetails Lib "winmm" ( _
    ByVal hmxobj As Long, _
    pmxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long _
) As Long

Private Declare Function mixerGetDevCaps Lib "winmm" _
Alias "mixerGetDevCapsA" ( _
    ByVal uMxId As Long, _
    pmxcaps As MIXERCAPS, _
    ByVal cbmxcaps As Long _
) As Long

' of course there are a lot more line ids
Public Enum MIXER_RECORDING_LINES
    MIXERLINE_ANALOG = &H100A&
    MIXERLINE_AUXILIARY = &H1009&
    MIXERLINE_COMPACTDISC = &H1005&
    MIXERLINE_DIGITAL = &H1001&
    MIXERLINE_LINE = &H1002&
    MIXERLINE_MICROPHONE = &H1003&
    MIXERLINE_PCSPEAKER = &H1007&
    MIXERLINE_SYNTHESIZER = &H1004&
    MIXERLINE_TELEPHONE = &H1006&
    MIXERLINE_UNDEFINED = &H1000&
    MIXERLINE_WAVEOUT = &H1008&
End Enum

Private Enum GMEMFlags
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
    GMEM_ZEROINIT = &H40
End Enum

Private Const MAXPNAMELEN                               As Long = 32&

Private Const MIXER_GETCONTROLDETAILSF_LISTTEXT         As Long = &H1&
Private Const MIXER_GETCONTROLDETAILSF_VALUE            As Long = &H0&

Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE          As Long = &H2&
    
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE          As Long = &H3&
Private Const MIXER_GETLINEINFOF_LINEID                 As Long = &H2&
Private Const MIXER_GETLINEINFOF_SOURCE                 As Long = &H1&
Private Const MIXER_GETLINEINFOF_DESTINATION            As Long = &H0&

Private Const MIXER_LONG_NAME_CHARS                     As Long = 64
Private Const MIXER_SHORT_NAME_CHARS                    As Long = 16

Private Const MIXERCONTROL_CONTROLTYPE_VOLUME           As Long = &H50030001
Private Const MIXERCONTROL_CONTROLTYPE_MUTE             As Long = &H20010002

Private Const MIXER_SETCONTROLDETAILSF_VALUE            As Long = &H0&

Private Const MIXER_OBJECTF_WAVEIN                      As Long = &H20000000

Private Const MIXERCONTROL_CT_UNITS_BOOLEAN             As Long = &H10000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE          As Long = &H1000000
Private Const MIXERCONTROL_CT_SC_LIST_SINGLE            As Long = 0&
Private Const MIXERCONTROL_CT_CLASS_LIST                As Long = &H70000000

Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT   As Long = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT     As Long = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_MIXER            As Long = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MUX              As Long = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)

Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN        As Long = &H7&



Private Const MMSYSERR_NOERROR                          As Long = 0&



Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1                                    As Long
    dwParam2                                    As Long
    szName                                      As Long
End Type

Private Type MIXERCONTROL
    cbStruct                                    As Long
    dwControlID                                 As Long
    dwControlType                               As Long
    fdwControl                                  As Long
    cMultipleItems                              As Long
    szShortName(MIXER_SHORT_NAME_CHARS / 2 - 1) As Integer
    szName(MIXER_LONG_NAME_CHARS / 2 - 1)       As Integer
    Bounds(5)                                   As Long
    Metrics(5)                                  As Long
End Type

Private Type MIXERCONTROLDETAILS
    cbStruct                                    As Long
    dwControlID                                 As Long
    cChannels                                   As Long
    item                                        As Long
    cbDetails                                   As Long
    paDetails                                   As Long
End Type

Private Type Target
    dwType                                      As Long
    dwDeviceID                                  As Long
    wMid                                        As Integer
    wPid                                        As Integer
    vDriverVersion                              As Long
    szPname                                     As String * MAXPNAMELEN
End Type

Private Type MIXERLINE
    cbStruct                                    As Long
    dwDestination                               As Long
    dwSource                                    As Long
    dwLineID                                    As Long
    fdwLine                                     As Long
    dwUser                                      As Long
    dwComponentType                             As Long
    cChannels                                   As Long
    cConnections                                As Long
    cControls                                   As Long
    szShortName                                 As String * MIXER_SHORT_NAME_CHARS
    szName                                      As String * MIXER_LONG_NAME_CHARS
    tTarget                                     As Target
End Type

Private Type MIXERLINECONTROLS
    cbStruct                                    As Long
    dwLineID                                    As Long
    dwControl                                   As Long
    cControls                                   As Long
    cbmxctrl                                    As Long
    pamxctrl                                    As Long
End Type

Private Type MIXERCAPS
    wMid                                        As Integer
    wPid                                        As Integer
    vDriverVersion                              As Long
    szPname                                     As String * 32
    fdwSupport                                  As Long
    cDestinations                               As Long
End Type

Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue                                     As Long
End Type

Private Type volume_stereo
    L                                           As Long
    R                                           As Long
End Type

Private Type WAVEINCAPS
    wMid                                        As Integer
    wPid                                        As Integer
    vDriverVer                                  As Long
    szPname                                     As String * MAXPNAMELEN
    dwFormats                                   As Long
    wChannels                                   As Integer
    wReserved1                                  As Integer
End Type

Private lngCurDev                               As Long

Private hMixer                                  As Long

' http://www.ureader.com/message/1359669.aspx
Public Function SetMixerLine2( _
    ByVal devid As Long, _
    ByVal Index As Long _
) As Boolean

    Dim retval      As Long
    Dim terr        As Long
    Dim mxl         As MIXERLINE
    Dim controls    As MIXERLINECONTROLS
    Dim control     As MIXERCONTROL
    Dim hControl    As Long
    Dim pControl    As Long
    Dim cd          As MIXERCONTROLDETAILS
    Dim hr          As Long
    Dim i           As Long
    Dim C           As Long
    Dim j           As Long
    Dim lv          As MIXERCONTROLDETAILS_BOOLEAN
    Dim hLV         As Long
    Dim pLV         As Long

    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_DESTINATION)

    If (hr <> 0) Then Exit Function

    For i = 0 To mxl.dwDestination - 1
        controls.cbStruct = Len(controls)
        controls.dwLineID = mxl.dwLineID
        controls.cControls = mxl.cControls
        controls.cbmxctrl = Len(control)

        hControl = GlobalAlloc(GMEM_ZEROINIT, Len(control) * controls.cControls)
        pControl = GlobalLock(hControl)

        controls.pamxctrl = pControl
        controls.dwControl = MIXERCONTROL_CONTROLTYPE_MUX

        terr = mixerGetLineControls(devid, controls, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If terr <> 0 Then Exit Function
        CpyMem control, ByVal pControl, Len(control)
        If Index > control.cMultipleItems Then
            GoTo SkipItem
        End If

        For C = 0 To controls.cControls - 1
            CpyMem control, ByVal pControl + Len(control) * C, Len(control)
            If (MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_MUX And control.dwControlType)) Then
                cd.cbStruct = Len(cd)
                cd.dwControlID = control.dwControlID
                cd.cChannels = 1
                cd.item = control.cMultipleItems
                cd.cbDetails = Len(lv)

                hLV = GlobalAlloc(GMEM_ZEROINIT, cd.cChannels * cd.item * cd.cbDetails)
                pLV = GlobalLock(hLV)
                
                cd.paDetails = pLV

                terr = mixerGetControlDetails(devid, cd, MIXER_GETCONTROLDETAILSF_VALUE)

                For j = 0 To cd.item - 1
                    CpyMem lv, ByVal pLV + Len(lv) * j, Len(lv)
                    If lv.fValue Then retval = i
                    lv.fValue = Abs(CBool(j = Index))
                    CpyMem ByVal pLV + Len(lv) * j, lv, Len(lv)
                Next

                terr = mixerSetControlDetails(devid, cd, MIXER_SETCONTROLDETAILSF_VALUE)

                GlobalUnlock hLV
                GlobalFree hLV
            End If
        Next

SkipItem:
        GlobalUnlock hControl
        GlobalFree hControl
    Next

    SetMixerLine2 = True
End Function

' http://www.ureader.com/message/1359669.aspx
Public Function GetMixerLine2( _
    ByVal devid As Long _
) As Long

    Dim retval      As Long
    Dim terr        As Long
    Dim mxl         As MIXERLINE
    Dim controls    As MIXERLINECONTROLS
    Dim control     As MIXERCONTROL
    Dim hControl    As Long
    Dim pControl    As Long
    Dim cd          As MIXERCONTROLDETAILS
    Dim hr          As Long
    Dim i           As Long
    Dim C           As Long
    Dim j           As Long
    Dim lv          As MIXERCONTROLDETAILS_BOOLEAN
    Dim hLV         As Long
    Dim pLV         As Long

    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_WAVEIN
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    hr = mixerGetLineInfo(devid, mxl, MIXER_GETLINEINFOF_DESTINATION)

    If (hr <> 0) Then Exit Function

    For i = 0 To mxl.dwDestination - 1
        controls.cbStruct = Len(controls)
        controls.dwLineID = mxl.dwLineID
        controls.cControls = mxl.cControls
        controls.cbmxctrl = Len(control)

        hControl = GlobalAlloc(GMEM_ZEROINIT, Len(control) * controls.cControls)
        pControl = GlobalLock(hControl)

        controls.pamxctrl = pControl
        controls.dwControl = MIXERCONTROL_CONTROLTYPE_MUX

        terr = mixerGetLineControls(devid, controls, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If terr <> 0 Then Exit Function

        For C = 0 To controls.cControls - 1
            CpyMem control, ByVal pControl + Len(control) * C, Len(control)
            If (MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_MUX And control.dwControlType)) Then
                cd.cbStruct = Len(cd)
                cd.dwControlID = control.dwControlID
                cd.cChannels = 1
                cd.item = control.cMultipleItems
                cd.cbDetails = Len(lv)

                hLV = GlobalAlloc(GMEM_ZEROINIT, cd.cChannels * cd.item * cd.cbDetails)
                pLV = GlobalLock(hLV)
                
                cd.paDetails = pLV

                terr = mixerGetControlDetails(devid, cd, MIXER_GETCONTROLDETAILSF_VALUE)

                For j = 0 To cd.item - 1
                    CpyMem lv, ByVal pLV + Len(lv) * j, Len(lv)
                    If lv.fValue Then retval = j
                Next

                GlobalUnlock hLV
                GlobalFree hLV
            End If
        Next

SkipItem:
        GlobalUnlock hControl
        GlobalFree hControl
    Next

    GetMixerLine2 = retval
End Function

Public Property Get DeviceCount() As Long

    DeviceCount = waveInGetNumDevs()
    
End Property

Public Property Get DeviceName(ByVal Index As Long) As String

    Dim udtInfo As WAVEINCAPS
    Dim strName As String

    waveInGetDevCaps Index, udtInfo, Len(udtInfo)
    strName = udtInfo.szPname

    If InStr(strName, Chr$(0)) > 0 Then
        strName = Left$(strName, InStr(strName, Chr$(0)) - 1)
    End If

    DeviceName = udtInfo.szPname
End Property

Public Function SelectDevice(ByVal Index As Long) As Boolean

    If Index < 0 Or Index > DeviceCount - 1 Then
        Exit Function
    End If

    lngCurDev = -1

    If hMixer <> 0 Then
        mixerClose hMixer
        hMixer = 0
    End If

    mixerOpen hMixer, Index, 0, 0, MIXER_OBJECTF_WAVEIN
    If hMixer = 0 Then Exit Function

    lngCurDev = Index

    SelectDevice = True
End Function

Public Property Get MixerLineCount() As Long

    Dim i   As Long

    If lngCurDev < 0 Then Exit Property

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerLineCount = MixerDestinationConnections(hMixer, i)
            Exit For
        End If
    Next
End Property

Public Property Get MixerLineName(ByVal Index As Long) As String

    Dim i   As Long

    If lngCurDev < 0 Then Exit Property

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerLineName = MixerConnectionName(hMixer, i, Index)
            Exit For
        End If
    Next
End Property

Public Function SelectMixerLine(ByVal Index As Long) As Boolean

    If lngCurDev < 0 Then Exit Function

    SelectMixerLine = SetMixerLine2(hMixer, MixerLineCount - Index - 1)
   
End Function

Public Property Get SelectedDevice() As Long

    SelectedDevice = lngCurDev
   
End Property

Public Property Get SelectedMixerLine() As Long

     SelectedMixerLine = MixerLineCount - GetMixerLine2(hMixer) - 1
     
End Property

Public Property Get MixerLineType(Optional ByVal Index As Long = -1) As Long

    Dim i       As Long
    Dim lngLine As Long

    If Index = -1 Then
        lngLine = SelectedMixerLine
    Else
        lngLine = Index
    End If

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerLineType = MixerConnectionType(hMixer, i, lngLine)
            Exit For
        End If
    Next
End Property

Public Property Get MixerLineVolume() As Long

    Dim i       As Long
    Dim udtVol  As volume_stereo

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            udtVol = MixerConnectionVolume(hMixer, i, SelectedMixerLine)
            Exit For
        End If
    Next

    MixerLineVolume = udtVol.L
End Property

Public Property Let MixerLineVolume(ByVal Volume As Long)

    Dim i       As Long
    Dim udtVol  As volume_stereo

    udtVol.L = Volume
    udtVol.R = Volume

    For i = 0 To MixerDestinations(hMixer) - 1
        If MixerDestinationType(hMixer, i) = MIXERLINE_COMPONENTTYPE_DST_WAVEIN Then
            MixerSetConnectionVolume hMixer, i, SelectedMixerLine, udtVol
            Exit For
        End If
    Next
End Property

' /////////////////////////////////////////////////
' ////////// WINDOWS MIXER FUNCTIONS
' /////////////////////////////////////////////////

Private Property Get MixerConnectionType( _
    ByVal devid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long _
) As Long

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo devid, udtML, MIXER_GETLINEINFOF_SOURCE

    MixerConnectionType = udtML.dwComponentType
End Property

Private Property Get MixerDestinations(ByVal deviceid As Long) As Long

    Dim udtCaps As MIXERCAPS

    mixerGetDevCaps deviceid, udtCaps, Len(udtCaps)

    MixerDestinations = udtCaps.cDestinations
End Property

Private Property Get MixerConnectionName( _
    ByVal deviceid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long _
) As String

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo deviceid, _
                     udtML, _
                     MIXER_GETLINEINFOF_SOURCE

    MixerConnectionName = udtML.szName
End Property

Private Property Get MixerDestinationConnections( _
    ByVal deviceid As Long, _
    ByVal Destination As Long _
) As Long

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination

    mixerGetLineInfo deviceid, _
                     udtML, _
                     MIXER_GETLINEINFOF_DESTINATION

    MixerDestinationConnections = udtML.cConnections
End Property

Private Property Get MixerDestinationType( _
    ByVal deviceid As Long, _
    ByVal Destination As Long _
) As Long

    Dim udtML   As MIXERLINE

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination

    mixerGetLineInfo deviceid, _
                     udtML, _
                     MIXER_GETLINEINFOF_DESTINATION

    MixerDestinationType = udtML.dwComponentType
End Property

Private Sub MixerSetConnectionVolume( _
    ByVal deviceid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long, _
    vol As volume_stereo _
)

    Dim udtML   As MIXERLINE
    Dim udtMCL  As MIXERCONTROLDETAILS
    Dim CtrlID  As Long

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo deviceid, udtML, MIXER_GETLINEINFOF_SOURCE

    CtrlID = MixerGetControlID(deviceid, _
                               udtML.dwComponentType, _
                               MIXERCONTROL_CONTROLTYPE_VOLUME, _
                               udtML.dwLineID)

    With udtMCL
        .cbDetails = 4
        .cbStruct = Len(udtMCL)
        .cChannels = udtML.cChannels
        .dwControlID = CtrlID
        .item = 0
        .paDetails = VarPtr(vol)
    End With

    mixerSetControlDetails deviceid, _
                           udtMCL, _
                           MIXER_SETCONTROLDETAILSF_VALUE
End Sub

Private Property Get MixerConnectionVolume( _
    ByVal deviceid As Long, _
    ByVal Destination As Long, _
    ByVal connection As Long _
) As volume_stereo

    Dim udtML   As MIXERLINE
    Dim udtMCL  As MIXERCONTROLDETAILS
    Dim udtVol  As volume_stereo
    Dim CtrlID  As Long

    udtML.cbStruct = Len(udtML)
    udtML.dwDestination = Destination
    udtML.dwSource = connection

    mixerGetLineInfo deviceid, udtML, MIXER_GETLINEINFOF_SOURCE

    CtrlID = MixerGetControlID(deviceid, _
                               udtML.dwComponentType, _
                               MIXERCONTROL_CONTROLTYPE_VOLUME, _
                               udtML.dwLineID)

    With udtMCL
        .cbDetails = 4
        .cbStruct = Len(udtMCL)
        .cChannels = udtML.cChannels
        .dwControlID = CtrlID
        .item = 0
        .paDetails = VarPtr(udtVol)
    End With

    mixerGetControlDetails deviceid, _
                           udtMCL, _
                           MIXER_GETCONTROLDETAILSF_VALUE

    MixerConnectionVolume = udtVol
End Property

Private Function MixerGetControlID( _
    ByVal deviceid As Long, _
    ByVal ComponentType As Long, _
    ByVal ControlType As Long, _
    ByVal LineID As Long _
) As Long

    Dim hMem     As Long
    Dim MC       As MIXERCONTROL
    Dim MxrLine  As MIXERLINE
    Dim MLC      As MIXERLINECONTROLS

    MxrLine.cbStruct = Len(MxrLine)
    MxrLine.dwComponentType = ComponentType

    If mixerGetLineInfo(deviceid, _
                        MxrLine, _
                        MIXER_GETLINEINFOF_COMPONENTTYPE) = 0 Then

        MLC.cbStruct = Len(MLC)
        MLC.dwLineID = LineID
        MLC.dwControl = ControlType
        MLC.cControls = 1
        MLC.cbmxctrl = Len(MC)

        hMem = GlobalAlloc(&H40, Len(MC))
        If hMem = 0 Then Exit Function
        MLC.pamxctrl = GlobalLock(hMem)

        MC.cbStruct = Len(MC)

        If mixerGetLineControls(deviceid, _
                                MLC, _
                                MIXER_GETLINECONTROLSF_ONEBYTYPE) = 0 Then

            CpyMem MC, ByVal MLC.pamxctrl, Len(MC)
            MixerGetControlID = MC.dwControlID
        End If

        GlobalUnlock hMem
        GlobalFree hMem
    End If
End Function
