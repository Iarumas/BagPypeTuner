VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWinMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function waveOutGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function mixerGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function midiInGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function midiOutGetNumDevs Lib "winmm" ( _
) As Long

Private Declare Function waveOutGetDevCaps Lib "winmm" _
Alias "waveOutGetDevCapsA" ( _
    ByVal uDeviceID As Long, _
    lpCaps As WAVEOUTCAPS, _
    ByVal uSize As Long _
) As Long

Private Declare Function waveInGetDevCaps Lib "winmm" _
Alias "waveInGetDevCapsA" ( _
    ByVal uDeviceID As Long, _
    pwic As WAVEINCAPS, _
    ByVal cbwic As Long _
) As Long

Private Declare Function mixerGetDevCaps Lib "winmm" _
Alias "mixerGetDevCapsA" ( _
    ByVal uDeviceID As Long, _
    pmxcaps As MIXERCAPS, _
    ByVal cbmxcaps As Long _
) As Long

Private Declare Function midiOutGetDevCaps Lib "winmm" _
Alias "midiOutGetDevCapsA" ( _
    ByVal uDeviceID As Long, _
    lpMidiOutCaps As MIDIOUTCAPS, _
    ByVal cbMidiOutCaps As Long _
) As Long

Private Declare Function midiInGetDevCaps Lib "winmm" _
Alias "midiInGetDevCapsA" ( _
    ByVal uDeviceID As Long, _
    lpMidiInCaps As MIDIINCAPS, _
    ByVal cbMidiInCaps As Long _
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


Private Const MAX_PNAME_LEN                             As Long = 32&

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

Private Const MIXER_OBJECTF_MIDIIN                      As Long = &H40000000
Private Const MIXER_OBJECTF_MIDIOUT                     As Long = &H30000000
Private Const MIXER_OBJECTF_WAVEIN                      As Long = &H20000000
Private Const MIXER_OBJECTF_WAVEOUT                     As Long = &H10000000
Private Const MIXER_OBJECTF_MIXER                       As Long = &H0

Private Const MIXERCONTROL_CT_UNITS_BOOLEAN             As Long = &H10000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE          As Long = &H1000000
Private Const MIXERCONTROL_CT_SC_LIST_SINGLE            As Long = 0&
Private Const MIXERCONTROL_CT_CLASS_LIST                As Long = &H70000000

Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT   As Long = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT     As Long = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_SINGLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_MIXER            As Long = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MUX              As Long = (MIXERCONTROL_CONTROLTYPE_SINGLESELECT + 1)


Public Enum MIXER_SOURCES
    MIXERLINE_SRC_ANALOG = &H100A&
    MIXERLINE_SRC_AUXILIARY = &H1009&
    MIXERLINE_SRC_COMPACTDISC = &H1005&
    MIXERLINE_SRC_DIGITAL = &H1001&
    MIXERLINE_SRC_LINE = &H1002&
    MIXERLINE_SRC_MICROPHONE = &H1003&
    MIXERLINE_SRC_PCSPEAKER = &H1007&
    MIXERLINE_SRC_SYNTHESIZER = &H1004&
    MIXERLINE_SRC_TELEPHONE = &H1006&
    MIXERLINE_SRC_UNDEFINED = &H1000&
    MIXERLINE_SRC_WAVEOUT = &H1008&
End Enum

Public Enum MIXER_DESTINATIONS
    MIXERLINE_DST_FIRST = &H0&
    MIXERLINE_DST_DIGITAL = &H1&
    MIXERLINE_DST_HEADPHONES = &H5&
    MIXERLINE_DST_LAST = &H8&
    MIXERLINE_DST_LINE = &H2&
    MIXERLINE_DST_MONITOR = &H3&
    MIXERLINE_DST_SPEAKERS = &H4&
    MIXERLINE_DST_TELEPHONE = &H6&
    MIXERLINE_DST_UNDEFINED = &H0&
    MIXERLINE_DST_VOICEIN = &H8&
    MIXERLINE_DST_WAVEIN = &H7&
End Enum

Public Enum MIXER_OPENBY
    MIXER_OPENBY_WAVEOUT_ID = MIXER_OBJECTF_WAVEOUT
    MIXER_OPENBY_WAVEIN_ID = MIXER_OBJECTF_WAVEIN
    MIXER_OPENBY_MIDIOUT_ID = MIXER_OBJECTF_MIDIOUT
    MIXER_OPENBY_MIDIIN_ID = MIXER_OBJECTF_MIDIIN
    MIXER_OPENBY_MIXER_ID = MIXER_OBJECTF_MIXER
End Enum

Private Enum GMEMFlags
    GMEM_FIXED = &H0
    GMEM_MOVEABLE = &H2
    GMEM_ZEROINIT = &H40
End Enum


Private Type WAVEOUTCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname             As String * MAX_PNAME_LEN
    dwFormats           As Long
    wChannels           As Integer
    dwSupport           As Long
End Type

Private Type WAVEINCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVer          As Long
    szPname             As String * MAX_PNAME_LEN
    dwFormats           As Long
    wChannels           As Integer
    wReserved1          As Integer
End Type

Private Type MIXERCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname             As String * MAX_PNAME_LEN
    fdwSupport          As Long
    cDestinations       As Long
End Type

Private Type MIDIINCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname             As String * MAX_PNAME_LEN
    fdwSupport          As Long
End Type

Private Type MIDIOUTCAPS
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname             As String * MAX_PNAME_LEN
    wTechnology         As Integer
    wVoices             As Integer
    wNotes              As Integer
    wChannelMask        As Integer
    dwSupport           As Long
End Type

Private Type TARGET
    dwType              As Long
    dwDeviceID          As Long
    wMid                As Integer
    wPid                As Integer
    vDriverVersion      As Long
    szPname             As String * MAX_PNAME_LEN
End Type

Private Type MIXERLINE
    cbStruct            As Long
    dwDestination       As Long
    dwSource            As Long
    dwLineID            As Long
    fdwLine             As Long
    dwUser              As Long
    dwComponentType     As Long
    cChannels           As Long
    cConnections        As Long
    cControls           As Long
    szShortName         As String * MIXER_SHORT_NAME_CHARS
    szName              As String * MIXER_LONG_NAME_CHARS
    tTarget             As TARGET
End Type

Private Type MIXERLINECONTROLS
    cbStruct            As Long
    dwLineID            As Long
    dwControl           As Long
    cControls           As Long
    cbmxctrl            As Long
    pamxctrl            As Long
End Type

Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1            As Long
    dwParam2            As Long
    szName              As Long
End Type

Private Type MIXERCONTROL
    cbStruct            As Long
    dwControlID         As Long
    dwControlType       As Long
    fdwControl          As Long
    cMultipleItems      As Long
    szShortName(15)     As Byte
    szName(63)          As Byte
    lMinimum            As Long
    lMaximum            As Long
    reserved(10)        As Long
End Type

Private Type MIXERCONTROLDETAILS
    cbStruct            As Long
    dwControlID         As Long
    cChannels           As Long
    item                As Long
    cbDetails           As Long
    paDetails           As Long
End Type

Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue             As Long
End Type

Private Type MIXERCONTROLDETAILS_UNSIGNED
     fValue             As Long
End Type

Private hMixer          As Long

Public Property Get DeviceCount( _
    ByVal devmode As MIXER_OPENBY _
) As Long

    Select Case devmode
        Case MIXER_OPENBY_MIXER_ID:     DeviceCount = mixerGetNumDevs()
        Case MIXER_OPENBY_WAVEIN_ID:    DeviceCount = waveInGetNumDevs()
        Case MIXER_OPENBY_WAVEOUT_ID:   DeviceCount = waveOutGetNumDevs()
        Case MIXER_OPENBY_MIDIIN_ID:    DeviceCount = midiInGetNumDevs()
        Case MIXER_OPENBY_MIDIOUT_ID:   DeviceCount = midiOutGetNumDevs()
    End Select
End Property

Public Property Get DeviceName( _
    ByVal index As Long, _
    ByVal devmode As MIXER_OPENBY _
) As String

    Dim udtDevWO    As WAVEOUTCAPS
    Dim udtDevWI    As WAVEINCAPS
    Dim udtDevMX    As MIXERCAPS
    Dim udtDevMI    As MIDIINCAPS
    Dim udtDevMO    As MIDIOUTCAPS

    Select Case devmode

        Case MIXER_OPENBY_MIXER_ID:
            waveOutGetDevCaps index, udtDevWO, Len(udtDevWO)
            DeviceName = udtDevWO.szPname

        Case MIXER_OPENBY_WAVEIN_ID:
            waveInGetDevCaps index, udtDevWI, Len(udtDevWI)
            DeviceName = udtDevWI.szPname

        Case MIXER_OPENBY_WAVEOUT_ID:
            mixerGetDevCaps index, udtDevMX, Len(udtDevMX)
            DeviceName = udtDevMX.szPname

        Case MIXER_OPENBY_MIDIIN_ID:
            midiInGetDevCaps index, udtDevMI, Len(udtDevMI)
            DeviceName = udtDevMI.szPname

        Case MIXER_OPENBY_MIDIOUT_ID:
            midiOutGetDevCaps index, udtDevMO, Len(udtDevMO)
            DeviceName = udtDevMO.szPname

    End Select
End Property

Public Function DeviceOpen( _
    ByVal devid As Long, _
    ByVal openas As MIXER_OPENBY _
) As Boolean

    DeviceClose

    If Not MM_IS_OK(mixerOpen(hMixer, devid, 0, 0, openas)) Then
        hMixer = 0
    End If

    DeviceOpen = hMixer <> 0
End Function

Public Sub DeviceClose()
    If hMixer <> 0 Then
        mixerClose hMixer
        hMixer = 0
    End If
End Sub

Public Property Get DestinationCount( _
) As Long

    Dim udtMxInfo   As MIXERCAPS

    If hMixer = 0 Then
        DestinationCount = -1
        Exit Property
    End If

    If Not MM_IS_OK(mixerGetDevCaps(hMixer, _
                                    udtMxInfo, _
                                    Len(udtMxInfo))) Then

        DestinationCount = -1
    Else
        DestinationCount = udtMxInfo.cDestinations
    End If
End Property

Public Property Get SourceCount( _
    ByVal destination As Long _
) As Long

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        SourceCount = -1
        Exit Property
    End If

    ' Mixer shall retrieve information about the
    ' destination by the destination's index
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        SourceCount = -1
    Else
        SourceCount = udtMXL.cConnections
    End If
End Property

Public Property Get DestinationName( _
    ByVal index As Long _
) As String

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        DestinationName = vbNullString
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        DestinationName = vbNullString
    Else
        DestinationName = TrimEx(udtMXL.szName)
    End If
End Property

Public Property Get SourceName( _
    ByVal destination As Long, _
    ByVal index As Long _
) As String

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        SourceName = vbNullString
        Exit Property
    End If

    ' Mixer shall retrieve information about the source
    ' by its destination and index
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        SourceName = vbNullString
    Else
        SourceName = TrimEx(udtMXL.szName)
    End If
End Property

Public Property Get DestinationChannels( _
    ByVal destination As Long _
) As Long

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        DestinationChannels = -1
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        DestinationChannels = -1
    Else
        DestinationChannels = udtMXL.cChannels
    End If
End Property

Public Property Get SourceChannels( _
    ByVal destination As Long, _
    ByVal index As Long _
) As Long

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        SourceChannels = -1
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        SourceChannels = -1
    Else
        SourceChannels = udtMXL.cChannels
    End If
End Property

Public Property Get DestinationMute( _
    ByVal index As Long _
) As Boolean

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtMute     As MIXERCONTROLDETAILS_BOOLEAN

    If hMixer = 0 Then
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        DestinationMute = -1
        Exit Property
    End If

    ' get the Control Id of mute setting for destination
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_MUTE
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        DestinationMute = -1
        Exit Property
    End If

    ' get the mute setting
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = udtMXL.cChannels
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtMute)
    udtMXCD.paDetails = VarPtr(udtMute)

    If Not MM_IS_OK(mixerGetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        DestinationMute = -1
        Exit Property
    End If

    DestinationMute = CBool(udtMute.fValue)
End Property

Public Property Let DestinationMute( _
    ByVal index As Long, _
    ByVal blnVal As Boolean _
)

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtMute     As MIXERCONTROLDETAILS_BOOLEAN

    If hMixer = 0 Then
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        Exit Property
    End If

    ' get the Control Id of mute setting for destination
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_MUTE
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        Exit Property
    End If

    ' set new mute setting
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = 1
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtMute)
    udtMXCD.paDetails = VarPtr(udtMute)

    udtMute.fValue = Abs(blnVal)

    If Not MM_IS_OK(mixerSetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        Exit Property
    End If
End Property

Public Property Get SourceMute( _
    ByVal destination As Long, _
    ByVal index As Long _
) As Boolean

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtMute     As MIXERCONTROLDETAILS_BOOLEAN

    If hMixer = 0 Then
        Exit Property
    End If

    ' information about the source
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        SourceMute = -1
        Exit Property
    End If

    ' Control Id of mute setting for source
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_MUTE
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        SourceMute = -1
        Exit Property
    End If

    ' get mute setting
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = udtMXL.cChannels
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtMute)
    udtMXCD.paDetails = VarPtr(udtMute)

    If Not MM_IS_OK(mixerGetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        SourceMute = -1
        Exit Property
    End If

    SourceMute = CBool(udtMute.fValue)
End Property

Public Property Let SourceMute( _
    ByVal destination As Long, _
    ByVal index As Long, _
    ByVal blnVal As Boolean _
)

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtMute     As MIXERCONTROLDETAILS_BOOLEAN

    If hMixer = 0 Then
        Exit Property
    End If

    ' information about source
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        Exit Property
    End If

    ' get the Control Id of the mute setting for the source
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_MUTE
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        Exit Property
    End If

    ' set new mute setting
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = 1
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtMute)
    udtMXCD.paDetails = VarPtr(udtMute)

    udtMute.fValue = Abs(blnVal)

    If Not MM_IS_OK(mixerSetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        Exit Property
    End If
End Property

Public Property Get DestinationVolume( _
    ByVal index As Long, _
    Optional ByVal channel As Long = -1 _
) As Long

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtVol()    As MIXERCONTROLDETAILS_UNSIGNED
    Dim i           As Long
    Dim lngVolMax   As Long
    Dim lngPer      As Long

    If hMixer = 0 Then
        DestinationVolume = -1
        Exit Property
    End If

    ' information about destination
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        DestinationVolume = -1
        Exit Property
    End If

    ' Control Id of volume setting for destination
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        DestinationVolume = -1
        Exit Property
    End If

    ' space for the volume of every channel
    ReDim udtVol(udtMXL.cChannels - 1) As MIXERCONTROLDETAILS_UNSIGNED

    ' get the volume for every channel of the destination
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = udtMXL.cChannels
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtVol(0))
    udtMXCD.paDetails = VarPtr(udtVol(0))

    If Not MM_IS_OK(mixerGetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        DestinationVolume = -1
        Exit Property
    End If

    ' volume in percent
    If channel = -1 Then
        ' all channels, return average
        For i = 0 To udtMXL.cChannels - 1
            lngPer = (udtVol(i).fValue / (udtMC.lMaximum - udtMC.lMinimum) * 100)

            If lngPer > lngVolMax Then
                lngVolMax = lngPer
            End If
        Next

        DestinationVolume = lngVolMax
    Else
        DestinationVolume = udtVol(channel).fValue / (udtMC.lMaximum - udtMC.lMinimum) * 100
    End If
End Property

Public Property Let DestinationVolume( _
    ByVal index As Long, _
    Optional ByVal channel As Long = -1, _
    ByVal lngPercent As Long _
)

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtVol()    As MIXERCONTROLDETAILS_UNSIGNED
    Dim i           As Long

    If hMixer = 0 Then
        Exit Property
    End If

    ' information about destination
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        Exit Property
    End If

    ' get Control Id of volume setting for destination
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        Exit Property
    End If

    ' space for the volume of every channel
    ReDim udtVol(udtMXL.cChannels - 1) As MIXERCONTROLDETAILS_UNSIGNED

    ' get the volume for every channel in the destination
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = udtMXL.cChannels
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtVol(0))
    udtMXCD.paDetails = VarPtr(udtVol(0))

    If Not MM_IS_OK(mixerGetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        Exit Property
    End If

    ' set the volume for every or one channeö
    If channel = -1 Then
        For i = 0 To udtMXL.cChannels - 1
            udtVol(i).fValue = ((CDbl(lngPercent) / 100#) + udtMC.lMinimum) * (udtMC.lMaximum - udtMC.lMinimum)
        Next
    Else
        udtVol(channel).fValue = ((CDbl(lngPercent) / 100#) + udtMC.lMinimum) * (udtMC.lMaximum - udtMC.lMinimum)
    End If

    If Not MM_IS_OK(mixerSetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        Exit Property
    End If
End Property

Public Property Get SourceVolume( _
    ByVal destination As Long, _
    ByVal index As Long, _
    Optional ByVal channel As Long = -1 _
) As Long

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtVol()    As MIXERCONTROLDETAILS_UNSIGNED
    Dim i           As Long
    Dim lngVolMax   As Long
    Dim lngPer      As Long

    If hMixer = 0 Then
        SourceVolume = -1
        Exit Property
    End If

    ' information about source
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        SourceVolume = -1
        Exit Property
    End If

    ' get Control Id of volume setting for source
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        SourceVolume = -1
        Exit Property
    End If

    ' space for volume of every channel
    ReDim udtVol(udtMXL.cChannels - 1) As MIXERCONTROLDETAILS_UNSIGNED

    ' get volume for every channel
    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = udtMXL.cChannels
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtVol(0))
    udtMXCD.paDetails = VarPtr(udtVol(0))

    If Not MM_IS_OK(mixerGetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        SourceVolume = -1
        Exit Property
    End If

    ' volume in percent
    If channel = -1 Then
        ' all channels requested, return average
        For i = 0 To udtMXL.cChannels - 1
            lngPer = (udtVol(i).fValue / (udtMC.lMaximum - udtMC.lMinimum) * 100)

            If lngPer > lngVolMax Then
                lngVolMax = lngPer
            End If
        Next

        SourceVolume = lngVolMax
    Else
        SourceVolume = udtVol(channel).fValue / (udtMC.lMaximum - udtMC.lMinimum) * 100
    End If
End Property

Public Property Let SourceVolume( _
    ByVal destination As Long, _
    ByVal index As Long, _
    Optional ByVal channel As Long = -1, _
    ByVal lngPercent As Long _
)

    Dim udtMXCD     As MIXERCONTROLDETAILS
    Dim udtMXLC     As MIXERLINECONTROLS
    Dim udtMXL      As MIXERLINE
    Dim udtMC       As MIXERCONTROL
    Dim udtVol()    As MIXERCONTROLDETAILS_UNSIGNED
    Dim i           As Long

    If hMixer = 0 Then
        Exit Property
    End If

    ' information about source
    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        Exit Property
    End If

    ' get Control Id of volume setting for source
    udtMXLC.cbStruct = Len(udtMXLC)
    udtMXLC.dwLineID = udtMXL.dwLineID
    udtMXLC.dwControl = MIXERCONTROL_CONTROLTYPE_VOLUME
    udtMXLC.cControls = 1
    udtMXLC.cbmxctrl = LenB(udtMC)
    udtMXLC.pamxctrl = VarPtr(udtMC)

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         udtMXLC, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        Exit Property
    End If

    ' space for volume of every channel
    ReDim udtVol(udtMXL.cChannels - 1) As MIXERCONTROLDETAILS_UNSIGNED

    udtMXCD.cbStruct = Len(udtMXCD)
    udtMXCD.cChannels = udtMXL.cChannels
    udtMXCD.dwControlID = udtMC.dwControlID
    udtMXCD.cbDetails = Len(udtVol(0))
    udtMXCD.paDetails = VarPtr(udtVol(0))

    If Not MM_IS_OK(mixerGetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        Exit Property
    End If

    ' return volume of all or one channel
    If channel = -1 Then
        For i = 0 To udtMXL.cChannels - 1
            udtVol(i).fValue = ((CDbl(lngPercent) / 100#) + udtMC.lMinimum) * (udtMC.lMaximum - udtMC.lMinimum)
        Next
    Else
        udtVol(channel).fValue = ((CDbl(lngPercent) / 100#) + udtMC.lMinimum) * (udtMC.lMaximum - udtMC.lMinimum)
    End If

    If Not MM_IS_OK(mixerSetControlDetails(hMixer, _
                                           udtMXCD, _
                                           MIXER_GETCONTROLDETAILSF_VALUE)) Then

        Exit Property
    End If
End Property

Public Property Get DestinationType( _
    ByVal index As Long _
) As MIXER_DESTINATIONS

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        DestinationType = -1
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        DestinationType = -1
    Else
        DestinationType = udtMXL.dwComponentType
    End If
End Property

Public Property Get SourceType( _
    ByVal destination As Long, _
    ByVal index As Long _
) As MIXER_SOURCES

    Dim udtMXL  As MIXERLINE

    If hMixer = 0 Then
        SourceType = -1
        Exit Property
    End If

    udtMXL.cbStruct = Len(udtMXL)
    udtMXL.dwDestination = destination
    udtMXL.dwSource = index

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     udtMXL, _
                                     MIXER_GETLINEINFOF_SOURCE)) Then

        SourceType = -1
    Else
        SourceType = udtMXL.dwComponentType
    End If
End Property

' ' http://www.ureader.com/message/1359669.aspx
Public Property Get SourceSelected( _
    ByVal destination As Long, _
    ByVal index As Long _
) As Boolean

    Dim mxl         As MIXERLINE
    Dim controls    As MIXERLINECONTROLS
    Dim control()   As MIXERCONTROL
    Dim cd          As MIXERCONTROLDETAILS
    Dim lv()        As MIXERCONTROLDETAILS_BOOLEAN
    Dim c           As Long
    Dim j           As Long

    If hMixer = 0 Then
        SourceSelected = -1
        Exit Property
    End If

    mxl.cbStruct = Len(mxl)
    mxl.dwDestination = destination

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     mxl, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        SourceSelected = -1
        Exit Property
    End If

    ReDim control(mxl.cControls - 1) As MIXERCONTROL

    controls.cbStruct = Len(controls)
    controls.dwLineID = mxl.dwLineID
    controls.cControls = mxl.cControls
    controls.cbmxctrl = Len(control(0))
    controls.pamxctrl = VarPtr(control(0))
    controls.dwControl = MIXERCONTROL_CONTROLTYPE_MUX

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         controls, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        Exit Property
    End If

    For c = 0 To controls.cControls - 1
        If (MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_MUX And control(c).dwControlType)) Then
            ReDim lv(control(c).cMultipleItems - 1) As MIXERCONTROLDETAILS_BOOLEAN

            cd.cbStruct = Len(cd)
            cd.dwControlID = control(c).dwControlID
            cd.cChannels = 1
            cd.item = control(c).cMultipleItems
            cd.cbDetails = Len(lv(0))
            cd.paDetails = VarPtr(lv(0))

            mixerGetControlDetails hMixer, _
                                   cd, _
                                   MIXER_GETCONTROLDETAILSF_VALUE

            For j = 0 To cd.item - 1
                If j = (SourceCount(destination) - index - 1) Then
                    SourceSelected = lv(j).fValue = 1
                    Exit For
                End If
            Next
        End If
    Next
End Property

Public Property Let SourceSelected( _
    ByVal destination As Long, _
    ByVal index As Long, _
    ByVal blnVal As Boolean _
)

    Dim mxl         As MIXERLINE
    Dim controls    As MIXERLINECONTROLS
    Dim control()   As MIXERCONTROL
    Dim cd          As MIXERCONTROLDETAILS
    Dim lv()        As MIXERCONTROLDETAILS_BOOLEAN
    Dim c           As Long
    Dim j           As Long

    If hMixer = 0 Then
        Exit Property
    End If

    mxl.cbStruct = Len(mxl)
    mxl.dwDestination = destination

    If Not MM_IS_OK(mixerGetLineInfo(hMixer, _
                                     mxl, _
                                     MIXER_GETLINEINFOF_DESTINATION)) Then

        Exit Property
    End If

    ReDim control(mxl.cControls - 1) As MIXERCONTROL

    controls.cbStruct = Len(controls)
    controls.dwLineID = mxl.dwLineID
    controls.cControls = mxl.cControls
    controls.cbmxctrl = Len(control(0))
    controls.pamxctrl = VarPtr(control(0))
    controls.dwControl = MIXERCONTROL_CONTROLTYPE_MUX

    If Not MM_IS_OK(mixerGetLineControls(hMixer, _
                                         controls, _
                                         MIXER_GETLINECONTROLSF_ONEBYTYPE)) Then

        Exit Property
    End If

    For c = 0 To controls.cControls - 1
        If MIXERCONTROL_CONTROLTYPE_MUX = (MIXERCONTROL_CONTROLTYPE_MUX And control(c).dwControlType) Then
            ReDim lv(control(c).cMultipleItems - 1) As MIXERCONTROLDETAILS_BOOLEAN

            cd.cbStruct = Len(cd)
            cd.dwControlID = control(c).dwControlID
            cd.cChannels = 1
            cd.item = control(c).cMultipleItems
            cd.cbDetails = Len(lv(0))
            cd.paDetails = VarPtr(lv(0))

            mixerGetControlDetails hMixer, _
                                   cd, _
                                   MIXER_GETCONTROLDETAILSF_VALUE

            For j = 0 To cd.item - 1
                lv(j).fValue = Abs(CBool(j = (SourceCount(destination) - index - 1)))
            Next

            mixerSetControlDetails hMixer, _
                                   cd, _
                                   MIXER_SETCONTROLDETAILSF_VALUE
        End If
    Next
End Property

Private Function MM_IS_OK( _
    ByVal errnr As Long _
) As Boolean

    MM_IS_OK = errnr = 0
End Function

Private Function TrimEx( _
    ByVal strVal As String _
) As String

    TrimEx = Trim$(Left$(strVal, InStr(strVal, Chr$(0)) - 1))
End Function

Private Sub Class_Terminate()
    DeviceClose
End Sub
