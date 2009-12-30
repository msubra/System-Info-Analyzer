VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SYSTEM_INFO
        dwOemID As Long
         
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Const PROCESSOR_ALPHA_21064 = 21064
Const PROCESSOR_ARCHITECTURE_ALPHA = 2
Const PROCESSOR_ARCHITECTURE_INTEL = 0
Const PROCESSOR_ARCHITECTURE_MIPS = 1
Const PROCESSOR_ARCHITECTURE_PPC = 3
Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF
Const PROCESSOR_INTEL_386 = 386
Const PROCESSOR_INTEL_486 = 486
Const PROCESSOR_INTEL_PENTIUM = 586
Const PROCESSOR_MIPS_R4000 = 4000

Public Enum ProcessorFeatures
     PF_FLOATING_POINT_PRECISION_ERRATA = 0
     PF_FLOATING_POINT_EMULATED = 1
     PF_COMPARE_EXCHANGE_DOUBLE = 2
     PF_MMX_INSTRUCTIONS_AVAILABLE = 3
     PF_PPC_MOVEMEM_64BIT_OK = 4
     PF_ALPHA_BYTE_INSTRUCTIONS = 5
End Enum

Private Declare Function IsProcessorFeaturePresent Lib "kernel32.dll" (ByVal ProcessorFeature As ProcessorFeatures) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Dim Info As SYSTEM_INFO
Public Property Get ProcessorType() As String

Dim m_CPUType  As String

Select Case (Info.dwProcessorType)
    Case PROCESSOR_INTEL_386:
        m_CPUType = "386 Processor"
    Case PROCESSOR_INTEL_486:
        m_CPUType = "486 Processor"
    Case PROCESSOR_INTEL_PENTIUM:
        m_CPUType = "Genuine Pentium Processor"
    Case PROCESSOR_MIPS_R4000:
        m_CPUType = "MIPS R4000 Processor"
    Case PROCESSOR_ALPHA_21064:
        m_CPUType = "ALPHA 21064 Processor"
End Select

ProcessorType = m_CPUType
End Property

Public Property Get ProcessorName() As String
Dim m_Microprocessor  As String

Select Case (Info.dwOemID)
    Case PROCESSOR_ARCHITECTURE_INTEL:
        m_Microprocessor = "Intel Microprocessor"
    Case PROCESSOR_ARCHITECTURE_MIPS:
        m_Microprocessor = "MIPS Microprocessor"
    Case PROCESSOR_ARCHITECTURE_ALPHA:
        m_Microprocessor = "ALPHA Microprocessor"
    Case PROCESSOR_ARCHITECTURE_PPC:
        m_Microprocessor = "PowerPC Microprocessor"
    Case PROCESSOR_ARCHITECTURE_UNKNOWN:
        m_Microprocessor = "Unknown Microprocessor"
End Select

ProcessorName = m_Microprocessor
End Property

Public Property Get CheckMathProcessor() As String

Dim v As New clsValues
Dim k As New clsKey

k.hKey = HKEY_LOCAL_MACHINE

k.Path = "Hardware\Description\System\FloatingPointProcessor"
Set v = k.Values

If v.Count <= 0 Then
    CheckMathProcessor = "Not Found"
ElseIf v.Count > 0 Then
    CheckMathProcessor = "Found"
End If

End Property
Public Property Get CountProcessors() As Integer
CountProcessors = Info.dwNumberOrfProcessors
End Property
Public Property Get MemoryPageSize() As Long
 MemoryPageSize = Info.dwPageSize
End Property
Public Property Get ActiveProcessor() As Long
ActiveProcessor = Info.dwActiveProcessorMask
End Property
Public Property Get FloatingPointPrecision() As Boolean
FloatingPointPrecision = IsProcessorFeaturePresent(PF_FLOATING_POINT_PRECISION_ERRATA)
End Property
Public Property Get FloatingPointEmulated() As Boolean
FloatingPointEmulated = IsProcessorFeaturePresent(PF_FLOATING_POINT_EMULATED)
End Property
Public Property Get MMX_Instruction() As Boolean
MMX_Instruction = IsProcessorFeaturePresent(PF_MMX_INSTRUCTIONS_AVAILABLE)
End Property
Public Property Get CompareExchangeDouble() As Boolean
CompareExchangeDouble = IsProcessorFeaturePresent(PF_COMPARE_EXCHANGE_DOUBLE)
End Property
Public Property Get AlphaInstructions() As Boolean
AlphaInstructions = IsProcessorFeaturePresent(PF_ALPHA_BYTE_INSTRUCTIONS)
End Property
Public Property Get MemMoveBit64() As Boolean
MemMoveBit64 = IsProcessorFeaturePresent(PF_PPC_MOVEMEM_64BIT_OK)
End Property
Public Property Get MinimumApplicationAddress() As Long
MinimumApplicationAddress = Info.lpMinimumApplicationAddress
End Property
Public Property Get MaximumApplicationAddress() As Long
MaximumApplicationAddress = Info.lpMaximumApplicationAddress
End Property
Public Property Get PerformanceCounter() As Currency

Dim lret As Long
Dim lCounter As LARGE_INTEGER

lret = QueryPerformanceCounter(lCounter)
PerformanceCounter = LargeIntToCurrency(lCounter)

End Property

Public Property Get PerformanceFrequency() As Currency
Dim lret As Long
Dim lFrequency As LARGE_INTEGER

lret = QueryPerformanceFrequency(lFrequency)
PerformanceFrequency = LargeIntToCurrency(lFrequency)


End Property
Private Sub Class_Initialize()
Call GetSystemInfo(Info)
End Sub
