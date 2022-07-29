##
## *  Microsoft Excel Developer's Toolkit
## *  Version 15.0
## *
## *  File:           INCLUDE\XLCALL.H
## *  Description:    Header file for for Excel callbacks
## *  Platform:       Microsoft Windows
## *
## *  DEPENDENCY:
## *  Include <windows.h> before you include this.
## *
## *  This file defines the constants and
## *  data types which are used in the
## *  Microsoft Excel C API.
## *
##

##
## * XL 12 Basic Datatypes
##

type
  BOOL* = INT32

##  Boolean

type
  XCHAR* = WCHAR

##  Wide Character

type
  RW* = INT32

##  XL 12 Row

type
  COL* = INT32

##  XL 12 Column

type
  IDSHEET* = DWORD_PTR

##  XL12 Sheet ID
##
## * XLREF structure
## *
## * Describes a single rectangular reference.
##

type
  XLREF* {.bycopy.} = object
    rwFirst*: WORD
    rwLast*: WORD
    colFirst*: BYTE
    colLast*: BYTE

  LPXLREF* = ptr XLREF

##
## * XLMREF structure
## *
## * Describes multiple rectangular references.
## * This is a variable size structure, default
## * size is 1 reference.
##

type
  XLMREF* {.bycopy.} = object
    count*: WORD
    reftbl*: array[1, XLREF]    ##  actually reftbl[count]

  LPXLMREF* = ptr XLMREF

##
## * XLREF12 structure
## *
## * Describes a single XL 12 rectangular reference.
##

type
  XLREF12* {.bycopy.} = object
    rwFirst*: RW
    rwLast*: RW
    colFirst*: COL
    colLast*: COL

  LPXLREF12* = ptr XLREF12

##
## * XLMREF12 structure
## *
## * Describes multiple rectangular XL 12 references.
## * This is a variable size structure, default
## * size is 1 reference.
##

type
  XLMREF12* {.bycopy.} = object
    count*: WORD
    reftbl*: array[1, XLREF12]  ##  actually reftbl[count]

  LPXLMREF12* = ptr XLMREF12

##
## * FP structure
## *
## * Describes FP structure.
##

type
  FP* {.bycopy.} = object
    rows*: cushort
    columns*: cushort
    array*: array[1, cdouble]   ##  Actually, array[rows][columns]


##
## * FP12 structure
## *
## * Describes FP structure capable of handling the big grid.
##

type
  FP12* {.bycopy.} = object
    rows*: INT32
    columns*: INT32
    array*: array[1, cdouble]   ##  Actually, array[rows][columns]


##
## * XLOPER structure
## *
## * Excel's fundamental data type: can hold data
## * of any type. Use "R" as the argument type in the
## * REGISTER function.
##

type
  INNER_C_STRUCT_XLCALL_177* {.bycopy.} = object
    count*: WORD               ##  always = 1
    `ref`*: XLREF

  INNER_C_STRUCT_XLCALL_177* {.bycopy.} = object
    lpmref*: ptr XLMREF
    idSheet*: IDSHEET

  INNER_C_STRUCT_XLCALL_177* {.bycopy.} = object
    lparray*: ptr xloper
    rows*: WORD
    columns*: WORD

  INNER_C_UNION_XLCALL_177* {.bycopy, union.} = object
    level*: cshort             ##  xlflowRestart
    tbctrl*: cshort            ##  xlflowPause
    idSheet*: IDSHEET          ##  xlflowGoto

  INNER_C_STRUCT_XLCALL_177* {.bycopy.} = object
    valflow*: INNER_C_UNION_XLCALL_177
    rw*: WORD                  ##  xlflowGoto
    col*: BYTE                 ##  xlflowGoto
    xlflow*: BYTE

  INNER_C_UNION_XLCALL_177* {.bycopy, union.} = object
    lpbData*: ptr BYTE          ##  data passed to XL
    hdata*: HANDLE             ##  data returned from XL

  INNER_C_STRUCT_XLCALL_177* {.bycopy.} = object
    h*: INNER_C_UNION_XLCALL_177
    cbData*: clong

  INNER_C_UNION_XLCALL_177* {.bycopy, union.} = object
    num*: cdouble              ##  xltypeNum
    str*: LPSTR                ##  xltypeStr
    err*: WORD                 ##  xltypeErr
    w*: cshort                 ##  xltypeInt
    sref*: INNER_C_STRUCT_XLCALL_177 ##  xltypeSRef
    mref*: INNER_C_STRUCT_XLCALL_177 ##  xltypeRef
    array*: INNER_C_STRUCT_XLCALL_177 ##  xltypeMulti
    flow*: INNER_C_STRUCT_XLCALL_177 ##  xltypeFlow
    bigdata*: INNER_C_STRUCT_XLCALL_177 ##  xltypeBigData

  XLOPER* {.bycopy.} = object
    val*: INNER_C_UNION_XLCALL_177
    xltype*: WORD

  LPXLOPER* = ptr XLOPER

##
## * XLOPER12 structure
## *
## * Excel 12's fundamental data type: can hold data
## * of any type. Use "U" as the argument type in the
## * REGISTER function.
##

type
  INNER_C_STRUCT_XLCALL_235* {.bycopy.} = object
    count*: WORD               ##  always = 1
    `ref`*: XLREF12

  INNER_C_STRUCT_XLCALL_235* {.bycopy.} = object
    lpmref*: ptr XLMREF12
    idSheet*: IDSHEET

  INNER_C_STRUCT_XLCALL_235* {.bycopy.} = object
    lparray*: ptr xloper12
    rows*: RW
    columns*: COL

  INNER_C_UNION_XLCALL_235* {.bycopy, union.} = object
    level*: cint               ##  xlflowRestart
    tbctrl*: cint              ##  xlflowPause
    idSheet*: IDSHEET          ##  xlflowGoto

  INNER_C_STRUCT_XLCALL_235* {.bycopy.} = object
    valflow*: INNER_C_UNION_XLCALL_235
    rw*: RW                    ##  xlflowGoto
    col*: COL                  ##  xlflowGoto
    xlflow*: BYTE

  INNER_C_UNION_XLCALL_235* {.bycopy, union.} = object
    lpbData*: ptr BYTE          ##  data passed to XL
    hdata*: HANDLE             ##  data returned from XL

  INNER_C_STRUCT_XLCALL_235* {.bycopy.} = object
    h*: INNER_C_UNION_XLCALL_235
    cbData*: clong

  INNER_C_UNION_XLCALL_235* {.bycopy, union.} = object
    num*: cdouble              ##  xltypeNum
    str*: ptr XCHAR             ##  xltypeStr
    xbool*: BOOL               ##  xltypeBool
    err*: cint                 ##  xltypeErr
    w*: cint
    sref*: INNER_C_STRUCT_XLCALL_235 ##  xltypeSRef
    mref*: INNER_C_STRUCT_XLCALL_235 ##  xltypeRef
    array*: INNER_C_STRUCT_XLCALL_235 ##  xltypeMulti
    flow*: INNER_C_STRUCT_XLCALL_235 ##  xltypeFlow
    bigdata*: INNER_C_STRUCT_XLCALL_235 ##  xltypeBigData

  XLOPER12* {.bycopy.} = object
    val*: INNER_C_UNION_XLCALL_235
    xltype*: DWORD

  LPXLOPER12* = ptr XLOPER12

##
## * XLOPER and XLOPER12 data types
## *
## * Used for xltype field of XLOPER and XLOPER12 structures
##

const
  xltypeNum* = 0x0001
  xltypeStr* = 0x0002
  xltypeBool* = 0x0004
  xltypeRef* = 0x0008
  xltypeErr* = 0x0010
  xltypeFlow* = 0x0020
  xltypeMulti* = 0x0040
  xltypeMissing* = 0x0080
  xltypeNil* = 0x0100
  xltypeSRef* = 0x0400
  xltypeInt* = 0x0800
  xlbitXLFree* = 0x1000
  xlbitDLLFree* = 0x4000
  xltypeBigData* = (xltypeStr or xltypeInt)

##
## * Error codes
## *
## * Used for val.err field of XLOPER and XLOPER12 structures
## * when constructing error XLOPERs and XLOPER12s
##

const
  xlerrNull* = 0
  xlerrDiv0* = 7
  xlerrValue* = 15
  xlerrRef* = 23
  xlerrName* = 29
  xlerrNum* = 36
  xlerrNA* = 42
  xlerrGettingData* = 43

##
## * Flow data types
## *
## * Used for val.flow.xlflow field of XLOPER and XLOPER12 structures
## * when constructing flow-control XLOPERs and XLOPER12s
##

const
  xlflowHalt* = 1
  xlflowGoto* = 2
  xlflowRestart* = 8
  xlflowPause* = 16
  xlflowResume* = 64

##
## * Return codes
## *
## * These values can be returned from Excel4(), Excel4v(), Excel12() or Excel12v().
##

const
  xlretSuccess* = 0
  xlretAbort* = 1
  xlretInvXlfn* = 2
  xlretInvCount* = 4
  xlretInvXloper* = 8
  xlretStackOvfl* = 16
  xlretFailed* = 32
  xlretUncalced* = 64
  xlretNotThreadSafe* = 128
  xlretInvAsynchronousContext* = 256
  xlretNotClusterSafe* = 512

##
## * XLL events
## *
## * Passed in to an xlEventRegister call to register a corresponding event.
##

const
  xleventCalculationEnded* = 1
  xleventCalculationCanceled* = 2

##
## * Function prototypes
##

## !!!Ignored construct:  int _cdecl Excel4 ( int xlfn , LPXLOPER operRes , int count , ... ) ;
## Error: token expected: ; but got: [identifier]!!!

##  followed by count LPXLOPERs

## !!!Ignored construct:  int pascal Excel4v ( int xlfn , LPXLOPER operRes , int count , LPXLOPER opers [ ] ) ;
## Error: token expected: ; but got: [identifier]!!!

## !!!Ignored construct:  int pascal XLCallVer ( void ) ;
## Error: token expected: ; but got: [identifier]!!!

## !!!Ignored construct:  long pascal LPenHelper ( int wCode , VOID * lpv ) ;
## Error: token expected: ; but got: [identifier]!!!

## !!!Ignored construct:  int _cdecl Excel12 ( int xlfn , LPXLOPER12 operRes , int count , ... ) ;
## Error: token expected: ; but got: [identifier]!!!

##  followed by count LPXLOPER12s

## !!!Ignored construct:  int pascal Excel12v ( int xlfn , LPXLOPER12 operRes , int count , LPXLOPER12 opers [ ] ) ;
## Error: token expected: ; but got: [identifier]!!!

##
## * Cluster Connector Async Callback
##

## !!!Ignored construct:  typedef int ( CALLBACK * PXL_HPC_ASYNC_CALLBACK ) ( LPXLOPER12 operAsyncHandle , LPXLOPER12 operReturn ) ;
## Error: token expected: ) but got: *!!!

##
## * Cluster connector entry point return codes
##

const
  xlHpcRetSuccess* = 0
  xlHpcRetSessionIdInvalid* = -1
  xlHpcRetCallFailed* = -2

##
## * Function number bits
##

const
  xlCommand* = 0x8000
  xlSpecial* = 0x4000
  xlIntl* = 0x2000
  xlPrompt* = 0x1000

##
## * Auxiliary function numbers
## *
## * These functions are available only from the C API,
## * not from the Excel macro language.
##

const
  xlFree* = (0 or xlSpecial)
  xlStack* = (1 or xlSpecial)
  xlCoerce* = (2 or xlSpecial)
  xlSet* = (3 or xlSpecial)
  xlSheetId* = (4 or xlSpecial)
  xlSheetNm* = (5 or xlSpecial)
  xlAbort* = (6 or xlSpecial)
  xlGetInst* = (7 or xlSpecial)   ##  Returns application's hinstance as an integer value, supported on 32-bit platform only
  xlGetHwnd* = (8 or xlSpecial)
  xlGetName* = (9 or xlSpecial)
  xlEnableXLMsgs* = (10 or xlSpecial)
  xlDisableXLMsgs* = (11 or xlSpecial)
  xlDefineBinaryName* = (12 or xlSpecial)
  xlGetBinaryName* = (13 or xlSpecial)

##  GetFooInfo are valid only for calls to LPenHelper

const
  xlGetFmlaInfo* = (14 or xlSpecial)
  xlGetMouseInfo* = (15 or xlSpecial)
  xlAsyncReturn* = (16 or xlSpecial) ## Set return value from an asynchronous function call
  xlEventRegister* = (17 or xlSpecial) ## Register an XLL event
  xlRunningOnCluster* = (18 or xlSpecial) ## Returns true if running on Compute Cluster
  xlGetInstPtr* = (19 or xlSpecial) ##  Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms

##  edit modes

const
  xlModeReady* = 0
  xlModeEnter* = 1
  xlModeEdit* = 2
  xlModePoint* = 4

##  document(page) types

const
  dtNil* = 0x7f
  dtSheet* = 0
  dtProc* = 1
  dtChart* = 2
  dtBasic* = 6

##  hit test codes

const
  htNone* = 0x00
  htClient* = 0x01
  htVSplit* = 0x02
  htHSplit* = 0x03
  htColWidth* = 0x04
  htRwHeight* = 0x05
  htRwColHdr* = 0x06
  htObject* = 0x07
  htTopLeft* = 0x08
  htBotLeft* = 0x09
  htLeft* = 0x0A
  htTopRight* = 0x0B
  htBotRight* = 0x0C
  htRight* = 0x0D
  htTop* = 0x0E
  htBot* = 0x0F

##  end size handles

const
  htRwGut* = 0x10
  htColGut* = 0x11
  htTextBox* = 0x12
  htRwLevels* = 0x13
  htColLevels* = 0x14
  htDman* = 0x15
  htDmanFill* = 0x16
  htXSplit* = 0x17
  htVertex* = 0x18
  htAddVtx* = 0x19
  htDelVtx* = 0x1A
  htRwHdr* = 0x1B
  htColHdr* = 0x1C
  htRwShow* = 0x1D
  htColShow* = 0x1E
  htSizing* = 0x1F
  htSxpivot* = 0x20
  htTabs* = 0x21
  htEdit* = 0x22

type
  FMLAINFO* {.bycopy.} = object
    wPointMode*: cint          ##  current edit mode.  0 => rest of struct undefined
    cch*: cint                 ##  count of characters in formula
    lpch*: cstring             ##  pointer to formula characters.  READ ONLY!!!
    ichFirst*: cint            ##  char offset to start of selection
    ichLast*: cint             ##  char offset to end of selection (may be > cch)
    ichCaret*: cint            ##  char offset to blinking caret

  MOUSEINFO* {.bycopy.} = object
    hwnd*: HWND                ##  input section
    ##  window to get info on
    pt*: POINT                 ##  mouse position to get info on
             ##  output section
    dt*: cint                  ##  document(page) type
    ht*: cint                  ##  hit test code
    rw*: cint                  ##  row @ mouse (-1 if #n/a)
    col*: cint                 ##  col @ mouse (-1 if #n/a)


##
## * User defined function
## *
## * First argument should be a function reference.
##

const
  xlUDF* = 255

