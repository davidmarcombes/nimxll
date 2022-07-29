
## XLOPER12 Definition
##
## This module wraps the C XLOPER12 structure
## and provides utility functions to use XLOPER12

# Windows types 
type
  XloperWideChar* = uint16 # wchar_t
  XloperInt32* = cint # INT32 
  XloperDwordPtr* = ByteAddress # DWORDPTR
  XloperDWord* = int32
  XloperWord* = int16
  XloperByte* = uint8
  XloperHandle* = int

# Xloper types
type
  Xloper12Bool* = XloperInt32 
  Xloper12IdSheet* = XloperDwordPtr

# Reference
type
  XLREF12* {.bycopy.} = object
    rwFirst*: XloperInt32
    rwLast*: XloperInt32
    colFirst*: XloperInt32
    colLast*: XloperInt32

  LPXLREF12* = ptr XLREF12

# Multireference
type
  XLMREF12* {.bycopy.} = object
    count*: XloperWord
    reftbl*: array[1, XLREF12]  ##  actually reftbl[count]

  LPXLMREF12* = ptr XLMREF12

# FP structure capable of handling the big grid.
type
  FP12* {.bycopy.} = object
    rows*: XloperInt32
    columns*: XloperInt32
    arraydef*: array[1, cdouble]   ##  Actually, array[rows][columns]

# XLOPER12
type
  Xloper12SRef* {.bycopy.} = object
    count*: XloperWord               ##  always = 1
    `ref`*: XLREF12

  Xloper12Ref* {.bycopy.} = object
    lpmref*: ptr XLMREF12
    Xloper12IdSheet*: Xloper12IdSheet

  Xloper12Multi* {.bycopy.} = object
    lparray*: ptr Xloper12
    rows*: XloperInt32
    columns*: XloperInt32

  Xloper12Valflow* {.bycopy, union.} = object
    level*: cint               ##  xlflowRestart
    tbctrl*: cint              ##  xlflowPause
    Xloper12IdSheet*: Xloper12IdSheet          ##  xlflowGoto

  Xloper12Flow* {.bycopy.} = object
    valflow*: Xloper12Valflow
    rw*: XloperInt32                    ##  xlflowGoto
    col*: XloperInt32                  ##  xlflowGoto
    xlflow*: XloperByte

  Xloper12Handle* {.bycopy, union.} = object
    lpbData*: ptr XloperByte          ##  data passed to XL
    hdata*: XloperHandle             ##  data returned from XL

  Xloper12BigData* {.bycopy.} = object
    h*: Xloper12Handle
    cbData*: clong

  Xloper12Inner* {.bycopy, union.} = object
    ## Interal data for an XLOPER 12
    num*: cdouble              ##  xltypeNum
    str*: ptr XloperWideChar             ##  xltypeStr
    bl*: Xloper12Bool               ##  xltypeXloper12Bool
    err*: cint                 ##  xltypeErr
    w*: cint
    sref*: Xloper12SRef ##  xltypeSRef
    mref*: Xloper12Ref ##  xltypeRef
    multi*: Xloper12Multi ##  xltypeMulti
    flow*: Xloper12Flow ##  xltypeFlow
    bigdata*: Xloper12BigData ##  xltypeBigData

  Xloper12* {.bycopy.} = object
    ## XLOPER12 is a variant style object
    val*: Xloper12Inner
    xltype*: XloperDWord

  PXloper12* = ptr XLOPER12

# Data types
const
  xltypeNum = 0x0001
  xltypeStr = 0x0002
  xltypeBool = 0x0004
  xltypeRef = 0x0008
  xltypeErr = 0x0010
  xltypeFlow = 0x0020
  xltypeMulti = 0x0040
  xltypeMissing = 0x0080
  xltypeNil = 0x0100
  xltypeSRef = 0x0400
  xltypeInt = 0x0800
  xlbitXLFree = 0x1000
  xlbitDLLFree = 0x4000
  xltypeBigData = (xltypeStr or xltypeInt)

# Error codes
const
  xlerrNull = 0
  xlerrDiv0 = 7
  xlerrValue = 15
  xlerrRef = 23
  xlerrName = 29
  xlerrNum = 36
  xlerrNA = 42
  xlerrGettingData = 43

# Nim Error Enum
type
  ExcelError* {.pure.} = enum
    ## Excel Error types
    Null,  #NULL 
    Div0,  #DIV0
    Value,  #VALUE
    Ref,  #REF
    Name,  #NAME
    Num,  #NUM
    Na,  #NA
    GettingData,  #GETTINGDATA
    Unknown

func errorCodeToEnum(e: cint): ExcelError {.inline.} =
  case e: 
    of xlerrNull: result = Null
    of xlerrDiv0: result = Div0
    of xlerrValue: result = Value
    of xlerrRef: result = Ref
    of xlerrName: result = Name
    of xlerrNum: result = Num
    of xlerrNA: result = Na
    of xlerrGettingData: result = GettingData
    else: result = Unknown

# Utility functions

func isNumeric*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of numerical type
  result = xlop.xltype == xltypeNum

func isString*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of string type
  result = xlop.xltype == xltypeStr

func isBool*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of boolean type
  result = xlop.xltype == xltypeBool

func isReference*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of reference type 
  result = xlop.xltype == xltypeRef

func isError*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of error type
  result = xlop.xltype == xltypeErr

# Flow ???

# Multi ???

func isMissing*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of missing type
  result = xlop.xltype == xltypeErr

func isNil*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of nil type
  result = xlop.xltype == xltypeNil

# SREf ???

func isInt*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of integer type
  result = xlop.xltype == xltypeInt

func isBigData*(xlop: Xloper12): bool {.inline.} = 
  ## Check if an Xloper12 is of big data type
  result = xlop.xltype == xltypeBigData

func getCDouble*(xlop: Xloper12): cdouble {.inline.} = 
  ## Get a cdouble from an Xloper12 
  ## 
  ## This function does not check the actual internal type
  result = xlop.val.num 

func getFloat*(xlop: Xloper12): float {.inline.} =
  ## Get a float from an Xloper12 
  ## 
  ## This function does not check the actual internal type
  result = xlop.val.num.float

func getFloat32*(xlop: Xloper12): float32 {.inline.} =
  ## Get a float32 from an Xloper12 
  ## 
  ## This function does not check the actual internal type
  result = xlop.val.num.float32

func getFloat64*(xlop: Xloper12): float64 {.inline.} =
  ## Get a float64 from an Xloper12 
  ## 
  ## This function does not check the actual internal type
  result = xlop.val.num.float64

func getErrorCode*(xlop: Xloper12): cint {.inline.} =
  ## Get the error code from an Xloper12
  ##  
  ## This function does not check the actual internal type
  ## 
  ## Prefer the getError function
  result = xlop.val.err

func getError*(xlop: Xloper12): ExcelError {.inline.} =
  ## Get the error from an Xloper12
  ##  
  ## This function does not check the actual internal type
  result = errorCodeToEnum(xlop.val.err)
