
when not defined(windows):
  {.error: "xlcall is a windows only module".}


import winlean

# https://github.com/dom96/osinfo/blob/master/src/osinfo/win.nim

proc getProcAddress(hModule: int, lpProcName: cstring): pointer{.stdcall,
    dynlib: "kernel32", importc: "GetProcAddress".}

proc getModuleHandleA(lpModuleName: cstring): int{.stdcall,
     dynlib: "kernel32", importc: "GetModuleHandleA".}

var pexcel12 {.global.}: pointer = nil

proc fetchExcel12EntryPt(): void {.inline.} =
  if pexcel12 == nil:
    let hmodule = getModuleHandleA(nil)
    if hmodule != 0:
      pexcel12 = getProcAddress(hmodule, "MdCallBack12")

type
    Xloper {.final, incompleteStruct, header: "<xlcall.h>", importc: "xloper".} = object
    PXloper {.header: "<xlcall.h>", importc: "LPXLOPER".} = ptr Xloper

proc excel12*(excelFunction: int, 
              result: PXloper,
              params: varargs[PXloper]
              ) int =
  result = 0