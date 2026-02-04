VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Створено: Окланд, 110 ОМБр, травень 2023
Option Explicit
Public ppmax, hwnd, Creator
Private PIBs(10000), sPIBs(10000), IPNs(10000), ZVs(10000), POSs(10000), PIDRs(10000), isSHTs(10000), RVKs(10000), Phones(10000), Addrs(10000), Pochatok(10000), Sexs(10000)

Function getSex(n)
    getSex = Sexs(n)
End Function

Sub setSex(n, a)
    If IsError(a) Then a = ""
    Sexs(n) = a
End Sub

Function getPochatok(n)
    getPochatok = Pochatok(n)
End Function

Sub setPochatok(n, a)
    If IsError(a) Then a = ""
    Pochatok(n) = a
End Sub

Function getAddr(n)
    getAddr = Addrs(n)
End Function

Sub setAddr(n, a)
    If IsError(a) Then a = ""
    Addrs(n) = a
End Sub

Function getPhone(n)
    getPhone = Phones(n)
End Function

Sub setPhone(n, a)
    If IsError(a) Then a = ""
    Phones(n) = a
End Sub

Function getPIB(n)
    getPIB = PIBs(n)
End Function

Sub setPIB(n, a)
    If IsError(a) Then a = ""
    PIBs(n) = a
End Sub

Function getsPIB(n)
    getsPIB = sPIBs(n)
End Function

Sub setsPIB(n, a)
    If IsError(a) Then a = ""
    sPIBs(n) = a
End Sub

Function getIPN(n)
    getIPN = IPNs(n)
End Function

Sub setIPN(n, a)
    If IsError(a) Then a = ""
    IPNs(n) = a
End Sub

Function getZV(n)
    getZV = ZVs(n)
End Function

Sub setZV(n, a)
    If IsError(a) Then a = ""
    ZVs(n) = a
End Sub

Function getPOS(n)
    getPOS = POSs(n)
End Function

Sub setPOS(n, a)
    If IsError(a) Then a = ""
    POSs(n) = a
End Sub

Function getPIDR(n)
    getPIDR = PIDRs(n)
End Function

Sub setPIDR(n, a)
    If IsError(a) Then a = ""
    PIDRs(n) = a
End Sub

Function getSHT(n)
    getSHT = isSHTs(n)
End Function

Sub setSHT(n, a)
    If IsError(a) Then a = ""
    isSHTs(n) = a
End Sub

Function getRVK(n)
    getRVK = RVKs(n)
End Function

Sub setRVK(n, a)
    If IsError(a) Then a = ""
    RVKs(n) = a
End Sub

' Створено: Окланд, 110 ОМБр, травень 2023

