Set WMI = GetObject("WinMgmts:root\WMI")
Set Objs = WMI.ExecQuery("SELECT * FROM MSStorageDriver_ATAPISmartData")
For Each Obj In Objs
    a = Obj.VendorSpecific
Next
For i = 2 To UBound(a)
    If a(i) = 4 Then
    MsgBox "Start stop count " & (a(i+11)+a(i+10)+a(i+9)+a(i+8)+a(i+7)+a(i+6))*256+a(i+5) & " times"
    End If
    If a(i) = 9 Then
    MsgBox "Power on " & (a(i+11)+a(i+10)+a(i+9)+a(i+8)+a(i+7)+a(i+6))*256+a(i+5) & " hours"
    End If
Next
