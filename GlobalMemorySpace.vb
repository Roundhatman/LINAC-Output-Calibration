
Module GlobalMemorySpace

    Public Const P0 As Single = 1013.25
    Public Const T0 As Single = 20.0
    Public Const Ke As Single = 1.0
    Public Const Hp As Single = 1.0
    Public Const fiveLines As String = vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine

    Public currentMachine As String = vbNullString
    Public machinePath As String = Application.StartupPath & "\Machines\"

    Public mouseOffset As Point
    Public isDragging As Boolean = False

    Public Ktp, Kpol, Ks, Kq, Dnw, PDD, M1, M2, M, DW, a0, a1, a2 As Single


End Module
