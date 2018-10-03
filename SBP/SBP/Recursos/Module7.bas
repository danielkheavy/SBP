Attribute VB_Name = "Module7"

'aqui vamos a poner los comentarios del sistema
'tconcare.Show 1
'parame adicionar aduana contador de agentes
'aduana
'aduanaga
'factura   agregar  aduana 11,dua 11
'detalle   igual
'Obtener Id del procesador de un PC
Private Function CpuId() As String

    Dim computer   As String

    Dim wmi        As Variant

    Dim processors As Variant

    Dim cpu        As Variant

    Dim cpu_ids    As String

    computer = "."
    Set wmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!" & computer & "rootcimv2")
    Set processors = wmi.ExecQuery("Select * from " & "Win32_Processor")

    For Each cpu In processors

        cpu_ids = cpu_ids & ", " & cpu.ProcessorId
    Next cpu

    If Len(cpu_ids) > 0 Then cpu_ids = Mid$(cpu_ids, 3)

    CpuId = cpu_ids

End Function

