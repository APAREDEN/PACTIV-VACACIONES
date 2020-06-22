
Public Class frmPrincipal
    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnProcesar01_Click(sender As Object, e As EventArgs) Handles btnProcesar01.Click, btnProcesar02.Click, btnProcesar03.Click, btnProcesar04.Click
        Select Case sender.name
            Case "btnProcesar01"
                VerDetalle(bsProgramacion, dtFecha.Value)
            Case "btnProcesar02"
                VerDetalle(bsProgramacion1, dtFecha1.Value)
            Case "btnProcesar03"
                VerDetalle(bsProgramacion2, dtFecha2.Value)
            Case "btnProcesar04"
                VerDetalle(bsProgramacion3, dtFecha3.Value)
        End Select


    End Sub
    Private Sub VerDetalle(ByRef bs As BindingSource, fechaInicial As DateTime)
        Dim ListaVacacion As New List(Of EntityProgramacion)
        Dim vacacion As EntityProgramacion
        bs.DataSource = Nothing
        Dim iNumDiasXsalir As Integer = 0, iContador As Integer = 0, iNumVac As Integer = Val(txtDvaca.Text), iNumBonus As Integer = Val(txtDbonus.Text), fecha As DateTime = fechaInicial, bSalida As Boolean

        Dim descanso As Boolean
        'Preparar fechas anteriores en el orden correcto

        Dim myAL As New ArrayList()
        fecha = DateAdd(DateInterval.Day, -1, fecha)

        bSalida = False
        iContador = 0
        descanso = GetDia(fecha)
        While descanso
            myAL.Add(fecha)
            fecha = DateAdd(DateInterval.Day, -1, fecha)
            descanso = GetDia(fecha)
        End While
        myAL.Sort()
        'Evalua días anteriores por si son descanso para agregarlos
        For Each fecha In myAL
            vacacion = New EntityProgramacion
            vacacion.Descanso = True
            iNumDiasXsalir += 1
            vacacion.Dia = iNumDiasXsalir
            vacacion.Fecha = fecha
            vacacion.NomDia = Format(fecha, "dddd")
            ListaVacacion.Add(vacacion)
        Next




        fecha = fechaInicial
        bSalida = vbFalse
        iContador = 0




        'Evalua Vacaciones
        If iNumVac > 0 Then
            bSalida = False
            While Not bSalida
                vacacion = New EntityProgramacion
                vacacion.Descanso = GetDia(fecha)
                iNumDiasXsalir += 1
                vacacion.Dia = iNumDiasXsalir
                vacacion.Fecha = fecha
                vacacion.NomDia = Format(fecha, "dddd")
                If vacacion.Descanso Then
                Else
                    iContador += 1
                    vacacion.TipoDiaTomado = "V" & Microsoft.VisualBasic.Right("00" + iContador.ToString.Trim, 2)
                End If
                ListaVacacion.Add(vacacion)
                fecha = DateAdd(DateInterval.Day, 1, fecha)
                If iContador = iNumVac Then
                    bSalida = True
                End If
            End While
        End If

        'EvaluaBonusDay
        If iNumBonus > 0 Then
            bSalida = False
            iContador = 0
            While Not bSalida
                vacacion = New EntityProgramacion
                vacacion.Descanso = GetDia(fecha)
                iNumDiasXsalir += 1
                vacacion.Dia = iNumDiasXsalir
                vacacion.Fecha = fecha
                vacacion.NomDia = Format(fecha, "dddd")
                If vacacion.Descanso Then
                Else
                    iContador += 1
                    vacacion.TipoDiaTomado = "B" & Microsoft.VisualBasic.Right("00" + iContador.ToString.Trim, 2)
                End If
                ListaVacacion.Add(vacacion)
                fecha = DateAdd(DateInterval.Day, 1, fecha)
                If iContador = iNumBonus Then
                    bSalida = True
                End If
            End While
        End If


        'Evalua días siguientes por si son descanso para agregarlos
        bSalida = False
        iContador = 0
        descanso = GetDia(fecha)
        While descanso
            vacacion = New EntityProgramacion
            vacacion.Descanso = descanso
            iNumDiasXsalir += 1
            vacacion.Dia = iNumDiasXsalir
            vacacion.Fecha = fecha
            vacacion.NomDia = Format(fecha, "dddd")
            ListaVacacion.Add(vacacion)
            fecha = DateAdd(DateInterval.Day, 1, fecha)
            descanso = GetDia(fecha)
        End While
        bs.DataSource = ListaVacacion
    End Sub
    Private Function GetDia(fecha As Date) As Boolean
        Dim FecBase As Date = CDate("2013-12-30")
        Dim Horario = {"D", "D", "", "", "D", "D", "D", "", "", "D", "D", "", "", ""}
        Dim diasAfehaBase As Long
        Dim Posicion As Long
        diasAfehaBase = DateDiff(DateInterval.Day, FecBase, fecha)
        Posicion = (diasAfehaBase Mod 14)
        If Horario(Posicion) = "D" Then Return True
        Return False
    End Function


End Class