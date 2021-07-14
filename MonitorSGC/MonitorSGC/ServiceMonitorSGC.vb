
Imports DataLibrarySGC

Public Class ServiceMonitorSGC
    Private WithEvents myTimer As System.Timers.Timer

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        Me.myTimer = New System.Timers.Timer()

        Me.myTimer.Enabled = True
        'ejecute cada 5 minutos
        Me.myTimer.Interval = 1000 * 60 * 5

        Me.myTimer.Start()
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub

    Protected Sub myTimer_Elapsed(ByVal sender As Object, e As EventArgs) Handles myTimer.Elapsed
        Dim data As New DataAccess
        Dim bussines As New ServicioBussinesCore

        'todas las tareas activas, incumplidas o finalizadas se omiten
        'de esta manera solo monitoreamos la ultima de la cadena
        Dim listaTareas = data.obtenerTareaSinFinalizar

        For Each tarea As MonitorTarea In listaTareas
            bussines.validaTiempoTarea(tarea)
        Next

    End Sub

End Class

