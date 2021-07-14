Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Threading
Imports DataLibrarySGC.My

Public Class ServicioBussinesCore
    Dim data As New DataAccess

    Private Const EstatusFinal As String = "Realizada"
    Private Const EstatusEscala As String = "Incumplida"
    Private Const EstatusInicial As String = "Asignada"


    'valido si aun tiene tiempo
    Public Sub validaTiempoTarea(ByVal tarea As MonitorTarea)
        Dim fechaValidar = tarea.FechaCreacion
        Dim fechaActual As DateTime = DateTime.Now
        'diferencia entre las dos fechaa
        Dim minutos As Long = DateDiff(DateInterval.Minute, fechaValidar, fechaActual)

        'si esta en uno aun esta activa
        If tarea.Status Then
            Select Case tarea.UnidadTiempo
                Case "Minutos"
                    'tarea se ha sobrepasado
                    If minutos >= tarea.Frecuencia Then
                        escalarTarea(tarea)
                    End If
                Case "Horas"
                    Dim hora As Integer = minutos \ 60
                    If hora >= tarea.Frecuencia Then
                        escalarTarea(tarea)
                    End If
                Case "Dias"
                    Dim dia As Integer = minutos \ 60 \ 24
                    If dia >= tarea.Frecuencia Then
                        escalarTarea(tarea)
                    End If
            End Select
        Else 'si se apago el buzon es que si se termino en tiempo, solo activas
            'ultima tarea en la cadena
            data.actualizaTareaEstatus(tarea.MonitorTareaId, EstatusFinal)

        End If
    End Sub

    Public Sub escalarTarea(ByVal tarea As MonitorTarea)

        Dim JefeEmpleadoId = data.obtenJefeEmpleado(tarea.IdEmpleado)
        'cuando llega al director general se deja de escalar
        If JefeEmpleadoId <> 0 Then
            'tarea incumplida
            data.actualizaTareaEstatus(tarea.MonitorTareaId, EstatusEscala)

            EnviaCorreoEscalamiento(tarea, JefeEmpleadoId)
            'voy a insertar una nueva tarea con el nuevo estatus para el jefe
            tarea.IdEmpleado = JefeEmpleadoId
            tarea.FechaCreacion = Date.Now
            tarea.Status = 1
            tarea.StatusTareaId = data.getEstatusTareaIdByName(EstatusInicial)
            tarea.PadreId = tarea.MonitorTareaId
            'inserto nueva tarea
            data.MonitorTareaInsert(tarea)
        End If
    End Sub

    Private Sub EnviaCorreoEscalamiento(ByVal tarea As MonitorTarea, ByVal JefeId As Integer)
        'Correo al jefe del empleado
        Dim queryMail = String.Format("Select E.nombre NombreEmpleado, E.email , P.nombre NombrePuesto, P.codigo FROM Empleado E INNER JOIN Puesto P On E.IDPuesto = P.IDPuesto WHERE  E.IDEmpleado={0}", JefeId)
        Dim dt As DataTable = data.obtenerDataQuery(queryMail)

        'datos del empleado
        Dim queryEmpleado = String.Format("Select E.nombre Empleado, P.nombre Puesto FROM Empleado E INNER JOIN Puesto P ON E.IDPuesto=P.IDPuesto WHERE E.IDEmpleado ={0}", tarea.IdEmpleado)
        Dim dtEmp = data.obtenerDataQuery(queryEmpleado)
        Dim nombreEmpleado As String = ""
        Dim puestoEmpleado As String = ""
        For Each row As DataRow In dtEmp.Rows
            nombreEmpleado = row("Empleado").ToString
            puestoEmpleado = row("Puesto").ToString
        Next

        'datos de la tarea
        Dim queryTarea As String = String.Format("SELECT NombrePantalla FROM ModulosSGC WHERE Id={0}", tarea.IdVista)
        Dim dtTarea = data.obtenerDataQuery(queryTarea)
        Dim NombreTarea As String = ""
        For Each row As DataRow In dtTarea.Rows
            NombreTarea = row("NombrePantalla").ToString
        Next

        'datos de la Tarea
        Dim queryBuzon As String = String.Format("SELECT IdRegistro FROM BuzonTarea WHERE BuzonTareaId={0}", tarea.BuzonTareaId)
        Dim dtbuzon = data.obtenerDataQuery(queryBuzon)
        Dim NoRegistro As Integer
        For Each row As DataRow In dtbuzon.Rows
            NoRegistro = CInt(row("IdRegistro"))
        Next
        'se envía correos a supervisor area
        For Each row As DataRow In dt.Rows
            Dim nombreCorreo = row("NombreEmpleado").ToString
            Dim puestoCorreo = row("NombrePuesto").ToString
            Dim mailCorreo = row("email").ToString
            Dim codigoPuesto = row("codigo").ToString
            Dim destino As New Destinatario
            With destino
                .Email = mailCorreo
                .Puesto = puestoCorreo
                .Nombre = nombreCorreo
            End With
            'Header del correo
            Dim body As String = "<div style='background:#C83535;font-size:20px;text-align: center;color: white;height:70px'> <h1 style='padding:16px;'> AMERICAN AXLE </h1> </div>" &
                "<div style='background:#E7E7E7;height:40px; margin-top:-20px;'> <h2 style='text-align:center; padding:16px;'></h2> </div>"
            'Sub Header destinatario
            body += String.Format("<div style='background:#F3F3F3; height:auto;'> <br><label style='margin-left:20px;padding:10px;font-weight:bold;'> {0} </label><br><label style='margin-left:20px;padding:10px;font-weight:bold;'> {1} </label> <br><br>", nombreCorreo, puestoCorreo)
            'body
            If codigoPuesto = "DG" Then
                body += String.Format("<p>Por este medio se le comunica que {0}, {1} no cumplio la tarea de {2} para el registro No.{3} en el tiempo establecido, por lo que se solicita su intervención a fin de dar cumplimiento a dicha actividad, para lo cual se establece un tiempo de {4} {5}<p>", nombreEmpleado, puestoEmpleado, NombreTarea, NoRegistro, tarea.Frecuencia, tarea.UnidadTiempo)

            Else
                body += String.Format("<p>Por este medio se le comunica que {0}, {1} no cumplio la tarea de {2} para el registro No.{3} en el tiempo establecido, se le comunica la anterior para solicitar su apoyo a fin de dar cumplimiento o cancelar dicha actividad.<p>", nombreEmpleado, puestoEmpleado, NombreTarea, NoRegistro)

            End If
            body += String.Format("<p>Gracias</p>")

            body += "<div style='background:#E7E7E7;height:50px; margin-top:-30px;'> <h4 style='text-align:center; padding:16px;'> Sistema de Gestión de Calidad Productiva </h4> </div>"
            body += "<div style='background-color:#3C3E43;bottom: 0; width: 100%; height:50px; color: white;'><center><label style='padding-top-20px; margin-top:15px;margin-left:-90px; position:absolute;'>EL CARMEN NUEVO LEON</label></center></div>"

            Dim subject As String = "Incumplimento de la Tarea:" + NombreTarea
            EnviaCorreo(body, destino, subject)

        Next
    End Sub


    Public Sub EnviaCorreo(ByVal Body As String, ByVal destino As Destinatario, ByVal subject As String)

        Dim correos As New MailMessage
        Dim envios As New SmtpClient
        Dim mail As Correo = ObtenerDatosMail()

        correos.To.Clear()
        correos.Body = " "
        correos.Subject = subject
        correos.BodyEncoding = Encoding.GetEncoding("iso-8859-1")
        correos.Body = Body

        correos.IsBodyHtml = True
        correos.To.Add(Trim(destino.Email))
        correos.From = New MailAddress(mail.Direccion)
        envios.Credentials = New NetworkCredential(mail.Direccion, mail.PassWord)
        envios.Host = mail.Host
        envios.Port = mail.Puerto
        envios.EnableSsl = False
        envios.Send(correos)
    End Sub

    Public Function ObtenerDatosMail() As Correo
        Dim mail As New Correo
        If Resources.correocliente = "1" Then
            mail.Direccion = Resources.correosfromcliente
            mail.Host = Resources.Hostcliente
            mail.Puerto = Resources.Puertocliente
        Else
            mail.Direccion = Resources.correosfrominterno
            mail.Host = Resources.HostInterno
            mail.Puerto = Resources.Puertointerno
        End If
        mail.PassWord = Resources.Password

        Return mail
    End Function

End Class

