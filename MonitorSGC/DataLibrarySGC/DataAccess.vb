Imports DataLibrarySGC.My
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Public Class DataAccess
    Private Const SPTareasNoFinalizadas = "spMonitorTareaNoFinGet"
    Private Const SPActualizaEstatusTarea = "spActualizaTareaEstatus"
    Private Const SPMonitorTareaInsert = "spMonitorTareaInsert"
    Private Const cadenaConnLocal = "Data Source=DESKTOP-M4JHF48\SQLEXPRESS;Initial Catalog=SistemaCalidad;User ID=sa;Password=becky"
    Private Const cadenaConnProd = "Data Source=192.168.1.129;Initial Catalog=SistemaCalidad;User ID=Admin;Password=123456"
    Private Const isProduccion As Integer = 1
    Private Const FileConexion As String = "c:\configMonitor\conexion.txt"

    Public Function getCadenaConexion() As String
        Dim conexion As String = ""
        Dim objReader As New StreamReader(FileConexion)
        Dim sLine As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not (sLine Is Nothing) Then
                conexion = sLine
            End If
        Loop Until sLine Is Nothing
        objReader.Close()
        Return conexion
    End Function
    ''' <summary>
    ''' Esta Función permite insertar una nueva tarea
    ''' </summary>
    Public Sub MonitorTareaInsert(ByVal tarea As MonitorTarea)
        Try
            Dim ConnString = getCadenaConexion()

            Using conn As New SqlConnection(ConnString)
                Using comm As New SqlCommand(SPMonitorTareaInsert)
                    comm.CommandType = CommandType.StoredProcedure
                    comm.Connection = conn
                    conn.Open()
                    comm.Parameters.AddWithValue("@BuzonTareaId", tarea.BuzonTareaId)
                    comm.Parameters.AddWithValue("@StatusTareaId", tarea.StatusTareaId)
                    comm.Parameters.AddWithValue("@IdVista", tarea.IdVista)
                    comm.Parameters.AddWithValue("@IdEmpleado", tarea.IdEmpleado)
                    comm.Parameters.AddWithValue("@Status", ConvertBooleantoInt(tarea.Status))
                    comm.Parameters.AddWithValue("@FechaCreacion", tarea.FechaCreacion)
                    comm.Parameters.AddWithValue("@FrecuenciaTareaId", tarea.FrecuenciaTareaId)
                    comm.Parameters.AddWithValue("@UnidadTiempoId", tarea.UnidadTiempoId)
                    comm.Parameters.AddWithValue("@PadreId", tarea.PadreId)
                    comm.ExecuteNonQuery()
                End Using
                conn.Close()
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Esta Función actualiza una tarea como terminada
    ''' </summary>
    Public Sub actualizaTareaEstatus(ByVal MonitorTareId As Integer, ByVal EstatusTarea As String)
        Try

            Dim ConnString = getCadenaConexion()

            Using conn As New SqlConnection(ConnString)
                Using comm As New SqlCommand(SPActualizaEstatusTarea)
                    comm.CommandType = CommandType.StoredProcedure
                    comm.Connection = conn
                    conn.Open()
                    comm.Parameters.AddWithValue("@MonitorTareaId", MonitorTareId)
                    comm.Parameters.AddWithValue("@StatusTarea", EstatusTarea)
                    comm.ExecuteNonQuery()
                End Using
                conn.Close()
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    ''' <summary>
    ''' Esta Función regresa el id del jefe de un empleado
    ''' </summary>
    Public Function obtenJefeEmpleado(ByVal EmpleadoId As Integer) As Integer
        Dim query As String = String.Format("SELECT Coalesce(IDJefe,0) IDJefe from Empleado WHERE IDEmpleado={0}", EmpleadoId)
        Dim dt = obtenerDataQuery(query)
        Dim JefeId As Integer
        For Each row As DataRow In dt.Rows
            JefeId = CType(row("IDJefe"), Integer)
        Next
        Return JefeId
    End Function

    ''' <summary>
    ''' Esta Función obtiene el estatus de una tarea por su nombre
    ''' </summary>
    Public Function getEstatusTareaIdByName(ByVal EstatusTareaName As String)
        Dim query As String = String.Format("SELECT StatusTareaId from StatusTarea WHERE Descripcion='{0}'", EstatusTareaName)
        Dim dt = obtenerDataQuery(query)
        Dim EstatusTareaId As Integer
        For Each row As DataRow In dt.Rows
            EstatusTareaId = CType(row("StatusTareaId"), Integer)
        Next
        Return EstatusTareaId
    End Function
    Public Function obtenerDataQuery(ByVal query As String) As DataTable

        Try
            Using cnSql As New SqlConnection(getCadenaConexion)
                Dim Dt As DataTable
                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand
                If Not cnSql.State = ConnectionState.Open Then
                    cnSql.Open()
                End If
                With Cmd
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Connection = cnSql
                End With
                Da.SelectCommand = Cmd
                Dt = New DataTable
                'modelo.Configuration.ProxyCreationEnabled = False
                Da.Fill(Dt)
                If cnSql.State = ConnectionState.Open Then
                    cnSql.Close()
                End If
                Return Dt
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    ''' <summary>
    ''' Esta Función regresa una lista de todas las tareas pendientes
    ''' </summary>
    Public Function obtenerTareaSinFinalizar() As List(Of MonitorTarea)
        Try
            Using cnString As New SqlConnection(getCadenaConexion)
                Dim Dt As New DataTable
                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand(SPTareasNoFinalizadas)
                If Not cnString.State = ConnectionState.Open Then
                    cnString.Open()
                End If
                With Cmd
                    .CommandType = CommandType.StoredProcedure
                    .Connection = cnString
                End With
                Using dReader As SqlDataReader = Cmd.ExecuteReader()
                    Dt.Load(dReader)
                    If cnString.State = ConnectionState.Open Then
                        cnString.Close()
                    End If
                    Return dtToMonitorTareaList(Dt)
                End Using

            End Using

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function dtToMonitorTareaList(ByVal dt As DataTable) As List(Of MonitorTarea)
        Dim lista As New List(Of MonitorTarea)
        For Each row As DataRow In dt.Rows
            Dim monitor As New MonitorTarea
            With monitor
                .MonitorTareaId = CType(row("MonitorTareaId"), Integer)
                .IdEmpleado = CType(row("IdEmpleado"), Integer)
                .IdVista = CType(row("IdVista"), Integer)
                .FechaCreacion = CType(row("FechaCreacion"), DateTime)
                .UnidadTiempoId = CType(row("UnidadTiempoId"), Integer)
                .UnidadTiempo = row("UnidadTiempo").ToString
                .StatusTareaId = CType(row("StatusTareaId"), Integer)
                .Frecuencia = CType(row("Frecuencia"), Integer)
                .FrecuenciaTareaId = CType(row("FrecuenciaTareaId"), Integer)
                .BuzonTareaId = CType(row("BuzonTareaId"), Integer)
                .Status = CType(row("Status").ToString, Boolean)
            End With
            lista.Add(monitor)
        Next
        Return lista
    End Function

    Public Function ConvertBooleantoInt(ByVal value As Boolean) As Integer
        If value Then
            Return 1
        Else
            Return 0
        End If
    End Function
End Class

