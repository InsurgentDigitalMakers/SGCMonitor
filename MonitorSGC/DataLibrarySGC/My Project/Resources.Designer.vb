﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Este código fue generado por una herramienta.
'     Versión de runtime:4.0.30319.42000
'
'     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
'     se vuelve a generar el código.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'StronglyTypedResourceBuilder generó automáticamente esta clase
    'a través de una herramienta como ResGen o Visual Studio.
    'Para agregar o quitar un miembro, edite el archivo .ResX y, a continuación, vuelva a ejecutar ResGen
    'con la opción /str o recompile su proyecto de VS.
    '''<summary>
    '''  Clase de recurso fuertemente tipado, para buscar cadenas traducidas, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "16.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Devuelve la instancia de ResourceManager almacenada en caché utilizada por esta clase.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("DataLibrarySGC.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Reemplaza la propiedad CurrentUICulture del subproceso actual para todas las
        '''  búsquedas de recursos mediante esta clase de recurso fuertemente tipado.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a 0.
        '''</summary>
        Friend ReadOnly Property correocliente() As String
            Get
                Return ResourceManager.GetString("correocliente", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a SGC@aam.com.
        '''</summary>
        Friend ReadOnly Property correosfromcliente() As String
            Get
                Return ResourceManager.GetString("correosfromcliente", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a novocast@idmdev.tech.
        '''</summary>
        Friend ReadOnly Property correosfrominterno() As String
            Get
                Return ResourceManager.GetString("correosfrominterno", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a 10.0.127.29.
        '''</summary>
        Friend ReadOnly Property Hostcliente() As String
            Get
                Return ResourceManager.GetString("Hostcliente", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a smtp.idmdev.tech.
        '''</summary>
        Friend ReadOnly Property HostInterno() As String
            Get
                Return ResourceManager.GetString("HostInterno", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a xq5kyMcW8.
        '''</summary>
        Friend ReadOnly Property Password() As String
            Get
                Return ResourceManager.GetString("Password", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a 25.
        '''</summary>
        Friend ReadOnly Property Puertocliente() As String
            Get
                Return ResourceManager.GetString("Puertocliente", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Busca una cadena traducida similar a 587.
        '''</summary>
        Friend ReadOnly Property Puertointerno() As String
            Get
                Return ResourceManager.GetString("Puertointerno", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
