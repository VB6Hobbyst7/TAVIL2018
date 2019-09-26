Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Drawing
Imports System.Windows.Forms
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad
Imports TAVIL2018

Public Class ClsUnion
    Private _HANDLE As String = ""
    Private _UNION As String = ""
    Private _UNITS As String = ""
    Private _T1HANDLE As String = ""
    Private _T1INFEED As String = ""
    Private _T1INCLINATION As String = ""
    Private _T2HANDLE As String = ""
    Private _T2OUTFEED As String = ""
    Private _T2INCLINATION As String = ""
    Private _ROTATION As String = ""
    Private _ExcelFilaUnion As ExcelFila = Nothing

#Region "PROPERTIES"
    Public Property HANDLE As String
        Get
            Return _HANDLE
        End Get
        Set(value As String)
            _HANDLE = value
        End Set
    End Property

    Public Property UNION As String
        Get
            Return _UNION
        End Get
        Set(value As String)
            If value <> _UNION Then
                _UNION = value
                clsA.XPonDato(Me.HANDLE, "UNION", value)
            End If
        End Set
    End Property

    Public Property UNITS As String
        Get
            Return _UNITS
        End Get
        Set(value As String)
            If value <> _UNITS Then
                _UNITS = value
                clsA.XPonDato(Me.HANDLE, "UNITS", value)
            End If
        End Set
    End Property

    Public Property T1HANDLE As String
        Get
            Return _T1HANDLE
        End Get
        Set(value As String)
            If value <> _T1HANDLE Then
                _T1HANDLE = value
                clsA.XPonDato(Me.HANDLE, "T1HANDLE", value)
            End If
        End Set
    End Property

    Public Property T1INFEED As String
        Get
            Return _T1INFEED
        End Get
        Set(value As String)
            If value <> _T1INFEED Then
                _T1INFEED = value
                clsA.XPonDato(Me.HANDLE, "T1INFEED", value)
            End If
        End Set
    End Property

    Public Property T1INCLINATION As String
        Get
            Return _T1INCLINATION
        End Get
        Set(value As String)
            If value <> _T1INCLINATION Then
                _T1INCLINATION = value
                clsA.XPonDato(Me.HANDLE, "T1INCLINATION", value)
            End If
        End Set
    End Property

    Public Property T2HANDLE As String
        Get
            Return _T2HANDLE
        End Get
        Set(value As String)
            If value <> _T2HANDLE Then
                _T2HANDLE = value
                clsA.XPonDato(Me.HANDLE, "T2HANDLE", value)
            End If
        End Set
    End Property

    Public Property T2OUTFEED As String
        Get
            Return _T2OUTFEED
        End Get
        Set(value As String)
            If value <> _T2OUTFEED Then
                _T2OUTFEED = value
                clsA.XPonDato(Me.HANDLE, "T2OUTFEED", value)
            End If
        End Set
    End Property

    Public Property T2INCLINATION As String
        Get
            Return _T2INCLINATION
        End Get
        Set(value As String)
            If value <> _T2INCLINATION Then
                _T2INCLINATION = value
                clsA.XPonDato(Me.HANDLE, "T2INCLINATION", value)
            End If
        End Set
    End Property
    Public Property ROTATION As String
        Get
            Return _ROTATION
        End Get
        Set(value As String)
            If value <> _ROTATION Then
                _ROTATION = value
                clsA.XPonDato(Me.HANDLE, "ROTATION", value)
            End If
        End Set
    End Property

    Public Property ExcelFilaUnion As ExcelFila
        Get
            Return _ExcelFilaUnion
        End Get
        Set(value As ExcelFila)
            _ExcelFilaUnion = value
        End Set
    End Property
#End Region
    ''' <summary>
    ''' Editar uniones existentes o Nuevas inserciones masivas.
    ''' Solo Handle para sacar/poner el resto de los XData del AcadBlockReference
    ''' </summary>
    ''' <param name="handle"></param>
    Public Sub New(handle As String)
        Me.HANDLE = handle
        'PonDatosX()
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        Dim oUnion As AcadObject = Nothing
        Try
            oUnion = Eventos.COMDoc().HandleToObject(Me.HANDLE)
        Catch ex As Exception
            Exit Sub
        End Try
        '
        Me._UNION = clsA.XLeeDato(oUnion.Handle, "UNION")
        Me._UNITS = clsA.XLeeDato(oUnion.Handle, "UNITS")
        Me._T1HANDLE = clsA.XLeeDato(oUnion.Handle, "T1HANDLE")
        Me._T1INFEED = clsA.XLeeDato(oUnion.Handle, "T1INFEED")
        Me._T1INCLINATION = clsA.XLeeDato(oUnion.Handle, "T1INCLINATION")
        Me._T2HANDLE = clsA.XLeeDato(oUnion.Handle, "T2HANDLE")
        Me._T2OUTFEED = clsA.XLeeDato(oUnion.Handle, "T2OUTFEED")
        Me._T2INCLINATION = clsA.XLeeDato(oUnion.Handle, "T2INCLINATION")
        Me._ROTATION = clsA.XLeeDato(oUnion.Handle, "ROTATION")
        oUnion = Nothing
    End Sub

    ''' <summary>
    ''' Nuevas uniones. Tenemos que pasar todos los parámetros.
    ''' </summary>
    ''' <param name="handle"></param>
    ''' <param name="Union"></param>
    ''' <param name="Units"></param>
    ''' <param name="T1Handle"></param>
    ''' <param name="T1Infeed"></param>
    ''' <param name="T1Inclination"></param>
    ''' <param name="T2Handle"></param>
    ''' <param name="T2Outfeed"></param>
    ''' <param name="T2Inclination"></param>
    ''' <param name="Rotation"></param>
    Public Sub New(handle As String, Union As String, Units As String,
                   T1Handle As String, T1Infeed As String, T1Inclination As String,
                    T2Handle As String, T2Outfeed As String, T2Inclination As String,
                    Rotation As String)
        Me.HANDLE = handle
        'PonDatosX(Union, Units, T1Handle, T1Handle, T1Inclination, T2Handle, T2Outfeed, T2Inclination, Rotation)
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        Dim oUnion As AcadObject = Nothing
        Try
            oUnion = Eventos.COMDoc().HandleToObject(Me.HANDLE)
        Catch ex As Exception
            Exit Sub
        End Try
        '
        Me.UNION = Union
        Me.UNITS = Units
        Me.T1HANDLE = T1Handle
        Me.T1INFEED = T1Infeed
        Me.T1INCLINATION = T1Inclination
        Me.T2HANDLE = T2Handle
        Me.T2OUTFEED = T2Outfeed
        Me.T2INCLINATION = T2Inclination
        Me.ROTATION = Rotation
        oUnion = Nothing
    End Sub

    Public Function OUnionDame() As AcadBlockReference
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        Return Eventos.COMDoc().HandleToObject(Me.HANDLE)
    End Function
    Public Sub PonAtributos()
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        Dim oUnion As AcadBlockReference = OUnionDame()
        If oUnion IsNot Nothing Then
            Try
                Dim dicNomVal As New Dictionary(Of String, String)
                dicNomVal.Add("UNION", Me.UNION)
                dicNomVal.Add("UNITS", Me.UNITS)
                clsA.Bloque_AtributoEscribe(oUnion.ObjectID, dicNomVal)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
    End Sub
End Class

