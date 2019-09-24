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
    Private _ID As String
    Private _UNION As String
    Private _UNITS As String
    Private _T1ID As String
    Private _T1INFEED As String
    Private _T1INCLINATION As INCLINATION
    Private _T2ID As String
    Private _T2OUTFEED As String
    Private _T2INCLINATION As INCLINATION
    Private _ROTATION As Double

#Region "PROPERTIES"
    Public Property ID As String
        Get
            Return _ID
        End Get
        Set(value As String)
            _ID = value
        End Set
    End Property

    Public Property UNION As String
        Get
            Return _UNION
        End Get
        Set(value As String)
            _UNION = value
        End Set
    End Property

    Public Property UNITS As String
        Get
            Return _UNITS
        End Get
        Set(value As String)
            _UNITS = value
        End Set
    End Property

    Public Property T1ID As String
        Get
            Return _T1ID
        End Get
        Set(value As String)
            _T1ID = value
        End Set
    End Property

    Public Property T1INFEED As String
        Get
            Return _T1INFEED
        End Get
        Set(value As String)
            _T1INFEED = value
        End Set
    End Property

    Public Property T1INCLINATION As INCLINATION
        Get
            Return _T1INCLINATION
        End Get
        Set(value As INCLINATION)
            _T1INCLINATION = value
        End Set
    End Property

    Public Property T2ID As String
        Get
            Return _T2ID
        End Get
        Set(value As String)
            _T2ID = value
        End Set
    End Property

    Public Property T2OUTFEED As String
        Get
            Return _T2OUTFEED
        End Get
        Set(value As String)
            _T2OUTFEED = value
        End Set
    End Property

    Public Property T2INCLINATION As INCLINATION
        Get
            Return _T2INCLINATION
        End Get
        Set(value As INCLINATION)
            _T2INCLINATION = value
        End Set
    End Property
    Public Property ROTATION As Double
        Get
            Return _ROTATION
        End Get
        Set(value As Double)
            If _ROTATION <> value Then
                _ROTATION = value
                Dim oUnion As AcadBlockReference = Eventos.COMDoc().ObjectIdToObject(Me.ID)
                clsA.XPonDato(oUnion, "ROTATION", value)
            End If
        End Set
    End Property
#End Region


    ''' <summary>
    ''' Editar uniones existentes o Nuevas inserciones masivas.
    ''' Solo Id para sacar/poner el resto de los XData del AcadBlockReference
    ''' </summary>
    ''' <param name="id"></param>
    Public Sub New(id As Long)
        Me.ID = id
        PonDatosX()
    End Sub

    ''' <summary>
    ''' Nuevas uniones. Tenemos que pasar todos los parámetros.
    ''' </summary>
    ''' <param name="id"></param>
    ''' <param name="Union"></param>
    ''' <param name="Units"></param>
    ''' <param name="T1Id"></param>
    ''' <param name="T1Infeed"></param>
    ''' <param name="T1Inclination"></param>
    ''' <param name="T2Id"></param>
    ''' <param name="T2Outfeed"></param>
    ''' <param name="T2Inclination"></param>
    ''' <param name="Rotation"></param>
    Public Sub New(id As Long, Union As String, Units As String,
                   T1Id As String, T1Infeed As String, T1Inclination As String,
                    T2Id As String, T2Outfeed As String, T2Inclination As String,
                    Rotation As String)
        Me.ID = id
        PonDatosX(Union, Units, T1Id, T1Id, T1Inclination, T2Id, T2Outfeed, T2Inclination, Rotation)
    End Sub

    Public Sub PonDatosX(
                        Optional Union As String = "",
                        Optional Units As String = "",
                        Optional T1Id As String = "",
                        Optional T1Infeed As String = "",
                        Optional T1Inclination As String = "",
                        Optional T2Id As String = "",
                        Optional T2Outfeed As String = "",
                        Optional T2Inclination As String = "",
                        Optional Rotation As String = "")
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        Dim oUnion As AcadBlockReference = Eventos.COMDoc().ObjectIdToObject(Me.ID)
        'Dim nombres As String() = {"UNION", "UNITS", "T1ID", "T1INFEED", "T1INCLINATION", "T2ID", "T2OUTFEED", "T2INCLINATION", "ROTATION"}
        '
        If Union = "" Then
            Me.UNION = clsA.XLeeDato(oUnion, "UNION")
        Else
            Me.UNION = Union
            clsA.XPonDato(oUnion, "UNION", Union)
        End If
        If Units = "" Then
            Me.UNITS = clsA.XLeeDato(oUnion, "UNITS")
        Else
            Me.UNITS = Units
            clsA.XPonDato(oUnion, "UNITS", Units)
        End If

        If T1Id = "" Then
            Me.T1ID = clsA.XLeeDato(oUnion, "T1ID")
        Else
            Me.T1ID = T1Id
            clsA.XPonDato(oUnion, "T1ID", T1Id)
        End If

        If T1Infeed = "" Then
            Me.T1INFEED = clsA.XLeeDato(oUnion, "T1INFEED")
        Else
            Me.T1INFEED = T1Infeed
            clsA.XPonDato(oUnion, "T1INFEED", T1Infeed)
        End If

        If T1Inclination = "" Then
            Me.T1INCLINATION = clsA.XLeeDato(oUnion, "T1INCLINATION")
        Else
            Me.T1INCLINATION = modTavil.INCLINATION.FLAT 'T1Inclination
            clsA.XPonDato(oUnion, "T1INCLINATION", T1Inclination)
        End If

        If T2Id = "" Then
            Me.T2ID = clsA.XLeeDato(oUnion, "T2ID")
        Else
            Me.T2ID = T2Id
            clsA.XPonDato(oUnion, "T2ID", T2Id)
        End If

        If T2Outfeed = "" Then
            Me.T2OUTFEED = clsA.XLeeDato(oUnion, "T2OUTFEED")
        Else
            Me.T2OUTFEED = T2Outfeed
            clsA.XPonDato(oUnion, "T2OUTFEED", T2Outfeed)
        End If

        If T2Inclination = "" Then
            Me.T2INCLINATION = clsA.XLeeDato(oUnion, "T2INCLINATION")
        Else
            Me.T2INCLINATION = modTavil.INCLINATION.FLAT 'T2Inclination
            clsA.XPonDato(oUnion, "T2INCLINATION", T2Inclination)
        End If

        If Rotation = "" Then
            Me.ROTATION = clsA.XLeeDato(oUnion, "ROTATION")
        Else
            Me.ROTATION = IIf(IsNumeric(Rotation), CInt(Rotation), 0)
            clsA.XPonDato(oUnion, "ROTATION", Rotation)
        End If
    End Sub
End Class

