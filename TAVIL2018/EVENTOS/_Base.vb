Imports System
Imports System.Text
Imports System.Linq
Imports System.Xml
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms
Imports System.IO

Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Geometry
Imports AXApp = Autodesk.AutoCAD.ApplicationServices.Application
Imports AXDoc = Autodesk.AutoCAD.ApplicationServices.Document
Imports AXWin = Autodesk.AutoCAD.Windows
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad


'
''' <summary>
''' Inicializar los eventos que necesitemos (Poner esto al inicio de AddIn. Primera linea que se ejecuta)
''' Eventos.Eventos_Inicializa()
''' </summary>
Partial Public Class Eventos

    ' ***** OBJETOS AUTOCAD
    '
    ' ***** Los tipos de objetos que vamos a controlar con eventos
    Public Shared lTypesAXObj As String() = {"BlockReference", "Circle"}
    Public Shared lTypesCOMObj As String() = {"AcDbCircle", "AcDbBlockReference"}
    Public Shared lIds As List(Of ObjectId)     ' Lista de objetos modificados
    '
    ' ***** Variables
    Public Shared ultimoObjectId As ObjectId = Nothing
    Public Shared ultimoAXDoc As String = ""
    '
    Public Shared Sub Eventos_Inicializa()
        lIds = New List(Of ObjectId)
        'AXEventM = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        If AXDocM.Count > 0 Then
            Subscribe_AXDoc()
        End If
        Subscribe_AXEventM()
        Subscribe_AXApp()
        Subscribe_COMApp()
        Subscribe_AXDocM()
    End Sub
    Public Shared Sub Eventos_Vacia()
        lIds = Nothing
        If AXDocM.Count > 0 Then
            Unsubscribe_AXDoc()
        End If
        Unsubscribe_AXEventM()
        Unsubscribe_AXApp()
        Unsubscribe_COMApp()
        Unsubscribe_AXDocM()
    End Sub

    Public Shared Sub AcadBlockReference_PonEventosModified()
        If clsA Is Nothing Then clsA = New a2.A2acad(COMApp, cfg._appFullPath, regAPPCliente)
        'Dim AcadBlockReference As ArrayList = clsA.SeleccionaDameBloquesTODOS(regAPPA)
        'For Each oBl As AcadBlockReference In AcadBlockReference
        '    Dim queTipo As String = clsA.XLeeDato(oBl, "tipo")
        '    If queTipo = "cinta" Then
        '        AddHandler oBl.Modified, AddressOf modTavil.AcadBlockReference_Modified
        '    End If
        'Next
    End Sub
    '
    Public Shared Function AXEventM() As Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager
        Return Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance
    End Function

    Public Shared Function AXDocM() As Autodesk.AutoCAD.ApplicationServices.DocumentCollection
        Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
    End Function

    Public Shared Function AXDoc() As Autodesk.AutoCAD.ApplicationServices.Document
        Return AXDocM.MdiActiveDocument
    End Function
    Public Shared Function AXDb() As Autodesk.AutoCAD.DatabaseServices.Database
        Return AXDoc.Database
    End Function
    Public Shared Function AXEditor() As Autodesk.AutoCAD.EditorInput.Editor
        Return AXDoc.Editor
    End Function
    Public Shared Function COMApp() As Autodesk.AutoCAD.Interop.AcadApplication
        Return CType(AXApp.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication).Application
    End Function
    Public Shared Function COMDoc() As Autodesk.AutoCAD.Interop.AcadDocument
        'Return COMApp.ActiveDocument
        Return AXDocM.MdiActiveDocument.GetAcadDocument
    End Function
    Public Shared Function COMDb() As Autodesk.AutoCAD.Interop.Common.AcadDatabase
        Return COMDoc.Database
    End Function
End Class