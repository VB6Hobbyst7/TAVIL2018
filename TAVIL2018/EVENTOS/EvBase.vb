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
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows
Imports AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application
Imports AcadDocument = Autodesk.AutoCAD.ApplicationServices.Document
Imports AcadWindows = Autodesk.AutoCAD.Windows
Imports a2 = AutoCAD2acad.A2acad
'
Namespace Eventos
    Partial Public Class AutoCADEventos
        ' ***** OBJETOS AUTOCAD
        Public WithEvents EvApp As Autodesk.AutoCAD.Interop.AcadApplication = Nothing
        Public WithEvents EvDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing

        Public Sub New(queApp As Autodesk.AutoCAD.Interop.AcadApplication)

        End Sub

        Public Sub AcadBlockReference_PonEventosModified()
            If clsA Is Nothing Then clsA = New a2.A2acad(EvApp, cfg._appFullPath, regAPPCliente)
            'Dim AcadBlockReference As ArrayList = clsA.SeleccionaDameBloquesTODOS(regAPPA)
            'For Each oBl As AcadBlockReference In AcadBlockReference
            '    Dim queTipo As String = clsA.XLeeDato(oBl, "tipo")
            '    If queTipo = "cinta" Then
            '        AddHandler oBl.Modified, AddressOf modTavil.AcadBlockReference_Modified
            '    End If
            ''Next
            'oDoc = oApp.ActiveDocument
            'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
        End Sub
    End Class
End Namespace