Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports a2 = AutoCAD2acad.A2acad

Public Module modTavil
    Public Sub AcadBlockReference_PonEventosModified()
        If clsA Is Nothing Then clsA = New a2.A2acad(oApp, cfg._appFullPath, regAPPCliente)
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
    '
    Public Sub Zoom_Seleccion()
        Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
        oApp.ZoomScaled(0.6, AcZoomScaleType.acZoomScaledRelative)
    End Sub
End Module