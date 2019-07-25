Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports AutoCAD2acad.clsAutoCAD2acad

Module modTavil

    '
    Public Sub AcadBlockReference_PonEventosModified()
        If clsA Is Nothing Then clsA = New AutoCAD2acad.clsAutoCAD2acad(oApp, app_folderandnameExt)
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
End Module
