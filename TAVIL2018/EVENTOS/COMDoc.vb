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
Imports System.Runtime.InteropServices
'
Partial Public Class Eventos
    Public Shared Sub Subscribe_COMDoc()
        If COMdoc Is Nothing Then Exit Sub
        AddHandler COMDoc.Activate, AddressOf EvDoc_Activate
        AddHandler COMDoc.BeginClose, AddressOf EvDoc_BeginClose
        AddHandler COMDoc.BeginDocClose, AddressOf EvDoc_BeginDocClose
        AddHandler COMDoc.BeginCommand, AddressOf EvDoc_BeginCommand
        AddHandler COMDoc.BeginDoubleClick, AddressOf EvDoc_BeginDoubleClick
        AddHandler COMDoc.BeginCommand, AddressOf EvDoc_BeginCommand
        AddHandler COMDoc.BeginLisp, AddressOf EvDoc_BeginLisp
        AddHandler COMDoc.BeginPlot, AddressOf EvDoc_BeginPlot
        AddHandler COMDoc.BeginRightClick, AddressOf EvDoc_BeginRightClick
        AddHandler COMDoc.BeginSave, AddressOf EvDoc_BeginSave
        AddHandler COMDoc.BeginShortcutMenuCommand, AddressOf EvDoc_BeginShortcutMenuCommand
        AddHandler COMDoc.BeginShortcutMenuDefault, AddressOf EvDoc_BeginShortcutMenuDefault
        AddHandler COMDoc.BeginShortcutMenuEdit, AddressOf EvDoc_BeginShortcutMenuEdit
        AddHandler COMDoc.BeginShortcutMenuGrip, AddressOf EvDoc_BeginShortcutMenuGrip
        AddHandler COMDoc.BeginShortcutMenuOsnap, AddressOf EvDoc_BeginShortcutMenuOsnap
        AddHandler COMDoc.Deactivate, AddressOf EvDoc_Deactivate
        AddHandler COMDoc.EndCommand, AddressOf EvDoc_EndCommand
        AddHandler COMDoc.EndLisp, AddressOf EvDoc_EndLisp
        AddHandler COMDoc.EndPlot, AddressOf EvDoc_EndPlot
        AddHandler COMDoc.EndSave, AddressOf EvDoc_EndSave
        AddHandler COMDoc.EndShortcutMenu, AddressOf EvDoc_EndShortcutMenu
        AddHandler COMDoc.LayoutSwitched, AddressOf EvDoc_LayoutSwitched
        AddHandler COMDoc.LispCancelled, AddressOf EvDoc_LispCancelled
        AddHandler COMDoc.ObjectAdded, AddressOf EvDoc_ObjectAdded
        AddHandler COMDoc.ObjectErased, AddressOf EvDoc_ObjectErased
        AddHandler COMDoc.ObjectModified, AddressOf EvDoc_ObjectModified
        AddHandler COMDoc.SelectionChanged, AddressOf EvDoc_SelectionChanged
        AddHandler COMDoc.WindowChanged, AddressOf EvDoc_WindowChanged
        AddHandler COMDoc.WindowMovedOrResized, AddressOf EvDoc_WindowMovedOrResized
    End Sub

    Public Shared Sub Unsubscribe_COMDoc()
        If COMDoc() Is Nothing Then Exit Sub
        RemoveHandler COMDoc.Activate, AddressOf EvDoc_Activate
        RemoveHandler COMDoc.BeginClose, AddressOf EvDoc_BeginClose
        RemoveHandler COMDoc.BeginDocClose, AddressOf EvDoc_BeginDocClose
        RemoveHandler COMDoc.BeginCommand, AddressOf EvDoc_BeginCommand
        RemoveHandler COMDoc.BeginDoubleClick, AddressOf EvDoc_BeginDoubleClick
        RemoveHandler COMDoc.BeginCommand, AddressOf EvDoc_BeginCommand
        RemoveHandler COMDoc.BeginLisp, AddressOf EvDoc_BeginLisp
        RemoveHandler COMDoc.BeginPlot, AddressOf EvDoc_BeginPlot
        RemoveHandler COMDoc.BeginRightClick, AddressOf EvDoc_BeginRightClick
        RemoveHandler COMDoc.BeginSave, AddressOf EvDoc_BeginSave
        RemoveHandler COMDoc.BeginShortcutMenuCommand, AddressOf EvDoc_BeginShortcutMenuCommand
        RemoveHandler COMDoc.BeginShortcutMenuDefault, AddressOf EvDoc_BeginShortcutMenuDefault
        RemoveHandler COMDoc.BeginShortcutMenuEdit, AddressOf EvDoc_BeginShortcutMenuEdit
        RemoveHandler COMDoc.BeginShortcutMenuGrip, AddressOf EvDoc_BeginShortcutMenuGrip
        RemoveHandler COMDoc.BeginShortcutMenuOsnap, AddressOf EvDoc_BeginShortcutMenuOsnap
        RemoveHandler COMDoc.Deactivate, AddressOf EvDoc_Deactivate
        RemoveHandler COMDoc.EndCommand, AddressOf EvDoc_EndCommand
        RemoveHandler COMDoc.EndLisp, AddressOf EvDoc_EndLisp
        RemoveHandler COMDoc.EndPlot, AddressOf EvDoc_EndPlot
        RemoveHandler COMDoc.EndSave, AddressOf EvDoc_EndSave
        RemoveHandler COMDoc.EndShortcutMenu, AddressOf EvDoc_EndShortcutMenu
        RemoveHandler COMDoc.LayoutSwitched, AddressOf EvDoc_LayoutSwitched
        RemoveHandler COMDoc.LispCancelled, AddressOf EvDoc_LispCancelled
        RemoveHandler COMDoc.ObjectAdded, AddressOf EvDoc_ObjectAdded
        RemoveHandler COMDoc.ObjectErased, AddressOf EvDoc_ObjectErased
        RemoveHandler COMDoc.ObjectModified, AddressOf EvDoc_ObjectModified
        RemoveHandler COMDoc.SelectionChanged, AddressOf EvDoc_SelectionChanged
        RemoveHandler COMDoc.WindowChanged, AddressOf EvDoc_WindowChanged
        RemoveHandler COMDoc.WindowMovedOrResized, AddressOf EvDoc_WindowMovedOrResized
    End Sub
    Public Shared Sub EvDoc_Activate()

    End Sub
    Public Shared Sub EvDoc_BeginClose()

    End Sub

    Public Shared Sub EvDoc_BeginCommand(CommandName As String)
        'If CommandName = "SAVE" Or CommandName = "QSAVE" Then
        'If clsA Is Nothing Then clsA = New AutoCAD2acad.clsAutoCAD2acad(Ev.EvApp, cfg._appFullPath, regAPPCliente)
        'clsA.SeleccionaTodosObjetos("INSERT",, True)
        'Ev.EvApp.ActiveDocument.SendCommand("_UPDATEFIELD _ALL  ")

        '
        'Dim arrIdBloques As ArrayList     '' Arraylist con los IDs de los bloques que empiezan por 
        'arrIdBloques = clsA.SeleccionaTodosObjetos("INSERT",, True)
        '
        ' Procesamos todos los bloques (arralist de Ids)
        'Dim arrHighlight As New ArrayList
        'For Each queId As Long In arrIdBloques
        '    Dim oBl As AcadBlockReference = Ev.EvApp.ActiveDocument.ObjectIdToObject(queId)
        '    ''
        '    Dim capa As String = oBl.Layer
        '    Dim partes() As String = capa.Split("·"c)
        '    Dim nivel As String = modAutoCAD.BloqueAtributoDame(queId, "NIVEL")
        '    If partes.Length > 1 AndAlso (partes(1) = ultimoNivel Or nivel = ultimoNivel) Then
        '        arrHighlight.Add(CType(oBl, AcadEntity))
        '    End If
        'Next
        ''
        'If arrHighlight IsNot Nothing AndAlso arrHighlight.Count > 0 Then
        '    modAutoCAD.SeleccionCreaResalta(queEntidades:=arrHighlight, tiempo:=5000, conZoom:=True)
        '    'Autodesk.AutoCAD.Internal.Utils.SelectObjects(ids:=Autodesk.AutoCAD.DatabaseServices.ObjectId)
        '    'Autodesk.AutoCAD.Internal.Utils.ZoomObjects(False)
        '    Ev.EvApp.ZoomScaled(0.75, AcZoomScaleType.acZoomScaledRelative)
        'End If
        'End If
    End Sub

    Public Shared Sub EvDoc_BeginDocClose(ByRef Cancel As Boolean)

    End Sub

    Public Shared Sub EvDoc_BeginDoubleClick(PickPoint As Object)

    End Sub

    Public Shared Sub EvDoc_BeginLisp(FirstLine As String)

    End Sub

    Public Shared Sub EvDoc_BeginPlot(DrawingName As String)

    End Sub

    Public Shared Sub EvDoc_BeginRightClick(PickPoint As Object)

    End Sub

    Public Shared Sub EvDoc_BeginSave(FileName As String)

    End Sub

    Public Shared Sub EvDoc_BeginShortcutMenuCommand(ByRef ShortcutMenu As AcadPopupMenu, Command As String)

    End Sub

    Public Shared Sub EvDoc_BeginShortcutMenuDefault(ByRef ShortcutMenu As AcadPopupMenu)

    End Sub

    Public Shared Sub EvDoc_BeginShortcutMenuEdit(ByRef ShortcutMenu As AcadPopupMenu, ByRef SelectionSet As AcadSelectionSet)

    End Sub

    Public Shared Sub EvDoc_BeginShortcutMenuGrip(ByRef ShortcutMenu As AcadPopupMenu)

    End Sub

    Public Shared Sub EvDoc_BeginShortcutMenuOsnap(ByRef ShortcutMenu As AcadPopupMenu)

    End Sub

    Public Shared Sub EvDoc_Deactivate()

    End Sub

    Public Shared Sub EvDoc_EndCommand(CommandName As String)
        'If CommandName = "INSERT" Then

        '    'finComando = True
        'End If
    End Sub

    Public Shared Sub EvDoc_EndLisp()

    End Sub

    Public Shared Sub EvDoc_EndPlot(DrawingName As String)

    End Sub

    Public Shared Sub EvDoc_EndSave(FileName As String)

    End Sub

    Public Shared Sub EvDoc_EndShortcutMenu(ByRef ShortcutMenu As AcadPopupMenu)

    End Sub

    Public Shared Sub EvDoc_LayoutSwitched(LayoutName As String)

    End Sub

    Public Shared Sub EvDoc_LispCancelled()

    End Sub

    Public Shared Sub EvDoc_ObjectAdded([Object] As Object)
        'If clsA Is Nothing Then clsA = New AutoCAD2acad.A2acad.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        ''Debug.Print([Object].EntityName)
        'Dim acadO As AcadObject = EvCOMApp.ActiveDocument.ObjectIdToObject([Object].ObjectId)
        'Dim oId As ObjectId = New ObjectId(acadO.ObjectID)
        'Dim acadDb As DBObject = clsA.DBObject_Get(oId)
        'Debug.Print(acadO.ObjectName)
        ''
        'If lTypesAXObj.Contains(acadDb.GetType.Name) AndAlso
        '    lHasCode.Contains(acadDb.AcadObject.OBjectId.ToString) = False Then
        '    Subscribre_EvCOMObj(acadO)
        '    Subscribe_EvAXObj(acadDb)
        'End If
    End Sub

    Public Shared Sub EvDoc_ObjectErased(<ComAliasName("AXDBLib.LONG_PTR")> ObjectId As Long)
        'Try
        '    Unsubscribe_EvAXObj(clsA.DBObject_Get(ObjectId))
        '    Dim colIds As New ObjectIdCollection
        '    colIds.Add(New Autodesk.AutoCAD.DatabaseServices.ObjectId(New IntPtr(ObjectId)))
        '    EvAXDocM.CurrentDocument.Database.Purge(colIds)
        'Catch ex As System.Exception

        'End Try
    End Sub

    Public Shared Sub EvDoc_ObjectModified(queObj As Object)
        'If TypeOf queObj Is Autodesk.AutoCAD.Interop.Common.AcadBlockReference Then
        '    '        Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '    '            'modTavil.AcadBlockReference_Modified(CType(queObj, Autodesk.AutoCAD.Interop.Common.AcadBlockReference))
        '    '            Try
        '    '                Dim queTipo As String = clsA.XLeeDato(CType(queObj, AcadObject), "tipo")
        '    '                'MsgBox(oBlr.EffectiveName & " Modificado")
        '    '                If colIds Is Nothing Then colIds = New List(Of Long)
        '    '                If colHan Is Nothing Then colHan = New List(Of String)
        '    '                If colIds.Contains(CType(queObj, AcadBlockReference).ObjectID) = False And queTipo = "cinta" Then
        '    '                    colIds.Add(CType(queObj, AcadBlockReference).ObjectID)
        '    '                    'colHan.Add(CType(queObj, AcadBlockReference).Handle)
        '    '                    ' Si vamos a modificar algo, poner app_procesointer = true (Para que no active eventos)
        '    '                    app_procesointerno = True
        '    '                    'clsA.SeleccionaPorHandle(Ev.EvApp.ActiveDocument, queObj, "_UPDATEFIELD")
        '    '                End If
        '    '            Catch ex As Exception
        '    '                Debug.Print(ex.ToString)
        '    '            End Try
        '    '        End Using
        '    '    End If
        '    '    ' Volver a false para activar los eventos.
        '    '    app_procesointerno = False
        '    '    'Ev.EvApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        'ElseIf TypeOf queObj Is Autodesk.AutoCAD.Interop.Common.AcadCircle Then
        '    Debug.Print(CType(queObj, Autodesk.AutoCAD.Interop.Common.AcadCircle).Diameter)
        'End If
    End Sub

    Public Shared Sub EvDoc_SelectionChanged()

    End Sub

    Public Shared Sub EvDoc_WindowChanged(WindowState As AcWindowState)

    End Sub

    Public Shared Sub EvDoc_WindowMovedOrResized(<ComAliasName("AXDBLib.LONG_PTR")> HWNDFrame As Long, bMoved As Boolean)

    End Sub
End Class
'
'Activate:                  Se activa cuando se activa una ventana de documento.
'BeginClose:                Se activa inmediatamente después de que AutoCAD recibe una solicitud para cerrar un dibujo. 
'BeginDocClose:             Se activa justo después de recibir una solicitud para cerrar un dibujo.
'BeginCommand:              Se activa inmediatamente después de que se emite un comando, pero antes de que se complete. 
'BeginDoubleClick:          Se activa después de que el usuario haga doble clic en un objeto del dibujo. 
'BeginLISP:                 Se activa inmediatamente después de que AutoCAD recibe una solicitud para evaluar una expresión LISP.
'BeginPlot:                 Se activa inmediatamente después de que AutoCAD recibe una solicitud para imprimir un dibujo.
'BeginRightClick:           Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo.
'BeginSave:                 Se activa inmediatamente después de que AutoCAD recibe una solicitud para guardar el dibujo.
'BeginShortcutMenuCommand:  Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Comando..
'BeginShortcutMenuDefault:  Se activa después de que el usuario haga clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Predeterminado. 
'BeginShortcutMenuEdit:     Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Editar.
'BeginShortcutMenuGrip:     Se activa después de que el usuario haga clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Agarre.
'BeginShortcutMenuOsnap:    Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en modo Osnap.
'Deactivate:                Se activa cuando la ventana de dibujo está desactivada.
'EndCommand:                Se activa inmediatamente después de que se completa un comando. 
'EndLISP:                   Se activa al finalizar la evaluación de una expresión LISP.
'EndPlot:                   Se activa después de que se haya enviado un documento a la impresora.
'EndSave:                   Se activa cuando AutoCAD ha terminado de guardar el dibujo.
'EndShortcutMenu:           Se activa después de que aparece el menú contextual.
'LayoutSwitched:            Se activa después de que el usuario cambia a un diseño diferente (presentacion).
'LISPCancelled:             Se activa cuando se cancela la evaluación de una expresión LISP.
'ObjectAdded:               Se activa cuando se agrega un objeto al dibujo.
'ObjectErased:              Se activa cuando un objeto ha sido borrado del dibujo.
'ObjectModified:            Se activa cuando un objeto en el dibujo ha sido modificado.
'SelectionChanged:          Se activa cuando cambia el conjunto de selección actual de pickfirst.
'WindowChanged:             Se activa cuando hay un cambio en la ventana de documento.
'WindowMovedOrResized:      Se activa justo después de que la ventana Dibujo se haya movido o cambiado de tamaño.
