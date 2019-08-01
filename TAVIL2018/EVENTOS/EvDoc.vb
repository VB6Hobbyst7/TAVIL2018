Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports oAppS = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
'
Namespace Eventos
    Partial Public Class AutoCADEventos

        Private Sub EvDoc_BeginCommand(CommandName As String) Handles EvDoc.BeginCommand
            'If CommandName = "SAVE" Or CommandName = "QSAVE" Then
            'If clsA Is Nothing Then clsA = New AutoCAD2acad.clsAutoCAD2acad(oApp, cfg._appFullPath, regAPPCliente)
            'clsA.SeleccionaTodosObjetos("INSERT",, True)
            'oApp.ActiveDocument.SendCommand("_UPDATEFIELD _ALL  ")

            '
            'Dim arrIdBloques As ArrayList     '' Arraylist con los IDs de los bloques que empiezan por 
            'arrIdBloques = clsA.SeleccionaTodosObjetos("INSERT",, True)
            '
            ' Procesamos todos los bloques (arralist de Ids)
            'Dim arrHighlight As New ArrayList
            'For Each queId As Long In arrIdBloques
            '    Dim oBl As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(queId)
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
            '    oApp.ZoomScaled(0.75, AcZoomScaleType.acZoomScaledRelative)
            'End If
            'End If
        End Sub

        Private Sub EvDoc_ObjectModified(queObj As Object) Handles EvDoc.ObjectModified
            'If (app_procesointerno = False) Then
            '    If TypeOf queObj Is Autodesk.AutoCAD.Interop.Common.AcadBlockReference Then
            '        Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            '            'modTavil.AcadBlockReference_Modified(CType(queObj, Autodesk.AutoCAD.Interop.Common.AcadBlockReference))
            '            Try
            '                Dim queTipo As String = clsA.XLeeDato(CType(queObj, AcadObject), "tipo")
            '                'MsgBox(oBlr.EffectiveName & " Modificado")
            '                If colIds Is Nothing Then colIds = New List(Of Long)
            '                If colHan Is Nothing Then colHan = New List(Of String)
            '                If colIds.Contains(CType(queObj, AcadBlockReference).ObjectID) = False And queTipo = "cinta" Then
            '                    colIds.Add(CType(queObj, AcadBlockReference).ObjectID)
            '                    'colHan.Add(CType(queObj, AcadBlockReference).Handle)
            '                    ' Si vamos a modificar algo, poner app_procesointer = true (Para que no active eventos)
            '                    app_procesointerno = True
            '                    'clsA.SeleccionaPorHandle(oApp.ActiveDocument, queObj, "_UPDATEFIELD")
            '                End If
            '            Catch ex As Exception
            '                Debug.Print(ex.ToString)
            '            End Try
            '        End Using
            '    End If
            '    ' Volver a false para activar los eventos.
            '    app_procesointerno = False
            '    'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            'End If
        End Sub
    End Class
End Namespace
'
'Activate
'*** Se activa cuando se activa una ventana de documento.

'BeginClose
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para cerrar un dibujo. 

'BeginDocClose
'*** Se activa justo después de recibir una solicitud para cerrar un dibujo.

'BeginCommand
'*** Se activa inmediatamente después de que se emite un comando, pero antes de que se complete. 

'BeginDoubleClick
'*** Se activa después de que el usuario haga doble clic en un objeto del dibujo. 

'BeginLISP
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para evaluar una expresión LISP.

'BeginPlot
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para imprimir un dibujo.

'BeginRightClick
'*** Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo.

'BeginSave
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para guardar el dibujo.

'BeginShortcutMenuCommand
'*** Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Comando..

'BeginShortcutMenuDefault
'*** Se activa después de que el usuario haga clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Predeterminado. 

'BeginShortcutMenuEdit
' *** Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Editar.

'BeginShortcutMenuGrip
'*** Se activa después de que el usuario haga clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en el modo Agarre.

'BeginShortcutMenuOsnap
'*** Se activa después de que el usuario hace clic con el botón derecho en la ventana Dibujo y antes de que aparezca el menú contextual en modo Osnap.

'Deactivate
'*** Se activa cuando la ventana de dibujo está desactivada.

'EndCommand
'*** Se activa inmediatamente después de que se completa un comando. 

'EndLISP
'*** Se activa al finalizar la evaluación de una expresión LISP.

'EndPlot
'*** Se activa después de que se haya enviado un documento a la impresora.

'EndSave
'*** Se activa cuando AutoCAD ha terminado de guardar el dibujo.

'EndShortcutMenu
'*** Se activa después de que aparece el menú contextual.

'LayoutSwitched
'*** Se activa después de que el usuario cambia a un diseño diferente (presentacion).

'LISPCancelled
'*** Se activa cuando se cancela la evaluación de una expresión LISP.

'ObjectAdded
'*** Se activa cuando se agrega un objeto al dibujo.

'ObjectErased
'*** Se activa cuando un objeto ha sido borrado del dibujo.

'ObjectModified
'*** Se activa cuando un objeto en el dibujo ha sido modificado.

'SelectionChanged
'*** Se activa cuando cambia el conjunto de selección actual de pickfirst.

'WindowChanged
'*** Se activa cuando hay un cambio en la ventana de documento.

'WindowMovedOrResized
'*** Se activa justo después de que la ventana Dibujo se haya movido o cambiado de tamaño.
