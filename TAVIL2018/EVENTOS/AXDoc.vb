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
Imports a2 = AutoCAD2acad.A2acad
'
Partial Public Class Eventos
    Public Shared Sub Subscribe_AXDoc()
        AddHandler AXDoc.BeginDocumentClose, AddressOf AXDoc_BeginDocumentClose
        AddHandler AXDoc.BeginDwgOpen, AddressOf AXDoc_BeginDwgOpen
        AddHandler AXDoc.CloseAborted, AddressOf AXDoc_CloseAborted
        AddHandler AXDoc.CloseWillStart, AddressOf AXDoc_CloseWillStart
        AddHandler AXDoc.CommandCancelled, AddressOf AXDoc_CommandCancelled
        AddHandler AXDoc.CommandEnded, AddressOf AXDoc_CommandEnded
        AddHandler AXDoc.CommandFailed, AddressOf AXDoc_CommandFailed
        AddHandler AXDoc.CommandWillStart, AddressOf AXDoc_CommandWillStart
        AddHandler AXDoc.EndDwgOpen, AddressOf AXDoc_EndDwgOpen
        AddHandler AXDoc.ImpliedSelectionChanged, AddressOf AXDoc_ImpliedSelectionChanged
        AddHandler AXDoc.LayoutSwitched, AddressOf AXDoc_LayoutSwitched
        AddHandler AXDoc.LayoutSwitching, AddressOf AXDoc_LayoutSwitching
        AddHandler AXDoc.LispCancelled, AddressOf AXDoc_LispCancelled
        AddHandler AXDoc.LispEnded, AddressOf AXDoc_LispEnded
        AddHandler AXDoc.LispWillStart, AddressOf AXDoc_LispWillStart
        AddHandler AXDoc.UnknownCommand, AddressOf AXDoc_UnknownCommand
        AddHandler AXDoc.ViewChanged, AddressOf AXDoc_ViewChanged
        '
        Subscribe_AXDB()
        Subscribe_COMDoc()
        Subscribe_AXBlockTR()
        Subscribe_AXEditor()
    End Sub

    Public Shared Sub Unsubscribe_AXDoc()
        RemoveHandler AXDoc.BeginDocumentClose, AddressOf AXDoc_BeginDocumentClose
        RemoveHandler AXDoc.BeginDwgOpen, AddressOf AXDoc_BeginDwgOpen
        RemoveHandler AXDoc.CloseAborted, AddressOf AXDoc_CloseAborted
        RemoveHandler AXDoc.CloseWillStart, AddressOf AXDoc_CloseWillStart
        RemoveHandler AXDoc.CommandCancelled, AddressOf AXDoc_CommandCancelled
        RemoveHandler AXDoc.CommandEnded, AddressOf AXDoc_CommandEnded
        RemoveHandler AXDoc.CommandFailed, AddressOf AXDoc_CommandFailed
        RemoveHandler AXDoc.CommandWillStart, AddressOf AXDoc_CommandWillStart
        RemoveHandler AXDoc.EndDwgOpen, AddressOf AXDoc_EndDwgOpen
        RemoveHandler AXDoc.ImpliedSelectionChanged, AddressOf AXDoc_ImpliedSelectionChanged
        RemoveHandler AXDoc.LayoutSwitched, AddressOf AXDoc_LayoutSwitched
        RemoveHandler AXDoc.LayoutSwitching, AddressOf AXDoc_LayoutSwitching
        RemoveHandler AXDoc.LispCancelled, AddressOf AXDoc_LispCancelled
        RemoveHandler AXDoc.LispEnded, AddressOf AXDoc_LispEnded
        RemoveHandler AXDoc.LispWillStart, AddressOf AXDoc_LispWillStart
        RemoveHandler AXDoc.UnknownCommand, AddressOf AXDoc_UnknownCommand
        RemoveHandler AXDoc.ViewChanged, AddressOf AXDoc_ViewChanged
        '
        Unsubscribe_AXDB()
        Unsubscribe_COMDoc()
        Unsubscribe_AXBlockTR()
        Unsubscribe_AXEditor()
    End Sub
    Public Shared Sub AXDoc_BeginDocumentClose(sender As Object, e As DocumentBeginCloseEventArgs)

    End Sub

    Public Shared Sub AXDoc_BeginDwgOpen(sender As Object, e As DrawingOpenEventArgs)

    End Sub

    Public Shared Sub AXDoc_CloseAborted(sender As Object, e As EventArgs)

    End Sub

    Public Shared Sub AXDoc_CloseWillStart(sender As Object, e As EventArgs)

    End Sub

    Public Shared Sub AXDoc_CommandCancelled(sender As Object, e As CommandEventArgs)

    End Sub

    Public Shared Sub AXDoc_CommandEnded(sender As Object, e As CommandEventArgs)
        ' Si no hay elementos añadidos o modificados, salir (lIds.Count = 0)
        If lIds.Count = 0 Then Exit Sub
        '
        For Each oId As ObjectId In lIds
            Dim oObj As DBObject = clsA.DBObject_Get(oId)
            If oObj.GetType.Name = "Circle" Then
                AXEditor.WriteMessage(vbLf & "Radius: " & CType(oObj, Circle).Radius)
            ElseIf oObj.GetType.Name = "BlockReference" Then
                While AXApp.IsQuiescent
                    System.Windows.Forms.Application.DoEvents()
                End While
                AXEditor.WriteMessage(vbLf & "Block Name: " & CType(oObj, BlockReference).Name)
            End If
        Next
        lIds.Clear()
        ultimoObjectId = Nothing
        '
        '*********************************************
    End Sub

    Public Shared Sub AXDoc_CommandFailed(sender As Object, e As CommandEventArgs)

    End Sub

    Public Shared Sub AXDoc_CommandWillStart(sender As Object, e As CommandEventArgs)
        lIds = New List(Of ObjectId)
    End Sub

    Public Shared Sub AXDoc_EndDwgOpen(sender As Object, e As DrawingOpenEventArgs)

    End Sub

    Public Shared Sub AXDoc_ImpliedSelectionChanged(sender As Object, e As EventArgs)

    End Sub

    Public Shared Sub AXDoc_LayoutSwitched(sender As Object, e As LayoutSwitchedEventArgs)

    End Sub

    Public Shared Sub AXDoc_LayoutSwitching(sender As Object, e As LayoutSwitchingEventArgs)

    End Sub

    Public Shared Sub AXDoc_LispCancelled(sender As Object, e As EventArgs)

    End Sub

    Public Shared Sub AXDoc_LispEnded(sender As Object, e As EventArgs)

    End Sub

    Public Shared Sub AXDoc_LispWillStart(sender As Object, e As LispWillStartEventArgs)

    End Sub

    Public Shared Sub AXDoc_UnknownCommand(sender As Object, e As UnknownCommandEventArgs)

    End Sub

    Public Shared Sub AXDoc_ViewChanged(sender As Object, e As EventArgs)

    End Sub
End Class
'
'BeginDocumentClose         Se activa justo después de que se recibe una solicitud Para cerrar un dibujo.
'BeginDwgOpen               Se activa cuando se va a abrir un dibujo.
'CloseAborted               Se activa cuando se cancela un intento de cerrar un dibujo.
'CloseWillStart             Se activa después del evento BeginDocumentClose y antes de cerrar el dibujo comienza.
'CommandCancelled           Se activa cuando un comando se cancela antes de que se complete.
'CommandEnded               Se activa inmediatamente después de que se completa un comando.
'CommandFailed              Se activa cuando un comando no se completa y no se cancela.
'CommandWillStart           Se activa inmediatamente después de que se emite un comando, pero antes de que se complete.
'EndDwgOpen                 Se activa cuando se abre un dibujo.
'ImpliedSelectionChanged    Se activa cuando cambia el conjunto de selección actual de pickfirst.
'LayoutSwitched             Se activa después de que un diseño se esté configurando como actual.
'LispCancelled              Se activa cuando se cancela la evaluación de una expresión LISP.
'LispEnded                  Se activa al finalizar la evaluación de una expresión LISP.
'LispWillStart              Se activa inmediatamente después de que AutoCAD recibe una solicitud Para evaluar una expresión LISP.
'UnknownCommand             Se dispara inmediatamente cuando se ingresa un comando desconocido en el símbolo del sistema.
'ViewChanged                Se activa después de que la vista de un dibujo ha cambiado.

