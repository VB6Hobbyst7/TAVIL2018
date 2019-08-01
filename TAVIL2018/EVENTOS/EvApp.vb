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
        Private Sub EvApp_AppActivate() Handles EvApp.AppActivate

        End Sub

        Private Sub EvApp_EndCommand(CommandName As String) Handles EvApp.EndCommand
            'If CommandName = "OPEN" And oApp.Documents.Count > 0 Then
            'modTavil.AcadBlockReference_PonEventosModified()
            'End If
            If EvApp.Documents.Count = 0 Then Exit Sub
            'If oDoc IsNot Nothing AndAlso oDoc.Equals(oApp.ActiveDocument) = False Then oDoc = oApp.ActiveDocument
            'If CommandName.ToUpper.Contains("COPY") Then
            '    'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            'End If
        End Sub

        Private Sub EvApp_EndOpen(FileName As String) Handles EvApp.EndOpen
            'While oApp.GetAcadState.IsQuiescent = True
            '    System.Windows.Forms.Application.DoEvents()
            'End While
            ' modTavil.AcadBlockReference_PonEventosModified()
            EvDoc = EvApp.ActiveDocument
            ''
            'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
            '
        End Sub

    End Class
End Namespace
'
'AppActivate
'*** Se activa justo antes de que se active la ventana principal de la aplicación. 

'AppDeactivate
'*** Se activa justo antes de que se desactive la ventana principal de la aplicación. 

'ARXLoaded
'*** Se activa cuando se carga una aplicación ObjectARX. 

'ARXUnloaded
'*** Se activa cuando una aplicación ObjectARX se ha descargado. 

'BeginCommand
'*** Se activa inmediatamente después de que se emite un comando, pero antes de que se complete. 

'BeginFileDrop
'*** Se activa cuando se suelta un archivo en la ventana principal de la aplicación.

'BeginLISP
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para evaluar una expresión LISP.

'BeginModal
'*** Se activa justo antes de que se muestre un cuadro de diálogo modal. 

'BeginOpen
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para abrir un dibujo existente. 

'BeginPlot
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para imprimir un dibujo.

'BeginQuit
'*** Se activa justo antes de que finalice una sesión de AutoCA. 

'BeginSave
'*** Se activa inmediatamente después de que AutoCAD recibe una solicitud para guardar el dibujo. 

'EndCommand
'*** Se activa inmediatamente después de que se completa un comandos.

'EndLISP
'*** Se activa al finalizar la evaluación de una expresión LISP.

'EndModal
'*** Se activa justo después de que se descarta un cuadro de diálogo modal.

'EndOpen
'*** Se activa inmediatamente después de que AutoCAD termina de abrir un dibujo existente.

'EndPlot
'*** Se activa después de que se haya enviado un documento a la impresora.

'EndSave
'*** Se activa cuando AutoCAD ha terminado de guardar el dibujog.

'LISPCancelled
'*** Se activa cuando se cancela la evaluación de una expresión LISP.

'NewDrawing
'*** Se activa justo antes de que se cree un nuevo dibujo.

'SysVarChanged
'*** Se activa cuando se cambia el valor de una variable del sistema.

'WindowChanged
'*** Se activa cuando hay un cambio en la ventana de la aplicación.

'WindowMovedOrResized
'*** Se activa justo después de que la ventana de la Aplicación se haya movido o cambiado de tamaño. 
