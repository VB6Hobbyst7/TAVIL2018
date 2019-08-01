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
        Public Sub Application_Idle(ByVal sender As Object, ByVal e As EventArgs)
            ' Si no hay colIds, salir. No se han borrado uniones
            If colIds Is Nothing OrElse colIds.Count = 0 Then Exit Sub
            ' Se ha borrado algún objeto. Comprobar si era una unión

            'RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Application_Idle
            'If oApp Is Nothing Then Exit Sub
            'Try
            '    If oApp.ActiveDocument Is Nothing Then Exit Sub
            'Catch ex As Exception
            '    Exit Sub
            'End Try
            'If oApp.Documents.Count = 0 Then Exit Sub
            'If docAct = Nothing Then Exit Sub

            '' Actualizar campos al guardar, imprimir, eTransmit, etc... Todos.
            'If oApp.ActiveDocument.GetVariable("FIELDEVAL") <> 31 Then oApp.ActiveDocument.SetVariable("FIELDEVAL", 31)
            'If (app_procesointerno = False) Then
            '    app_procesointerno = True
            '    If colIds IsNot Nothing AndAlso colIds.Count > 0 Then
            '        Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
            '            For Each oId As Long In colIds
            '                Dim oBlR As AcadBlockReference = oApp.ActiveDocument.ObjectIdToObject(oId)
            '                Dim oPars As Hashtable = clsA.BloqueDinamicoParametrosDameTodos(oBlR.ObjectID)
            '                Dim largo As Double = Nothing
            '                Dim ancho As Double = Nothing
            '                For Each nPar As String In oPars.Keys
            '                    If nPar = "DISTANCIA1" Or nPar = "DISTANCE1" Or nPar = "LENGTH" Then
            '                        largo = CDbl(oPars(nPar))
            '                    ElseIf nPar = "DISTANCIA2" Or nPar = "DISTANCE2" Or nPar = "WIDTH" Then
            '                        ancho = CDbl(oPars(nPar))
            '                    End If
            '                Next
            '                ' Poner el atributo "REFERENCE"
            '                'Dim queRef As String = clsA.BloqueAtributoDame(oBlR.ObjectID, "REFERENCE")
            '                'If queRef <> "" Then
            '                '    Dim prefijo As String = queRef.Substring(0, 3)
            '                '    clsA.BloqueAtributoEscribe(oBlR.ObjectID, "REFERENCE", prefijo & ancho.ToString & "_" & largo.ToString)
            '                'End If

            '                'Dim arrPatas As ArrayList = clsA.SeleccionaDameBloquesTODOStexto(regAPPA,, PatasCapa, "*" & oBlR.EffectiveName & "*")
            '                'Dim arrPatas As ArrayList = clsA.SeleccionaDameBloquesTODOS(regAPPCliente,, PatasCapa)
            '                'If arrPatas IsNot Nothing AndAlso arrPatas.Count > 0 Then
            '                '    For Each oBlP As AcadBlockReference In arrPatas
            '                '        Dim queCinta As String = clsA.XLeeDato(oBlP, "cinta")
            '                '        If queCinta = oBlR.Handle.ToString Then
            '                '            clsA.BloqueDinamico_ParametroEscribe(oBlP.ObjectID, "WIDTH", ancho) ' ancho.ToString)
            '                '        End If
            '                '    Next
            '                '    'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            '                'End If
            '            Next
            '            colIds.Clear()
            '            colIds = Nothing
            '        End Using
            '    End If

            '    If colHan IsNot Nothing AndAlso colHan.Count > 0 Then
            '        For Each queHandle As String In colHan
            '            Dim oBlr As AcadBlockReference = oApp.ActiveDocument.HandleToObject(queHandle)
            '            Call clsA.SeleccionaDameBloqueUno(oBlr.Name, oBlr.Layer)
            '            oApp.ActiveDocument.SendCommand("_UPDATEFIELD _all  ")
            '            'clsA.SeleccionaPorHandle(oApp.ActiveDocument, oApp.ActiveDocument.HandleToObject(queHandle), "_UPDATEFIELD")
            '        Next
            '        colHan.Clear()
            '        colHan = Nothing
            '    End If

            '    AutoEnumera_AppIdle(sender, e)
            'app_procesointerno = False
            'End If
        End Sub
    End Class
End Namespace
'
'BeginCustomizationMode 
'*** Se activa justo antes de que AutoCAD entre en modo de personalización.

'BeginDoubleClick
'*** Se activa cuando se hace doble clic en el botón del mouse. 

'BeginQuit
'*** Se activa justo antes de que finalice una sesión de AutoCAD.

'DisplayingCustomizeDialog
'*** Se activa justo antes de que se muestre el cuadro de diálogo Personalizar.

'DisplayingDraftingSettingsDialog
'*** Se activa justo antes de que se muestre el cuadro de diálogo Configuración de dibujo. 

'DisplayingOptionDialog
'*** Se activa justo antes de que se muestre el cuadro de diálogo Opciones. 

'EndCustomizationMode
'*** Se activa cuando AutoCAD sale del modo de personalización. 

'EnterModal
'*** Se activa justo antes de que se muestre un cuadro de diálogo modal.

'Idle
'*** Se activa cuando el texto de AutoCAD. 

'LeaveModal
'*** Se activa cuando se cierra un cuadro de diálogo modal. 

'PreTranslateMessage
'*** Se activa justo antes de que AutoCAD traduzca un mensaje. 

'QuitAborted
'*** Se activa cuando se cancela un intento de apagar AutoCAD.

'QuitWillStart
'*** Se activa después del evento BeginQuit y antes de que comience el apagado. 

'SystemVariableChanged
'*** Se activa cuando se realiza un intento de cambiar una variable del sistema. 

'SystemVariableChanging
'*** Se activa justo antes de que se intente cambiar una variable del sistema. 

