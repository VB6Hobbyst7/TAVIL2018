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

Partial Public Class Eventos
    Public Shared Sub Subscribe_AXApp()
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginCloseAll, AddressOf AXApp_BeginCloseAll
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginCustomizationMode, AddressOf AXApp_BeginCustomizationMode
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginDoubleClick, AddressOf AXApp_BeginDoubleClick
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginQuit, AddressOf AXApp_BeginQuit
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.DisplayingCustomizeDialog, AddressOf AXApp_DisplayingCustomizeDialog
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.DisplayingDraftingSettingsDialog, AddressOf AXApp_DisplayingDraftingSettingsDialog
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.DisplayingOptionDialog, AddressOf AXApp_DisplayingOptionDialog
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.EndCustomizationMode, AddressOf AXApp_EndCustomizationMode
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.EnterModal, AddressOf AXApp_EnterModal
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf AXApp_Idle
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.LeaveModal, AddressOf AXApp_LeaveModal
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.PreTranslateMessage, AddressOf AXApp_PreTranslateMessage
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.QuitAborted, AddressOf AXApp_QuitAborted
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.QuitWillStart, AddressOf AXApp_QuitWillStart
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanged, AddressOf AXApp_SystemVariableChanged
        AddHandler Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanging, AddressOf AXApp_SystemVariableChanging
    End Sub
    Public Shared Sub Unsubscribe_AXApp()
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginCloseAll, AddressOf AXApp_BeginCloseAll
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginCustomizationMode, AddressOf AXApp_BeginCustomizationMode
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginDoubleClick, AddressOf AXApp_BeginDoubleClick
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginQuit, AddressOf AXApp_BeginQuit
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.DisplayingCustomizeDialog, AddressOf AXApp_DisplayingCustomizeDialog
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.DisplayingDraftingSettingsDialog, AddressOf AXApp_DisplayingDraftingSettingsDialog
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.DisplayingOptionDialog, AddressOf AXApp_DisplayingOptionDialog
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.EndCustomizationMode, AddressOf AXApp_EndCustomizationMode
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.EnterModal, AddressOf AXApp_EnterModal
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf AXApp_Idle
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.LeaveModal, AddressOf AXApp_LeaveModal
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.PreTranslateMessage, AddressOf AXApp_PreTranslateMessage
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.QuitAborted, AddressOf AXApp_QuitAborted
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.QuitWillStart, AddressOf AXApp_QuitWillStart
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanged, AddressOf AXApp_SystemVariableChanged
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanging, AddressOf AXApp_SystemVariableChanging
    End Sub
    Public Shared Sub AXApp_BeginCloseAll(sender As Object, e As BeginCloseAllEventArgs) ' As BeginCloseAllEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_BeginCloseAll")
        If logeventos Then PonLogEv("AXApp_BeginCloseAll")
    End Sub

    Public Shared Sub AXApp_BeginCustomizationMode(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_BeginCustomizationMode")
        If logeventos Then PonLogEv("AXApp_BeginCustomizationMode")
    End Sub

    Public Shared Sub AXApp_BeginDoubleClick(sender As Object, e As BeginDoubleClickEventArgs) ' e As BeginDoubleClickEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_BeginDoubleClick")
        If logeventos Then PonLogEv("AXApp_BeginDoubleClick")
    End Sub

    Public Shared Sub AXApp_BeginQuit(sender As Object, e As BeginQuitEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_BeginQuit")
        If logeventos Then PonLogEv("AXApp_BeginQuit")
    End Sub

    Public Shared Sub AXApp_DisplayingCustomizeDialog(sender As Object, e As TabbedDialogEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_DisplayingCustomizeDialog")
        If logeventos Then PonLogEv("AXApp_DisplayingCustomizeDialog")
    End Sub

    Public Shared Sub AXApp_DisplayingDraftingSettingsDialog(sender As Object, e As TabbedDialogEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_BeginCloseAll")
        If logeventos Then PonLogEv("AXApp_BeginCloseAll")
    End Sub

    Public Shared Sub AXApp_DisplayingOptionDialog(sender As Object, e As TabbedDialogEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_DisplayingOptionDialog")
        If logeventos Then PonLogEv("AXApp_DisplayingOptionDialog")
    End Sub

    Public Shared Sub AXApp_EndCustomizationMode(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_EndCustomizationMode")
        If logeventos Then PonLogEv("AXApp_EndCustomizationMode")
    End Sub

    Public Shared Sub AXApp_EnterModal(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_EnterModal")
        If logeventos Then PonLogEv("AXApp_EnterModal")
    End Sub

    Public Shared Sub AXApp_Idle(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_Idle")
        'If logeventos Then PonLogEv("AXApp_Idle")
        If (app_procesointerno = True) Then Exit Sub
        If coneventos = False Then Exit Sub  ' Para que no haga nada después de un comando.
        ' Actualizar campos al guardar, imprimir, eTransmit, etc... Todos.
        Try
            If AXApp.GetSystemVariable("FIELDEVAL") <> 31 Then AXApp.SetSystemVariable("FIELDEVAL", 31)
        Catch ex As System.Exception
        End Try
        '
        ' Si no hay colIds, salir. No se han borrado uniones
        'If colIds Is Nothing OrElse colIds.Count = 0 Then Exit Sub
        '*********************************************
        '    For Each oId As ObjectId In lIds
        '        Dim oObj As DBObject = clsA.DBObject_Get(oId)
        '        If oObj.GetType.Name = "Circle" Then
        '            EvAXEditor.WriteMessage(vbLf & "Radius: " & CType(oObj, Circle).Radius)
        '        ElseIf oObj.GetType.Name = "BlockReference" Then
        '            EvAXEditor.WriteMessage(vbLf & "BlockName: " & CType(oObj, BlockReference).BlockName)
        '        End If
        '    Next
        '    lIds.Clear()
        '*********************************************

        'If (app_procesointerno = False) Then
        'app_procesointerno = True
        '    If colIds IsNot Nothing AndAlso colIds.Count > 0 Then
        '        Using Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        '            For Each oId As Long In colIds
        '                Dim oBlR As AcadBlockReference = Ev.EvApp.ActiveDocument.ObjectIdToObject(oId)
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
        '                '    'Ev.EvApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        '                'End If
        '            Next
        '            colIds.Clear()
        '            colIds = Nothing
        '        End Using
        'End If
        '    If colHan IsNot Nothing AndAlso colHan.Count > 0 Then
        '        For Each queHandle As String In colHan
        '            Dim oBlr As AcadBlockReference = Ev.EvApp.ActiveDocument.HandleToObject(queHandle)
        '            Call clsA.SeleccionaDameBloqueUno(oBlr.Name, oBlr.Layer)
        '            Ev.EvApp.ActiveDocument.SendCommand("_UPDATEFIELD _all  ")
        '            'clsA.SeleccionaPorHandle(Ev.EvApp.ActiveDocument, Ev.EvApp.ActiveDocument.HandleToObject(queHandle), "_UPDATEFIELD")
        '        Next
        '        colHan.Clear()
        '        colHan = Nothing
        '    End If

        ' AutoEnumera_AppIdle(sender, e)
        'app_procesointerno = False
        'End If

        'Angel. 27/09/19
        ReseteaXDataIncorrectos()
        ActualizaProxyIncorrectos()
        'Angel. 27/09/19
        app_procesointerno = False
    End Sub

    Public Shared Sub AXApp_LeaveModal(sender As Object, e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_LeaveModal")
        If logeventos Then PonLogEv("AXApp_LeaveModal")
    End Sub

    Public Shared Sub AXApp_PreTranslateMessage(sender As Object, e As PreTranslateMessageEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_PreTranslateMessage")
        'If logeventos Then PonLogEv("AXApp_PreTranslateMessage")
    End Sub

    Public Shared Sub AXApp_QuitAborted(sender As Object, e As EventArgs) ' e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_QuitAborted")
        If logeventos Then PonLogEv("AXApp_QuitAborted")
    End Sub

    Public Shared Sub AXApp_QuitWillStart(sender As Object, e As EventArgs) ' AcadApplication. e As EventArgs) ' e As EventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_QuitWillStart")
        If logeventos Then PonLogEv("AXApp_QuitWillStart")
    End Sub

    Public Shared Sub AXApp_SystemVariableChanged(sender As Object, e As Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_SystemVariableChanged")
        If logeventos Then PonLogEv("AXApp_SystemVariableChanged")
    End Sub

    Public Shared Sub AXApp_SystemVariableChanging(sender As Object, e As Autodesk.AutoCAD.ApplicationServices.SystemVariableChangingEventArgs)
        'AXDoc.Editor.WriteMessage("AXApp_SystemVariableChanging")
        If logeventos Then PonLogEv("AXApp_SystemVariableChanging")
    End Sub
End Class
'
'BeginCloseAll                      Se activa justo antes de cerrar todos los documentos.
'BeginCustomizationMode:            Se activa justo antes de que AutoCAD entre en modo de personalización.
'BeginDoubleClick:                  Se activa cuando se hace doble clic en el botón del mouse. 
'BeginQuit:                         Se activa justo antes de que finalice una sesión de AutoCAD.
'DisplayingCustomizeDialog:         Se activa justo antes de que se muestre el cuadro de diálogo Personalizar.
'DisplayingDraftingSettingsDialog:  Se activa justo antes de que se muestre el cuadro de diálogo Configuración de dibujo. 
'DisplayingOptionDialog:            Se activa justo antes de que se muestre el cuadro de diálogo Opciones. 
'EndCustomizationMode:              Se activa cuando AutoCAD sale del modo de personalización. 
'EnterModal:                        Se activa justo antes de que se muestre un cuadro de diálogo modal.
'Idle:                              Se activa cuando el texto de AutoCAD. 
'LeaveModal:                        Se activa cuando se cierra un cuadro de diálogo modal. 
'PreTranslateMessage:               Se activa justo antes de que AutoCAD traduzca un mensaje. 
'QuitAborted:                       Se activa cuando se cancela un intento de apagar AutoCAD.
'QuitWillStart:                     Se activa después del evento BeginQuit y antes de que comience el apagado. 
'SystemVariableChanged:             Se activa cuando se realiza un intento de cambiar una variable del sistema. 
'SystemVariableChanging:            Se activa justo antes de que se intente cambiar una variable del sistema. 

