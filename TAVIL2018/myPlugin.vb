' (C) Copyright 2018 by 
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports uau = UtilesAlberto.Utiles


' This line is not mandatory, but improves loading performances
<Assembly: ExtensionApplication(GetType(TAVIL2018.MyPlugin))>

Namespace TAVIL2018

    ' This class is instantiated by AutoCAD once and kept alive for the 
    ' duration of the session. If you don't do any one time initialization 
    ' then you should remove this class.
    Public Class MyPlugin
        Implements IExtensionApplication

        Public Sub New()
            Me.Initialize()
        End Sub

        'Public Shared docAct As Document
        'Public Shared Event DBObjectErased As ObjectErasedEventHandler 'Cuando un objeto es eliminado de la base de datos.
        'Public Shared Event DBObjectModified As ObjectEventHandler 'Cuando un objeto es modificado en la base de datos
        'Public Shared Event DBObjectAppended As ObjectEventHandler 'Cuando un objeto es añadido a la base de datos
        'Public Shared Event AppIdle As EventHandler 'Cuando Autocad no esta ejecutando nada.

        '
        'Private Sub Database_ObjectErased(ByVal sender As Object, ByVal e As ObjectErasedEventArgs)
        '    RaiseEvent DBObjectErased(sender, e)
        'End Sub
        'Private Sub Database_ObjectModified(ByVal sender As Object, ByVal e As ObjectEventArgs)
        '    RaiseEvent DBObjectModified(sender, e)
        'End Sub
        'Private Sub Database_ObjectAppended(ByVal sender As Object, ByVal e As ObjectEventArgs)
        '    RaiseEvent DBObjectAppended(sender, e)
        'End Sub

        'Private Sub Application_Idle(ByVal sender As Object, ByVal e As EventArgs)
        '    RaiseEvent AppIdle(sender, e)
        'End Sub


        Public Sub Initialize() Implements IExtensionApplication.Initialize
            ' Add one time initialization here
            ' One common scenario is to setup a callback function here that 
            ' unmanaged code can call. 
            ' To do this:
            ' 1. Export a function from unmanaged code that takes a function
            '    pointer and stores the passed in value in a global variable.
            ' 2. Call this exported function in this function passing delegate.
            ' 3. When unmanaged code needs the services of this managed module
            '    you simply call acrxLoadApp() and by the time acrxLoadApp 
            '    returns  global function pointer is initialized to point to
            '    the C# delegate.
            ' For more info see: 
            ' http:'msdn2.microsoft.com/en-US/library/5zwkzwf4(VS.80).aspx
            ' http:'msdn2.microsoft.com/en-us/library/44ey4b32(VS.80).aspx
            ' http:'msdn2.microsoft.com/en-US/library/7esfatk4.aspx
            ' as well as some of the existing AutoCAD managed apps.

            ' Initialize your plug-in application hereTry
            oApp = CType(Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication, Autodesk.AutoCAD.Interop.AcadApplication)
            cfg = New UtilesAlberto.Conf(System.Reflection.Assembly.GetExecutingAssembly)
            clsA = New AutoCAD2acad.A2acad.A2acad(oApp, cfg._appFullPath, regAPPCliente)
            '
            'If Autodesk.AutoCAD.Runtime.ExtensionLoader.IsLoaded(autocad2acad) = True Then
            '    'MsgBox("Cargado " & autocad2acad)
            'Else
            '    dll2acad = Autodesk.AutoCAD.Runtime.ExtensionLoader.Load(autocad2acad)
            '    Call AppDomain.CurrentDomain.GetAssemblies()
            'End If

            '''' Leer configuración
            Try
                Dim txtError As String() = INICargar()
                If txtError(0) <> "" Then
                    MsgBox(txtError(0))
                    If Log Then cfg.PonLog(txtError(0), True)
                    Exit Sub
                Else
                    If Log Then cfg.PonLog(txtError(1), True)
                End If
            Catch ex As System.Exception
                MsgBox("Activate--> INICargar --> " & ex.Message)
                Exit Sub
            End Try
            ' 
            ' Rellenar las clasese con los datos de Excel
            'cPT = New PT
            '
            ' Cargar los datos del fichero LAYOUTDB.xlsx
            ' Si está abierto, propone cerrarlo o cancelar, salimos sin cargar el AddIn
            If uau.FicheroEstaAbiertoMensaje(LAYOUTDB, cfg._appname) = MsgBoxResult.Cancel Then
                Exit Sub
            End If

            Try
                clsD = New clsLAYOUTDBS4
            Catch ex As System.Exception
                MsgBox(ex.ToString)
            End Try
            If Log Then cfg.PonLog("Llenada coleccion seleccionables (clsD)", False)
            ' Poner los eventos del documento necesarios.
            AddHandler Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.DocumentActivated, AddressOf DocumentManager_DocumentActivated
            'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.DocumentCreated, AddressOf DocumentManager_DocumentActivated

            'Cuando arranca el Autocad, por defecto aparece un dibujo cargado. Si no se hace lo siguiente, pierde la asignación de los eventos y si se ejecuta funciones del plugin da error.
            If (Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Count > 0) Then
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.CurrentDocument
                DocumentManager_DocumentActivated(Nothing, Nothing)

            End If
            AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
            If Log Then cfg.PonLog("Añadidos enventos", False)

            'If (docAct <> Nothing) Then
            '    AddHandler docAct.Database.ObjectModified, AddressOf Database_ObjectModified
            '    AddHandler docAct.Database.ObjectErased, AddressOf Database_ObjectErased
            '    AddHandler docAct.Database.ObjectAppended, AddressOf Database_ObjectAppended

            'End If
        End Sub


        'Private Sub DocumentManager_DocumentActivated(sender As Object, e As DocumentCollectionEventArgs)
        '    docAct = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        '    If (docAct <> Nothing) Then
        '        AddHandler docAct.Database.ObjectModified, AddressOf Database_ObjectModified
        '        AddHandler docAct.Database.ObjectErased, AddressOf Database_ObjectErased
        '        AddHandler docAct.Database.ObjectAppended, AddressOf Database_ObjectAppended
        '    End If
        'End Sub

        Public Sub Terminate() Implements IExtensionApplication.Terminate
            ' Do plug-in application clean up here
            oApp = Nothing
            clsA = Nothing
            cfg = Nothing
            'Me.Finalize()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Sub
    End Class

End Namespace