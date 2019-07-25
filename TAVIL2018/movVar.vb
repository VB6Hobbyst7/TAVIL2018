﻿Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports AutoCAD2acad.A2acad
Imports System.Linq
Imports ua = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad

Module movVar
    ' ***** OBJETOS AUTOCAD
    Public WithEvents oApp As Autodesk.AutoCAD.Interop.AcadApplication = Nothing
    Public WithEvents oDoc As Autodesk.AutoCAD.Interop.AcadDocument = Nothing
    Public ed As Autodesk.AutoCAD.EditorInput.Editor
    Public oSel As Autodesk.AutoCAD.Interop.AcadSelectionSet
    Public oBlR As AcadBlockReference = Nothing     ' AcadBlockReference de la cinta que seleccionemos.
    ' ***** OBJETOS AUTOCAD ACTIVE X
    Public docAct As Document

    '
    ' ***** CLASES
    Public clsA As a2.A2acad = Nothing
    Public clsD As clsLAYOUTDBS4 = Nothing
    Public cfg As UtilesAlberto.Conf
    'Public cXML As ClosedXML2acad.ClosedXML2acad
    Public pataD As clsBloquePataDatos     ' Datos de un bloque de pata (Parametros y Atributos)
    ' ***** CLASES con los datos de cada Hoja Excel
    Public cPT As PT     ' Hoja PT de LAYOUTDBS4.xlsx

    '
    ' ***** FORMULARIO
    Public frmCo As frmConfigura
    Public frmAu As frmAutonumera
    Public frmPa As frmPatas1
    Public frmUn As frmUniones
    Public frmAg As frmAgrupa
    Public frmBo As frmBomDatos
    Public frmBlo As frmBloques
    Public frmBloE As frmBloquesEditar
    '
    ' ***** COLECCIONES
    Public colIds As List(Of Long)
    Public colHan As List(Of String)
    Public dicBloques As Dictionary(Of String, String) = Nothing          ' Key=nombre bloque (Sin extension), Value=Directorio Bloque
    Public lnBloquesPatas As List(Of String)     ' Nombres únicos de bloques de patas
    Public lnRADIUS As List(Of String)           ' Nombres únicos de RADIUS de patas
    Public lnWIDTH As List(Of String)            ' Nombres únicos de WIDTH de patas
    Public lnHEIGHT As List(Of String)           ' Nombres únicos de HEIGHT de patas
    Public patasSIPlanta As String() = {"PLANTA", "TVIEW"}
    Public patasNoPlanta As String() = {"ALÇAT1", "ALÇAT2", "FVIEW", "SVIEW"}
    '
    ' ***** CONSTANTES
    Public Const regAPPCliente As String = "TAVIL2acad"
    Public Const fijoCliente As String = "TAVIL"
    Public Const fijoYear As String = "2018"
    Public Const fijoClienteYear As String = fijoCliente & fijoYear
    Public Const capaproxytabla As String = "CAPABOM"
    Public Const capaproxy As String = "PROXY"
    Public Const estilotexto As String = fijoCliente & "_TEXTO"
    Public Const estilotabla As String = fijoCliente & "_TABLA"
    Public Const cGRUPO As String = "GRUPO"
    '
    ' ***** VARIABLES CONFIGURACION
    Public Log As Boolean = False
    Public BloqueRecursos As String = ""    ' DWG con todos los recursos del desarrollo
    Public sizeImg As Integer = 100
    Public LAYOUTDB As String = "LAYOUTDBS4.xlsx"                  ' fullPath del fichero LAYOUTDB.xlsx que hace de base de datos.
    Public BloquesDir As String = "BLOQUES"
    Public PatasDir As String = "PEUS"
    Public PatasCapa As String = "PEUS I ACCESOS"
    Public PatasRef As String = "REF.PEUS"
    Public CintasDir As String = "CINTES"
    Public CintasCapa As String = "CINTES"
    Public CintasRef As String = "REF.CINTES"
    '
    '
    ' ***** VARIABLES APLICACION
    'Public app_folder As String = My.Application.Info.DirectoryPath     '' Solo Directorio
    'Public app_name As String = My.Application.Info.AssemblyName        '' 
    'Public app_folderandname As String = IO.Path.Combine(app_folder, app_name)        '' 
    'Public app_fullPath As String = IO.Path.Combine(app_folder, app_name & ".dll")
    'Public app_version As String = My.Application.Info.Version.ToString
    'Public app_nameandversion As String = app_name & " - v" & app_version
    'Public app_log As String = app_folderandname & ".log"
    'Public app_conf As String = app_fullPath & ".config"
    Public app_procesointerno As Boolean = False 'afleta
    '
#Region "UTILITIES"
    Public Function INICargar() As String()
        If cfg Is Nothing Then cfg = New UtilesAlberto.Conf(System.Reflection.Assembly.GetExecutingAssembly)
        'If cXML Is Nothing Then cXML = New ClosedXML2acad.ClosedXML2acad
        Dim mensaje(1) As String
        '' Mensaje(0) contendrá los errores.
        '' mensaje(1) contendrá los valores leidos del .INI
        mensaje(0) = "" : mensaje(1) = ""
        mensaje(1) &= "***** Configuración :" & vbCrLf
        '[OPTIONS]
        'Log=1
        'BloqueRecursos=BloqueRecursos.dwg
        'sizeImg=175
        'LAYOUTDB=LAYOUTDB.xlsx
        'BloquesDir=BLOQUES
        'PatasDir=PATAS
        'PatasCapa=PEUS I ACCESOS
        'PatasRef=REF.PEUS
        'CintasDir=CINTES
        'CintasCapa=CINTES
        'CintasRef=REF.CINTES
        'patasSiPlanta=PLANTA,TVIEW
        'patasNoPlanta = ALCAT1,ALÇAT2,FVIEW,SVIEW

        Dim LogTemp As String = ua.IniGet(cfg._appini, "OPTIONS", "Log")
        Log = IIf(LogTemp = "1", True, False)
        mensaje(1) &= "Log = " & Log & vbCrLf
        '
        BloqueRecursos = ua.IniGet(cfg._appini, "OPTIONS", "BloqueRecursos")
        If BloqueRecursos <> "" Then
            BloqueRecursos = IO.Path.Combine(cfg._appfolder, BloqueRecursos)
        End If
        mensaje(1) &= "BloqueRecursos = " & BloqueRecursos & vbCrLf
        '
        Dim sizeImgTemp As String = ua.IniGet(cfg._appini, "OPTIONS", "sizeImg")
        sizeImg = IIf(IsNumeric(sizeImgTemp), CInt(sizeImgTemp), 100)
        mensaje(1) &= "sizeImg = " & sizeImg.ToString & vbCrLf
        '
        LAYOUTDB = ua.IniGet(cfg._appini, "OPTIONS", "LAYOUTDB")
        If LAYOUTDB.StartsWith(".\") Then
            LAYOUTDB = LAYOUTDB.Replace(".\", cfg._appfolder & "\")
        ElseIf LAYOUTDB.Contains("\") = False And LAYOUTDB.Contains(":") = False Then
            LAYOUTDB = IO.Path.Combine(cfg._appfolder, LAYOUTDB)
        End If
        mensaje(1) &= "LAYOUTDB = " & LAYOUTDB & vbCrLf
        '
        BloquesDir = ua.IniGet(cfg._appini, "OPTIONS", "BloquesDir")
        If BloquesDir.StartsWith(".\") Then
            BloquesDir = BloquesDir.Replace(".\", cfg._appfolder & "\")
        ElseIf BloquesDir.Contains("\") = False And BloquesDir.Contains(":") = False Then
            BloquesDir = IO.Path.Combine(cfg._appfolder, BloquesDir)
        End If
        mensaje(1) &= "BloquesDir = " & BloquesDir & vbCrLf
        '
        PatasDir = ua.IniGet(cfg._appini, "OPTIONS", "PatasDir")
        PatasDir = IO.Path.Combine(BloquesDir, PatasDir)
        mensaje(1) &= "PatasDir = " & PatasDir & vbCrLf
        '
        PatasCapa = ua.IniGet(cfg._appini, "OPTIONS", "PatasCapa")
        PatasRef = ua.IniGet(cfg._appini, "OPTIONS", "PatasRef")
        mensaje(1) &= "PatasCapa = " & PatasCapa & vbCrLf
        mensaje(1) &= "PatasRef = " & PatasRef & vbCrLf
        '
        CintasDir = ua.IniGet(cfg._appini, "OPTIONS", "CintasDir")
        CintasDir = IO.Path.Combine(BloquesDir, CintasDir)
        mensaje(1) &= "CintasDir = " & CintasDir & vbCrLf
        '
        CintasCapa = ua.IniGet(cfg._appini, "OPTIONS", "CintasCapa")
        CintasRef = ua.IniGet(cfg._appini, "OPTIONS", "CintasRef")
        mensaje(1) &= "CintasCapa = " & CintasCapa & vbCrLf
        mensaje(1) &= "CintasRef = " & CintasRef & vbCrLf
        '
        Dim partes() As String = ua.IniGet(cfg._appini, "OPTIONS", "patasSiPlanta").Split(","c)
        If partes IsNot Nothing AndAlso partes.Count > 0 Then
            ReDim patasSIPlanta(partes.Count - 1) : partes.CopyTo(patasSIPlanta, 0)
        End If
        ReDim partes(-1)
        partes = ua.IniGet(cfg._appini, "OPTIONS", "patasNoPlanta").Split(","c)
        If partes IsNot Nothing AndAlso partes.Count > 0 Then
            ReDim patasNoPlanta(partes.Count - 1) : partes.CopyTo(patasNoPlanta, 0)
        End If
        '
        ' Al fichero log la configuración leida.
        'If Log Then cfg.PonLog(mensaje(1), True)
        '
        ' ***** Comprobar ficheros y directorios. Si no existe alguno, devolvemos error.
        ' FICHEROS
        If IO.File.Exists(BloqueRecursos) = False Then
            mensaje(0) &= "No existe fichero " & BloqueRecursos & vbCrLf
        End If
        '
        If IO.File.Exists(LAYOUTDB) = False Then
            mensaje(0) &= "No existe fichero " & LAYOUTDB & vbCrLf
        End If
        ' DIRECTORIOS
        If IO.Directory.Exists(BloquesDir) = False Then
            mensaje(0) &= "No existe directorio " & BloquesDir & vbCrLf
        Else
            dicBloques_LlenaConDirRaiz(BloquesDir)
        End If
        '
        If IO.Directory.Exists(PatasDir) = False Then
            mensaje(0) &= "No existe directorio " & PatasDir & vbCrLf
        End If
        '
        If IO.Directory.Exists(CintasDir) = False Then
            mensaje(0) &= "No existe directorio " & CintasDir & vbCrLf
        End If
        '
        Return mensaje
    End Function
    '
    Public Sub dicBloques_LlenaConDirRaiz(queDirRaiz As String)
        If IO.Directory.Exists(queDirRaiz) = False Then Exit Sub
        '
        dicBloques = New Dictionary(Of String, String)
        For Each queFi As String In IO.Directory.GetFiles(queDirRaiz, "*.dwg", IO.SearchOption.AllDirectories)
            Dim nombre As String = IO.Path.GetFileNameWithoutExtension(queFi)
            Dim directorio As String = IO.Path.GetDirectoryName(queFi)
            If dicBloques.ContainsKey(nombre) = False Then dicBloques.Add(nombre, directorio)
        Next
    End Sub
    '
    Public Sub CierraFormularios()
        If frmCo IsNot Nothing Then frmCo.Close()
        If frmAu IsNot Nothing Then frmAu.Close()
        If frmPa IsNot Nothing Then frmPa.Close()
        If frmUn IsNot Nothing Then frmUn.Close()
        If frmAg IsNot Nothing Then frmAg.Close()
        If frmBo IsNot Nothing Then frmBo.Close()
    End Sub

    Private Sub oApp_AppActivate() Handles oApp.AppActivate

    End Sub


    Private Sub oApp_EndCommand(CommandName As String) Handles oApp.EndCommand
        'If CommandName = "OPEN" And oApp.Documents.Count > 0 Then
        'modTavil.AcadBlockReference_PonEventosModified()
        'End If
        If oApp.Documents.Count = 0 Then Exit Sub
        'If oDoc IsNot Nothing AndAlso oDoc.Equals(oApp.ActiveDocument) = False Then oDoc = oApp.ActiveDocument
        'If CommandName.ToUpper.Contains("COPY") Then
        '    'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        'End If
    End Sub

    Private Sub oApp_EndOpen(FileName As String) Handles oApp.EndOpen
        'While oApp.GetAcadState.IsQuiescent = True
        '    System.Windows.Forms.Application.DoEvents()
        'End While
        ' modTavil.AcadBlockReference_PonEventosModified()
        'oDoc = oApp.ActiveDocument
        ''
        'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
        '
    End Sub

    Private Sub oDoc_BeginCommand(CommandName As String) Handles oDoc.BeginCommand
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

    Private Sub oDoc_ObjectModified(queObj As Object) Handles oDoc.ObjectModified
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
    Public Sub Tavil_AppIdle(ByVal sender As Object, ByVal e As EventArgs)
        RemoveHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
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

    Public Sub Database_ObjectModified(ByVal sender As Object, ByVal e As ObjectEventArgs)
        'AutoEnumera_DBObjectModified(sender, e)
        ''oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
    End Sub

    Public Sub Database_ObjectAppended(ByVal sender As Object, ByVal e As ObjectEventArgs)
        'Try
        '    If TypeOf e.DBObject.AcadObject Is AcadBlockReference AndAlso CType(e.DBObject.AcadObject, AcadBlockReference).EffectiveName.ToUpper.StartsWith("PT") Then
        '        Exit Sub
        '    End If
        'Catch ex As Exception
        '    Exit Sub
        'End Try
        'AutoEnumera_DBObjectAppended(sender, e)
        ''oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
        'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
    End Sub

    Public Sub Database_ObjectErased(ByVal sender As Object, ByVal e As ObjectErasedEventArgs)
        'AutoEnumera_ObjectErased(sender, e)
        'AddHandler Autodesk.AutoCAD.ApplicationServices.Application.Idle, AddressOf Tavil_AppIdle
    End Sub

    Public Sub DocumentManager_DocumentActivated(sender As Object, e As DocumentCollectionEventArgs)
        'docAct = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ''
        'If (docAct <> Nothing) Then
        '    oDoc = oApp.ActiveDocument
        '    AddHandler docAct.Database.ObjectModified, AddressOf Database_ObjectModified
        '    AddHandler docAct.Database.ObjectErased, AddressOf Database_ObjectErased
        '    AddHandler docAct.Database.ObjectAppended, AddressOf Database_ObjectAppended
        '    colP_ProxiesRellena()
        '    arrayProxiesEliminados = clsA.PropiedadCustomDocumento_Lee("ElementosProxiesEliminados").Split("·")
        '    If colP.Count > 0 Then
        '        Dim oI As ObjectId = New ObjectId(colP.First().Value.First.oMl.ObjectID)
        '        ElementoProxyRecomendado = RecomiendaElementoLibre(clsA.Entity_Get(oI))
        '    Else
        '        ElementoProxyRecomendado = RecomiendaElementoLibre()
        '    End If
        'End If
    End Sub
#End Region
End Module
