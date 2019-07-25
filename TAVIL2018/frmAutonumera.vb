Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports a2 = AutoCAD2acad.A2acad


Public Class frmAutonumera

    Private Sub frmAutonumera_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "AUTONUMERACION - v" & cfg._appversion

        ElementoProxyRecomendado = RecomiendaElementoLibre()

    End Sub

    Private Sub frmAutonumera_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmAu = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub btnInsertaNum_Click(sender As Object, e As EventArgs) Handles btnInsertaNum.Click
        app_procesointerno = True
        If ElementoProxyRecomendado.Contains(".") = False Then
            MsgBox("Numeración con Formato Incorrecto", MsgBoxStyle.Critical, "AVISOS AL USUARIO")
            Exit Sub
        End If

        If colP_ExisteElemento(ElementoProxyRecomendado) Then
            MsgBox("La Numeración Indicada ya existe en un Proxy", MsgBoxStyle.Critical, "AVISOS AL USUARIO")
            Exit Sub
        End If

        If ExisteEnArrayProxiesEliminados(ElementoProxyRecomendado) Then
            MsgBox("La Numeración Indicada fue eliminada", MsgBoxStyle.Critical, "AVISOS AL USUARIO")
            Exit Sub
        End If


        If clsA.oAppA.ActiveDocument.ActiveSpace = Autodesk.AutoCAD.Interop.Common.AcActiveSpace.acPaperSpace Then
            MsgBox("Proxy sólo se puede insertar en Espacio Modelo", MsgBoxStyle.Critical, "AVISOS AL USUARIO")
            Exit Sub
        End If

        '
        Me.Visible = False
        ' Cargar recursos
        clsA.ClonaTODODesdeDWG(BloqueRecursos)

        Try
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMLEADERSTYLE", fijoCliente)
        Catch ex As Exception
            MsgBox("No existe el estilo " & fijoCliente)
            Exit Sub
        End Try
        ' Activar la capa 'proxy' y quitar el contorno de cobertura
        clsA.CapaActiva(capaproxy)
        clsA.CoberturaOnOff(False)

        Dim bolFin As Boolean = False
        Do While bolFin = False
            'Filtro para buscar a que BlockReference asociar el proxy
            Dim arrayFilterType() As Type = {GetType(BlockReference)}
            Dim oIdSel As ObjectId = clsA.ObjectId_SelectOne(OpenMode.ForRead, arrayFilterType)

            If oIdSel <> Nothing Then
                Dim oEntity As Entity = clsA.Entity_Get(oIdSel)
                'Mira si el blockReference seleccionado tiene el atributo ELEMENTO y que esta vacía para osociarle uno nuevo.
                Dim strElementoEntity As String
                Try
                    'strElementoEntity = clsA.BloqueAtributoDame(CType(oEntity.AcadObject, AcadBlockReference).ObjectID, "ELEMENTO")
                    strElementoEntity = clsA.XLeeDato(CType(oEntity.AcadObject, AcadBlockReference), "ELEMENTO")
                Catch ex As Exception
                    strElementoEntity = "-1" 'Ha seleccionado un blockReference que no tiene atributo ELEMENTO y por lo tanto no se le puede asignar un proxy
                End Try

                If (strElementoEntity = "") Then
                    Dim oMl As AcadMLeader = Nothing
                    oMl = clsA.MLeader_InsertaCommand()
                    If oMl IsNot Nothing Then

                        'Actualiza el proxy con el Elemento (Familia.ID)
                        clsA.MLeaderBlock_PonValorAtributo(oMl, "ELEMENTO", ElementoProxyRecomendado.Split(".")(1))
                        clsA.XPonDato(oMl, "ELEMENTO", ElementoProxyRecomendado)    ', regAPPCliente)

                        colP_AddElemento(ElementoProxyRecomendado, oMl)

                        'clsA.BloqueAtributoEscribe(CType(oEntity.AcadObject, AcadBlockReference).ObjectID, "ELEMENTO", ElementoProxyRecomendado)
                        clsA.XPonDato(CType(oEntity.AcadObject, AcadBlockReference), "ELEMENTO", ElementoProxyRecomendado)  ', regAPPCliente)

                        ElementoProxyRecomendado = RecomiendaElementoLibre(oEntity)
                    End If
                ElseIf (strElementoEntity = "-1") Then
                    MsgBox("Error al leer XData del BlockReference")
                Else
                    MsgBox("El BlockReference seleccionado ya dispone del proxy: " & strElementoEntity)
                End If
            Else
                bolFin = True
            End If
        Loop
        app_procesointerno = False
        Me.Visible = True



    End Sub



    Private Sub btnCargarNumeracion_Click(sender As Object, e As EventArgs) Handles btnCargarNumeracion.Click
        app_procesointerno = True
        If clsA.oAppA.ActiveDocument.ActiveSpace = Autodesk.AutoCAD.Interop.Common.AcActiveSpace.acPaperSpace Then
            MsgBox("Proxy sólo puede insertarse en Espacio Modelo", MsgBoxStyle.Critical, "AVISOS AL USUARIO")
            Exit Sub
        End If
        '
        Me.Visible = False
        ' Cargar recursos
        clsA.ClonaTODODesdeDWG(BloqueRecursos)

        Try
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CMLEADERSTYLE", fijoCliente)
        Catch ex As Exception
            MsgBox("No existe el estilo " & fijoCliente)
            Exit Sub
        End Try
        ' Activar la capa 'proxy' y quitar el contorno de cobertura
        clsA.CapaActiva(capaproxy)
        clsA.CoberturaOnOff(False)
        ' Insertar el MLeader


        Dim arrayFilterType() As Type = {GetType(BlockReference), GetType(MLeader)}
        Dim oIdSel As ObjectId = clsA.ObjectId_SelectOne(OpenMode.ForRead, arrayFilterType)
        Dim oMl As AcadMLeader = Nothing

        Dim oEntity As Entity = Nothing

        If oIdSel <> Nothing Then
            oEntity = clsA.Entity_Get(oIdSel)
        End If
        ElementoProxyRecomendado = RecomiendaElementoLibre(oEntity)
        app_procesointerno = False
        Me.Visible = True

    End Sub




End Class