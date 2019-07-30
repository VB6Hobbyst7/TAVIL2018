Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports a2 = AutoCAD2acad.A2acad

Public Class frmAgrupa
    Public conMensaje As Boolean = True
    Private Sub frmAgrupa_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "AGRUPAR ELEMENTOS LINEA - v" & cfg._appversion
        If clsA Is Nothing Then clsA = New a2.A2acad(oApp, cfg._appFullPath, regAPPCliente)
        app_procesointerno = True
        gbAdministrar.Enabled = False
        btnGrupoBorrar.Enabled = False
        txtNombreGrupo.Text = ""
        lblInf.Text = ""
        tvGrupos_LlenaXDATA()
    End Sub

    Private Sub frmAgrupa_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        app_procesointerno = False
        frmAg = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cerrar_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cerrar_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub tvGrupos_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvGrupos.AfterSelect
        gbAdministrar.Enabled = False
        btnGrupoBorrar.Enabled = False
        txtNombreGrupo.Text = ""
        lblInf.Text = ""
        '
        If tvGrupos.SelectedNode Is Nothing Then Exit Sub
        '
        gbAdministrar.Enabled = True
        btnGrupoBorrar.Enabled = True
        txtNombreGrupo.Text = tvGrupos.SelectedNode.Text
        Me.Grupo_SeleccionarObjetos(tvGrupos.SelectedNode.Text)
    End Sub

    Private Sub btnGrupoAdd_Click(sender As Object, e As EventArgs) Handles btnGrupoAdd.Click
        Dim queGrupo As String = tvGrupos.SelectedNode.Text
        '
        Me.Visible = False
        oApp.ActiveDocument.Activate()
        Dim arrEntities As ArrayList = clsA.SeleccionaDameEntitiesONSCREEN(solouna:=False)
        If arrEntities IsNot Nothing AndAlso arrEntities.Count > 0 Then
            For Each oEnt As AcadEntity In arrEntities
                clsA.XPonDato(oEnt, cGRUPO, queGrupo, True)
            Next
            Dim oNode As TreeNode = tvGrupos.SelectedNode
            tvGrupos.SelectedNode = Nothing
            tvGrupos.SelectedNode = oNode
        End If
        Me.Visible = True
    End Sub

    Private Sub btnGrupoQuitar_Click(sender As Object, e As EventArgs) Handles btnGrupoQuitar.Click
        Me.Visible = False
        oApp.ActiveDocument.Activate()
        Dim grupoAhora As String = tvGrupos.SelectedNode.Tag
        Dim arrEntities As ArrayList = clsA.SeleccionaDameEntitiesONSCREEN(solouna:=False)
        If arrEntities IsNot Nothing AndAlso arrEntities.Count > 0 Then
            For Each oEnt As AcadEntity In arrEntities
                Dim queG As String = clsA.XLeeDato(oEnt, cGRUPO)
                If queG = grupoAhora Then
                    clsA.XPonDato(oEnt, cGRUPO, "", True)
                    If oApp.ActiveDocument.ActiveSelectionSet IsNot Nothing Then
                        Try
                            oApp.ActiveDocument.ActiveSelectionSet.RemoveItems(oEnt)
                        Catch ex As Exception

                        End Try
                    End If
                End If
            Next
            'oApp.ActiveDocument.Regen(AcRegenType.acActiveViewport)
            'If oApp.ActiveDocument.ActiveSelectionSet IsNot Nothing AndAlso oApp.ActiveDocument.ActiveSelectionSet.Count > 0 Then
            '    btnGrupoZoom_Click(Nothing, Nothing)
            'End If
        End If
        Me.Visible = True
        Dim oNode As TreeNode = tvGrupos.SelectedNode
        tvGrupos.SelectedNode = Nothing
        tvGrupos.SelectedNode = oNode
    End Sub

    Private Sub btnGrupoCrear_Click(sender As Object, e As EventArgs) Handles btnGrupoCrear.Click
REPITE:
        Dim nuevoGrupo As String = InputBox("Nombre del nuevo grupo :", "CREAR GRUPO", "")
        If nuevoGrupo = "" Then Exit Sub
        '
        If tvGrupos.Nodes.ContainsKey(nuevoGrupo) Then
            MsgBox("Ya existe ese nombre de grupo...", MsgBoxStyle.Exclamation)
            GoTo REPITE
        End If
        '
        Dim oNode As New TreeNode
        oNode.Text = nuevoGrupo
        oNode.Name = nuevoGrupo
        oNode.Tag = nuevoGrupo
        oNode.ToolTipText = oNode.Tag
        Dim id As Integer = tvGrupos.Nodes.Add(oNode)
        oNode = Nothing
        tvGrupos.Sort()
        '
        tvGrupos.SelectedNode = tvGrupos.Nodes.Item(id)
        tvGrupos_AfterSelect(Nothing, Nothing)
    End Sub

    Private Sub btnGrupoBorrar_Click(sender As Object, e As EventArgs) Handles btnGrupoBorrar.Click
        If conMensaje = True Then
            If MsgBox("Cada objeto de este grupo se quitará del grupo (No se borra) ¿Está seguro...?", MsgBoxStyle.YesNo, "AVISOS AL USUARIO") = MsgBoxResult.No Then
                Exit Sub
            End If
        End If
        '
        Dim grupo As String = tvGrupos.SelectedNode.Tag.ToString        '[Nombre Grupo]
        Dim lGrupo As List(Of Long) = clsA.SeleccionaTodosObjetosXData("GRUPO", grupo)
        If lGrupo IsNot Nothing AndAlso lGrupo.Count > 0 Then
            For Each queId As Long In lGrupo
                Dim oEnt As AcadEntity = oApp.ActiveDocument.ObjectIdToObject(queId)
                clsA.XPonDato(oEnt, cGRUPO, "")
            Next
        End If
        txtNombreGrupo.Text = ""
        tvGrupos.Nodes.Remove(tvGrupos.SelectedNode)
        tvGrupos.SelectedNode = Nothing
        tvGrupos_AfterSelect(Nothing, Nothing)
    End Sub

#Region "FUNCIONES"
    Public Sub tvGrupos_LlenaXDATA()
        ' Rellenar tvGrupos con los grupos que haya ([nombre grupo]) Sacado de XData elementos (regAPPCliente, XData = "GRUPO")
        tvGrupos.Nodes.Clear()
        If clsA Is Nothing Then clsA = New a2.A2acad(oApp, cfg._appFullPath, regAPPCliente)
        Dim arrTodos As List(Of Long) = clsA.SeleccionaTodosObjetos(,, True)
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            Exit Sub
        End If
        ' Filtrar lista de grupo. Sacar nombres únicos.
        For Each queId As Long In arrTodos
            Dim acadObj As AcadObject = oApp.ActiveDocument.ObjectIdToObject(queId)
            Dim grupo As String = clsA.XLeeDato(acadObj, cGRUPO)
            If grupo = "" Then Continue For
            '
            If tvGrupos.Nodes.ContainsKey(grupo) Then Continue For
            '
            Dim oNode As New TreeNode
            oNode.Text = grupo
            oNode.Name = grupo
            oNode.Tag = grupo
            oNode.ToolTipText = oNode.Tag
            tvGrupos.Nodes.Add(oNode)
            oNode = Nothing
            acadObj = Nothing
        Next
        tvGrupos.Sort()
        '
        tvGrupos.SelectedNode = Nothing
        tvGrupos_AfterSelect(Nothing, Nothing)
    End Sub

    Public Sub Grupo_SeleccionarObjetos(grupo As String)
        Dim lGrupos As List(Of Long) = clsA.SeleccionaTodosObjetosXData(cGRUPO, grupo)
        If lGrupos IsNot Nothing AndAlso lGrupos.Count > 0 Then
            Dim arrSeleccion As New ArrayList
            For Each queId As Long In lGrupos
                arrSeleccion.Add(oApp.ActiveDocument.ObjectIdToObject(queId))
            Next
            If arrSeleccion.Count > 0 Then
                lblInf.Text = arrSeleccion.Count & " Elementos"
                clsA.SeleccionCreaResalta(arrSeleccion, 0, False)
                If cbZoom.Checked Then
                    Zoom_Seleccion()
                End If
            End If
        Else
            'If tvGrupos.Nodes.ContainsKey(grupo) Then
            '    tvGrupos.Nodes.Item(grupo).Remove()
            'End If
            lblInf.Text = "0 Elementos"
            'tvGrupos.SelectedNode = Nothing
            'tvGrupos_AfterSelect(Nothing, Nothing)
        End If
    End Sub
#End Region
End Class