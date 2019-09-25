Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Linq
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad

Public Class frmUniones
    Private UltimoBloqueUnion As AcadBlockReference = Nothing
    Private UltimoBloqueT1 As AcadBlockReference = Nothing
    Private UltimoBloqueT2 As AcadBlockReference = Nothing
    Private UltimaClsUnion As ClsUnion = Nothing
    Private UltimaFilaExcel As ExcelFila
    Private HighlightedPictureBox As PictureBox = Nothing
    Private oTT As ToolTip = Nothing
    Private capaUnionesVisible As Boolean = True
    Private EsUnionNueva As Boolean = False

    Private Sub frmUniones_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = HojaUniones & " - v" & cfg._appversion
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp, cfg._appFullPath, regAPPCliente)
        ' Cargar recursos
        Using oLock As DocumentLock = Eventos.AXDoc.LockDocument
            clsA.ClonaTODODesdeDWG(BloqueRecursos, True, True)
            clsA.ClonaTODODesdeDWG(IO.Path.Combine(IO.Path.GetDirectoryName(BloqueRecursos), "UNION.dwg"), True, True)
        End Using
        PonToolTipControles()
        Dim queCapa As AcadLayer = clsA.CapaDame(HojaUniones)
        If queCapa IsNot Nothing Then
            capaUnionesVisible = queCapa.LayerOn
            queCapa.LayerOn = True
            queCapa.Lock = False
            queCapa.Freeze = False
            cbCapa.Checked = capaUnionesVisible
        End If
        clsA.SelectionSet_Delete()
        cbTipo.Text = "TODOS"
    End Sub

    Private Sub frmUniones_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If cbCapa.Enabled = True Then
            Dim queCapa As AcadLayer = clsA.CapaDame(HojaUniones)
            If queCapa IsNot Nothing Then
                queCapa.LayerOn = capaUnionesVisible
                queCapa.Freeze = Not (cbCapa.Checked)
            End If
        End If
        frmUn = Nothing
    End Sub

    Private Sub btnCerrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCerrar.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub
    '
    ' CONTROLES
    Private Sub cbTipo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbTipo.SelectedIndexChanged
        tvUniones_Rellena(cbTipo.SelectedIndex)
    End Sub
    '
    Private Sub tvUniones_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tvUniones.AfterSelect
        clsA.Selection_Quitar()
        UltimoBloqueT1 = Nothing
        UltimoBloqueT2 = Nothing
        UltimoBloqueUnion = Nothing
        UltimaClsUnion = Nothing
        If tvUniones.SelectedNode Is Nothing Then
            PonEstadoControlesInicial()
            Exit Sub
        End If
        BtnEditarUnion.Enabled = True
        BtnBorrarUnion.Enabled = True
        Me.Uniones_SeleccionarObjetos(tvUniones.SelectedNode.Tag.ToString)
        ' Mostrar los datos en sus controles.
        PonDatosUnion(tvUniones.SelectedNode.Tag.ToString)
    End Sub

    Private Sub PonDatosUnion(Optional handle As String = "")
        If tvUniones.SelectedNode Is Nothing Then Exit Sub
        If handle = "" Then handle = tvUniones.SelectedNode.Tag.ToString
        If handle = "" Then Exit Sub
        If IsNumeric(handle) Then Exit Sub
        '
        Try
            UltimoBloqueUnion = Eventos.COMDoc().HandleToObject(handle)
        Catch ex As Exception
            Exit Sub
        End Try
        '
        Dim lUniHANLE As IEnumerable(Of ClsUnion) = From x In LClsUnion
                                                    Where x.HANDLE = handle
                                                    Select x
        If lUniHANLE.Count > 0 Then
            UltimaClsUnion = lUniHANLE.First
        Else
            Exit Sub
        End If
        '
        ' T1
        If UltimaClsUnion.T1HANDLE <> "" Then
            Try
                UltimoBloqueT1 = Eventos.COMDoc().HandleToObject(UltimaClsUnion.T1HANDLE)
                If UltimaClsUnion.T1INCLINATION <> "" Then ListBox_SeleccionaPorTexto(LbInclinationT1, UltimaClsUnion.T1INCLINATION)
                LblT1.Text = "Datos T1:" & vbCrLf & UltimoBloqueT1.EffectiveName & vbCrLf & UltimaClsUnion.T1INFEED
            Catch ex As Exception
            End Try
        End If
        ' T2
        If UltimaClsUnion.T2HANDLE <> "" Then
            Try
                UltimoBloqueT2 = Eventos.COMDoc().HandleToObject(UltimaClsUnion.T2HANDLE)
                If UltimaClsUnion.T2INCLINATION <> "" Then ListBox_SeleccionaPorTexto(LbInclinationT2, UltimaClsUnion.T2INCLINATION)
                LblT2.Text = "Datos T2:" & vbCrLf & UltimoBloqueT2.EffectiveName & vbCrLf & UltimaClsUnion.T2OUTFEED
            Catch ex As Exception
            End Try
        End If
        '
        If UltimaClsUnion.ROTATION = "" OrElse UltimaClsUnion.ROTATION = "0" Then
            LbRotation.SelectedIndex = 0
        ElseIf UltimaClsUnion.ROTATION = "90" Then
            LbRotation.SelectedIndex = 1
        Else
            LbRotation.SelectedIndex = -1
        End If
        '
        LbUnion.Items.Clear()
        If UltimaClsUnion.UNION <> "" Then
            LbUnion.Tag = New String() {UltimaClsUnion.ExcelFilaUnion.UNION, UltimaClsUnion.ExcelFilaUnion.UNITS}
            LbUnion.Items.AddRange(UltimaClsUnion.ExcelFilaUnion.UNION.Split(";"c))
            LbUnion.Text = UltimaClsUnion.UNION
        End If
        '
        LblUnits.Text = UltimaClsUnion.UNITS
    End Sub
    Private Sub tvUniones_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles tvUniones.MouseDoubleClick
        Uniones_SeleccionarObjetos(tvUniones.SelectedNode.Tag, conZoom:=True)
    End Sub
    Private Sub BtnCrearUnion_Click(sender As Object, e As EventArgs) Handles BtnCrearUnion.Click
        EsUnionNueva = True
        BorraDatos()
        UltimoBloqueT1 = Nothing
        UltimoBloqueT2 = Nothing
        UltimoBloqueUnion = Nothing
        UltimaClsUnion = Nothing
        '
        GUnion.Enabled = True
        BtnInsertarUnion.Visible = True
        BtnInsertarUnion.Enabled = False
        BtnAceptar.Visible = False
        BtnAceptar.Enabled = False
    End Sub

    Private Sub BtnEditarUnion_Click(sender As Object, e As EventArgs) Handles BtnEditarUnion.Click
        If tvUniones.SelectedNode Is Nothing Then Exit Sub
        '
        EsUnionNueva = False
        GUnion.Enabled = True
        BtnInsertarUnion.Visible = False
        BtnInsertarUnion.Enabled = False
        BtnAceptar.Visible = True
        BtnAceptar.Enabled = True
    End Sub
    Private Sub BtnBorrarUnion_Click(sender As Object, e As EventArgs) Handles BtnBorrarUnion.Click
        clsA.Selection_Quitar()
        Dim oBl As AcadBlockReference = Eventos.COMDoc().HandleToObject(tvUniones.SelectedNode.Tag)
        If oBl Is Nothing Then Exit Sub
        '
        oBl.Delete()
        tvUniones.SelectedNode.Remove()
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
        BtnSeleccionar.Enabled = (tvUniones.Nodes.Count > 0)
    End Sub
    Private Sub BtnSeleccionar_Click(sender As Object, e As EventArgs) Handles BtnSeleccionar.Click
        Dim oBl As AcadBlockReference = clsA.Bloque_SeleccionaDame
        If oBl Is Nothing Then Exit Sub
        '
        tvUniones.SelectedNode = Nothing
        For Each oNode As TreeNode In tvUniones.Nodes
            If oNode.Tag = oBl.ObjectID Then
                tvUniones.SelectedNode = oNode
                tvUniones_MouseDoubleClick(tvUniones, Nothing)  ' Hacer Zoom en el elemento.
                Exit For
            End If
        Next
    End Sub
    Private Sub CbCapa_CheckedChanged(sender As Object, e As EventArgs) Handles cbCapa.CheckedChanged
        Dim queCapa As AcadLayer = clsA.CapaDame(HojaUniones)
        If queCapa IsNot Nothing Then
            queCapa.LayerOn = cbCapa.Checked
            queCapa.Freeze = Not (cbCapa.Checked)
            Eventos.COMDoc().Regen(AcRegenType.acActiveViewport)
            PonEstadoControlesInicial()
            tvUniones.Enabled = cbCapa.Checked
            BtnCrearUnion.Enabled = cbCapa.Checked
            BtnSeleccionar.Enabled = cbCapa.Checked
        End If
        BtnCancelar_Click(Nothing, Nothing)
    End Sub
    Private Sub BtnInsertaMultiplesUniones_Click(sender As Object, e As EventArgs) Handles BtnInsertaMultiplesUniones.Click
        Me.Visible = False
        clsA.ActivaAppAPI()
        Dim oBlr As AcadBlockReference = Nothing
        Try
            Do
                oBlr = clsA.Bloque_InsertaMultiple(, NombreBloqueUNION)
                'If oBlr IsNot Nothing Then
                '    'tvUniones_PonNodo(oBlr.ObjectID)
                '    'tvUniones.SelectedNode = Nothing
                '    'tvUniones_AfterSelect(Nothing, Nothing)
                'End If
            Loop While oBlr IsNot Nothing
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        tvUniones_Rellena()
        Me.Visible = True
    End Sub

    Private Sub BtnT1_Click(sender As Object, e As EventArgs) Handles BtnT1.Click
        Me.Visible = False
        clsA.ActivaAppAPI()
        UltimoBloqueT1 = clsA.Bloque_SeleccionaDame
        If UltimoBloqueT1 IsNot Nothing AndAlso UltimoBloqueUnion IsNot Nothing Then
            clsA.XPonDato(UltimoBloqueUnion.Handle, "T1HANDLE", UltimoBloqueT1.Handle)
        End If
        CompruebaDatos()
        Me.Visible = True
    End Sub
    Private Sub BtnT2_Click(sender As Object, e As EventArgs) Handles BtnT2.Click
        Me.Visible = False
        clsA.ActivaAppAPI()
        UltimoBloqueT2 = clsA.Bloque_SeleccionaDame
        If UltimoBloqueT2 IsNot Nothing AndAlso UltimoBloqueUnion IsNot Nothing Then
            clsA.XPonDato(UltimoBloqueUnion.Handle, "T2HANDLE", UltimoBloqueT2.Handle)
        End If
        CompruebaDatos()
        Me.Visible = True
    End Sub


    Private Sub LbT1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LbInclinationT1.SelectedIndexChanged, LbInclinationT2.SelectedIndexChanged, LbRotation.SelectedIndexChanged
        If GUnion.Enabled = False Then Exit Sub
        '
        If UltimoBloqueT1 IsNot Nothing AndAlso
            UltimoBloqueT2 IsNot Nothing AndAlso
            LbInclinationT1.SelectedIndex >= 0 AndAlso
            LbInclinationT2.SelectedIndex >= 0 AndAlso
            LbRotation.SelectedIndex >= 0 Then
            PonUnion()
        End If
    End Sub
    Private Sub LbUnion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LbUnion.SelectedIndexChanged
        If GUnion.Enabled = False Then Exit Sub
        '
        If LbUnion.Tag Is Nothing OrElse LbUnion.Tag.ToString = "" OrElse LbUnion.SelectedIndex = -1 OrElse LbUnion.Text = "" Then
            BtnInsertarUnion.Enabled = False
            BtnAceptar.Enabled = False
            LbUnion.Text = ""
            LblUnits.Text = ""
            Exit Sub
        End If
        '
        Dim datos() As String = LbUnion.Tag
        If datos Is Nothing Then Exit Sub
        If datos.GetLength(0) <> 2 Then Exit Sub
        If datos(0) = "" Then Exit Sub
        '
        '
        Dim union As String() = datos(0).Split(";"c)
        Dim units As String() = datos(1).Split(";"c)
        For x As Integer = 0 To UBound(union)
            If union(x) = LbUnion.Text Then
                LblUnits.Text = units(x)
                Exit For
            End If
        Next
        PonEstadoBotonesInsertarAceptar()
    End Sub

    Private Sub BtnInsertarUnion_Click(sender As Object, e As EventArgs) Handles BtnInsertarUnion.Click
        clsA.ActivaAppAPI()
        If UltimoBloqueUnion Is Nothing Then
            clsA.Bloque_Inserta(, NombreBloqueUNION)
            If clsA.oBlult IsNot Nothing Then
                UltimoBloqueUnion = Eventos.COMDoc.HandleToObject(clsA.oBlult.Handle)
                LClsUnion.Add(New ClsUnion(UltimoBloqueUnion.Handle, LbUnion.Text, LblUnits.Text, UltimoBloqueT1.Handle, IIf(UltimaFilaExcel IsNot Nothing, UltimaFilaExcel.INFEED_CONVEYOR, ""), LbInclinationT1.Text, UltimoBloqueT2.Handle, IIf(UltimaFilaExcel IsNot Nothing, UltimaFilaExcel.OUTFEED_CONVEYOR, ""), LbInclinationT2.Text, LbRotation.Text))
                tvUniones_Rellena()
            End If
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        UltimoBloqueT1 = Nothing
        UltimoBloqueT2 = Nothing
        UltimoBloqueUnion = Nothing
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
        BorraDatos()
    End Sub
    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        If UltimaClsUnion IsNot Nothing Then
            UltimaClsUnion.T1HANDLE = UltimoBloqueT1.Handle
            UltimaClsUnion.T1INCLINATION = LbInclinationT1.Text
            UltimaClsUnion.T1INFEED = IIf(UltimaFilaExcel IsNot Nothing, UltimaFilaExcel.INFEED_CONVEYOR, "")
            UltimaClsUnion.T2HANDLE = UltimoBloqueT2.Handle
            UltimaClsUnion.T2INCLINATION = LbInclinationT2.Text
            UltimaClsUnion.T2OUTFEED = IIf(UltimaFilaExcel IsNot Nothing, UltimaFilaExcel.OUTFEED_CONVEYOR, "")
            UltimaClsUnion.UNION = LbUnion.Text
            UltimaClsUnion.UNITS = LblUnits.Text
            UltimaClsUnion.ROTATION = LbRotation.Text
            UltimaClsUnion.ExcelFilaUnion = UltimaFilaExcel
        End If
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
    End Sub
#Region "FUNCIONES"
    Public Sub PonEstadoControlesInicial()
        ' Estado inicila de GUniones (Con todos los controles de selección y edición)
        tvUniones.SelectedNode = Nothing
        GUnion.Enabled = False
        BtnCrearUnion.Enabled = True
        BtnEditarUnion.Enabled = False
        BtnBorrarUnion.Enabled = False
        GUnion.Enabled = False
        BtnT1.Enabled = False : BtnT1.BackColor = btnOn
        LblT1.Text = "Datos T1: "
        BtnT2.Enabled = False : BtnT2.BackColor = btnOn
        LblT2.Text = "Datos T2: "
        BtnInsertarUnion.Enabled = False : BtnInsertarUnion.BackColor = btnOn
        BorraDatos()
    End Sub
    Public Sub PonToolTipControles()
        oTT = New ToolTip()
        oTT.AutoPopDelay = 5000 ' Tiempo que estará visible
        oTT.InitialDelay = 500  ' Tiempo inicial para mostrarse
        oTT.ReshowDelay = 100   ' Tiempo de espera entre controles
        oTT.ShowAlways = True   ' Forzar a que se muestre el tooltip, aunque no este activo el Form.
        '
        Me.tvUniones.ShowNodeToolTips = True
        oTT.SetToolTip(Me.tvUniones, "Listado de uniones")
        oTT.SetToolTip(Me.cbZoom, "Zoom uniones")
        '
        oTT.SetToolTip(Me.BtnCrearUnion, "Crear nueva unión")
        oTT.SetToolTip(Me.BtnEditarUnion, "Editar unión seleccionada")
        oTT.SetToolTip(Me.BtnBorrarUnion, "Borrar unión seleccionada")
        oTT.SetToolTip(Me.BtnSeleccionar, "Seleccionar bloque UNION en Dibujo")
        oTT.SetToolTip(Me.BtnBorrarUnion, "Mostra/Ocultar capa UNIONES")
        oTT.SetToolTip(Me.BtnInsertaMultiplesUniones, "Insertar múltiples bloques '" & NombreBloqueUNION & "'")
        '
        oTT.SetToolTip(Me.BtnT1, "Seleccionar Transportador 1")
        oTT.SetToolTip(Me.BtnT2, "Seleccionar Transportador 2")
        oTT.SetToolTip(Me.BtnInsertarUnion, "Insertar o Mover Unión")
        oTT.SetToolTip(Me.btnCerrar, "Cerrar Uniones")
        oTT.SetToolTip(Me.BtnAceptar, "Aceptar unión y escribir atributos")
    End Sub

    'Public Sub PonUnion()
    '    If GUnion.Enabled = False Then Exit Sub
    '    '
    '    LbUnion.Items.Clear()
    '    Dim angulo As String = ""
    '    If BtnInsertarUnion.Visible = True Then
    '        If (clsA.Bloque_LargoEnX(UltimoBloqueT1.ObjectID) = True AndAlso clsA.Bloque_LargoEnX(UltimoBloqueT2.ObjectID) = True) OrElse
    '        (clsA.Bloque_LargoEnX(UltimoBloqueT1.ObjectID) = False AndAlso clsA.Bloque_LargoEnX(UltimoBloqueT2.ObjectID) = False) Then
    '            angulo = "0"
    '            LbRotation.SelectedIndex = 0
    '        Else
    '            angulo = "90"
    '            LbRotation.SelectedIndex = 1
    '        End If
    '    Else
    '        If UltimaClsUnion.ROTATION = "" OrElse UltimaClsUnion.ROTATION = "0" Then
    '            LbRotation.SelectedIndex = 0
    '        ElseIf UltimaClsUnion.ROTATION = "90" Then
    '            LbRotation.SelectedIndex = 1
    '        Else
    '            LbRotation.SelectedIndex = -1
    '        End If
    '    End If
    '    UltimaFilaExcel = cU.Fila_BuscaDame(UltimoBloqueT1.EffectiveName.Split("_"c)(0), LbInclinationT1.Text, UltimoBloqueT2.EffectiveName.Split("_"c)(0), LbInclinationT2.Text, angulo)
    '    If UltimaFilaExcel IsNot Nothing Then
    '        LbUnion.Tag = UltimaFilaExcel
    '        Dim valores As String = UltimaFilaExcel.UNION
    '        valores = valores.Replace(" o ", ";")
    '        valores = valores.Replace(" ", "")
    '        Dim partes() As String = valores.Split(";")
    '        LbUnion.Items.AddRange(partes)
    '        'If CbUnion.Items.Contains(CbUnion.Text) = False Then CbUnion.Text = ""
    '        If LbUnion.Items.Count >= 1 AndAlso LbUnion.Items.Contains(UltimaClsUnion.UNION) Then
    '            ListBox_SeleccionaPorTexto(LbUnion, UltimaClsUnion.UNION)
    '        Else
    '            LbUnion.SelectedIndex = -1
    '            'CbUnion.Text = ""
    '        End If
    '    End If
    'End Sub
    Private Sub CompruebaDatos()
        ' Salir, si no esta activo GUnion
        If GUnion.Enabled = False Then Exit Sub
        If EsUnionNueva = True Then

        ElseIf EsUnionNueva = False AndAlso UltimaClsUnion IsNot Nothing Then

        End If
        ' Antiguo PonUnion
        LbUnion.Items.Clear()
        LbUnion.SelectedIndex = -1
        LbUnion.Text = ""
        Dim angulo As String = ""
        If BtnInsertarUnion.Visible = True Then
            If (clsA.Bloque_LargoEnX(UltimoBloqueT1.ObjectID) = True AndAlso clsA.Bloque_LargoEnX(UltimoBloqueT2.ObjectID) = True) OrElse
            (clsA.Bloque_LargoEnX(UltimoBloqueT1.ObjectID) = False AndAlso clsA.Bloque_LargoEnX(UltimoBloqueT2.ObjectID) = False) Then
                angulo = "0"
                LbRotation.SelectedIndex = 0
            Else
                angulo = "90"
                LbRotation.SelectedIndex = 1
            End If
        Else
            If UltimaClsUnion.ROTATION = "" OrElse UltimaClsUnion.ROTATION = "0" Then
                LbRotation.SelectedIndex = 0
            ElseIf UltimaClsUnion.ROTATION = "90" Then
                LbRotation.SelectedIndex = 1
            Else
                LbRotation.SelectedIndex = -1
            End If
        End If
        UltimaFilaExcel = cU.Fila_BuscaDame(UltimoBloqueT1.EffectiveName.Split("_"c)(0), LbInclinationT1.Text, UltimoBloqueT2.EffectiveName.Split("_"c)(0), LbInclinationT2.Text, LbRotation.Text)
        If UltimaFilaExcel IsNot Nothing Then
            LbUnion.Tag = UltimaFilaExcel
            Dim valores As String = UltimaFilaExcel.UNION
            valores = valores.Replace(" o ", ";")
            valores = valores.Replace(" ", "")
            Dim partes() As String = valores.Split(";")
            LbUnion.Items.AddRange(partes)
            'If CbUnion.Items.Contains(CbUnion.Text) = False Then CbUnion.Text = ""
            If LbUnion.Items.Count >= 1 AndAlso LbUnion.Items.Contains(UltimaClsUnion.UNION) Then
                ListBox_SeleccionaPorTexto(LbUnion, UltimaClsUnion.UNION)
            Else
                LbUnion.SelectedIndex = -1
                'CbUnion.Text = ""
            End If
        End If
        ' ****************************************************
        '
        ' Salir, si no esta activo GUnion
        If GUnion.Enabled = False Then Exit Sub
        '
        Me.SuspendLayout()
        BtnInsertarUnion.Enabled = False
        BtnAceptar.Enabled = False
        '
        ' Comprobar objetos. BtnT1 y BtnT2 siempre activo
        BtnT1.Enabled = True
        BtnT2.Enabled = True
        ' *** BtnT1
        If UltimoBloqueT1 IsNot Nothing Then
            'LblT1.Text = "Datos T1:" & vbCrLf & UltimoBloqueT1.EffectiveName
            BtnT1.BackColor = btnOn
        Else
            LblT1.Text = "Datos T1:"
            BtnT1.BackColor = btnOff
        End If
        ' *** BtnT2
        If UltimoBloqueT2 IsNot Nothing Then
            'LblT2.Text = "Datos T2:" & vbCrLf & UltimoBloqueT2.EffectiveName
            BtnT2.BackColor = btnOn
        Else
            LblT2.Text = "Datos T2:"
            BtnT2.BackColor = btnOff
        End If
        '
        PonEstadoBotonesInsertarAceptar()
        '****************************************
        ' Deshabilitar controles, por defecto.
        'pb1.BackColor = btnOff 'Drawing.Color.FromArgb(255, 192, 192)
        'pb2.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'btnAgregarUnion.BackColor = Drawing.Color.FromArgb(255, 192, 192)
        'pb2.Enabled = False ': Control_Borde(pb2)
        'btnAgregarUnion.Enabled = False ': Control_Borde(btnAgregarUnion)
        'pb1.BackColor = Drawing.Color.FromArgb(255, 192, 192)
    End Sub

    Public Sub ExcelFila_Actualizar()
        If GUnion.Enabled = False Then Exit Sub
        ' Si no tenemos estos datos, salir. No se puede buscar nueva fila Excel.
        If UltimoBloqueT1 Is Nothing OrElse
                UltimoBloqueT2 Is Nothing OrElse
                LbInclinationT1.SelectedIndex = -1 OrElse
                LbInclinationT2.SelectedIndex = -1 OrElse
                LbRotation.SelectedIndex = -1 Then
            Exit Sub
        End If

        UltimaFilaExcel = cU.Fila_BuscaDame(UltimoBloqueT1.EffectiveName.Split("_"c)(0), LbInclinationT1.Text, UltimoBloqueT2.EffectiveName.Split("_"c)(0), LbInclinationT2.Text, LbRotation.Text)
        If UltimaFilaExcel IsNot Nothing Then
            If UltimaClsUnion IsNot Nothing Then
                UltimaClsUnion.ExcelFilaUnion = UltimaFilaExcel
            End If
            LbUnion.Tag = UltimaFilaExcel
            Dim valores As String = UltimaFilaExcel.UNION
            valores = valores.Replace(" o ", ";")
            valores = valores.Replace(" ", "")
            Dim partes() As String = valores.Split(";")
            LbUnion.Items.AddRange(partes)
            'If CbUnion.Items.Contains(CbUnion.Text) = False Then CbUnion.Text = ""
            If LbUnion.Items.Count >= 1 AndAlso LbUnion.Items.Contains(UltimaClsUnion.UNION) Then
                ListBox_SeleccionaPorTexto(LbUnion, UltimaClsUnion.UNION)
            Else
                LbUnion.SelectedIndex = -1
                'CbUnion.Text = ""
            End If
        End If
    End Sub
    Public Sub PonEstadoBotonesInsertarAceptar()
        Me.SuspendLayout()
        ' *** BtnInsertarUnion y BtnAceptar
        If UltimoBloqueT1 IsNot Nothing AndAlso
                UltimoBloqueT2 IsNot Nothing AndAlso
                LbInclinationT1.SelectedIndex > -1 AndAlso
                LbInclinationT2.SelectedIndex > -1 AndAlso
                LbRotation.SelectedIndex > -1 AndAlso
                LbUnion.SelectedIndex > -1 AndAlso
                LblUnits.Text <> "" Then
            BtnInsertarUnion.Enabled = True
            BtnAceptar.Enabled = True
        Else
            BtnInsertarUnion.Enabled = False
            BtnAceptar.Enabled = False
        End If
        Me.ResumeLayout()
    End Sub
    Public Sub tvUniones_Rellena(Optional tipo As modTavil.EFiltro = EFiltro.TODOS)
        PonEstadoControlesInicial()
        ' Rellenar tvGrupos con los grupos que haya ([nombre grupo]) Sacado de XData elementos (regAPPCliente, XData = "GRUPO")
        tvUniones.Nodes.Clear()
        tvUniones.BeginUpdate()
        Dim arrTodos As List(Of TreeNode) = modTavil.tvUniones_DameListTreeNodes
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            tvUniones.EndUpdate()
            Exit Sub
        End If
        '
        tvUniones.Nodes.AddRange(arrTodos.ToArray)
        tvUniones.Sort()
        tvUniones.SelectedNode = Nothing
        tvUniones_AfterSelect(Nothing, Nothing)
        BtnSeleccionar.Enabled = (tvUniones.Nodes.Count > 0)
        cbCapa.Enabled = (tvUniones.Nodes.Count > 0)
        LblUniones.Text = tvUniones.Nodes.Count & " Uniones"
        tvUniones.EndUpdate()
    End Sub

    Public Sub Uniones_SeleccionarObjetos(HandleUnion As String, Optional conZoom As Boolean = False)
        Dim lGrupos As New List(Of Long)
        Try
            UltimoBloqueUnion = Eventos.COMDoc().HandleToObject(HandleUnion)
            lGrupos.Add(UltimoBloqueUnion.ObjectID)
        Catch ex As Exception
            Exit Sub
        End Try
        '
        Dim idT1 As String = clsA.XLeeDato(UltimoBloqueUnion.Handle, "T1HANDLE")
        Dim idT2 As String = clsA.XLeeDato(UltimoBloqueUnion.Handle, "T2HANDLE")
        If idT1 <> "" Then
            Try
                UltimoBloqueT1 = Eventos.COMDoc().HandleToObject(idT1)
                lGrupos.Add(UltimoBloqueT1.ObjectID)
            Catch ex As Exception
            End Try
        End If
        If idT2 <> "" Then
            Try
                UltimoBloqueT2 = Eventos.COMDoc().HandleToObject(idT2)
                lGrupos.Add(UltimoBloqueT2.ObjectID)
            Catch ex As Exception
            End Try
        End If

        If lGrupos.Count > 0 Then
            If cbZoom.Checked OrElse conZoom = True Then
                clsA.Selecciona_AcadID(lGrupos.ToArray)
                clsA.HazZoomObjeto(UltimoBloqueUnion, 3, False)
            Else
                clsA.Selecciona_AcadID(lGrupos.ToArray)
            End If
        End If
    End Sub

    ' PictureBox (U otros controles). Poner el borde de color rojo  
    Private Sub Control_Borde(oC As Control, Optional restaurar As Boolean = False)
        ' Restaurar botón.
        If restaurar = True Then
            oC.Invalidate()
            Exit Sub
        End If
        ' Poner borde Rojo
        If TypeOf oC Is PictureBox Then
            CType(oC, PictureBox).BorderStyle = BorderStyle.None
        End If
        'Rectangulo del control + Offset del rectangulo hacia fuera.  
        Dim BorderBounds As Rectangle = oC.ClientRectangle
        'BorderBounds.Inflate(-1, -1)

        'Use ControlPaint to draw the border.  
        'Change the Color.Red parameter to your own colour below.  
        ControlPaint.DrawBorder(oC.CreateGraphics,
                                        BorderBounds,
                                        Color.Red,
                                        ButtonBorderStyle.Outset)
    End Sub
    '
    Public Sub BorraDatos()
        Me.LbInclinationT1.SelectedIndex = -1
        Me.LbInclinationT2.SelectedIndex = -1
        Me.LbRotation.SelectedIndex = -1
        Me.LbUnion.SelectedIndex = -1
        Me.LblUnits.Text = ""
    End Sub
#End Region
End Class
'
'Efecto Click en PictureBox. Click (Borde Rojo), otro click (Sin borde)  
'Private Sub pb1_Click(sender As Object, e As EventArgs) Handles pb1.Click, pb2.Click 'etc  
'    Dim pbX As PictureBox = DirectCast(sender, PictureBox)
'    'Get the rectangle of the control and inflate it to represent the border area  
'    Dim BorderBounds As Rectangle = pbX.ClientRectangle
'    BorderBounds.Inflate(-1, -1)

'    'Use ControlPaint to draw the border.  
'    'Change the Color.Red parameter to your own colour below.  
'    ControlPaint.DrawBorder(pbX.CreateGraphics,
'                            BorderBounds,
'                            Color.Red,
'                            ButtonBorderStyle.Solid)

'    If Not (HighlightedPictureBox Is Nothing) Then
'        'Remove the border of the last PictureBox  
'        HighlightedPictureBox.Invalidate()
'    End If

'    'Rememeber the last highlighted PictureBox  
'    HighlightedPictureBox = CType(sender, PictureBox)
'    pbX = Nothing
'End Sub
