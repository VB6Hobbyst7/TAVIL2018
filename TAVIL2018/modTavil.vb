Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports TAVIL2018.TAVIL2018
Imports System.Windows.Forms
Imports uau = UtilesAlberto.Utiles
Imports a2 = AutoCAD2acad.A2acad
Imports Ev = TAVIL2018.Eventos

Public Module modTavil
    Public LClsUnion As List(Of ClsUnion)
    Public Function tvUniones_DameListTreeNodes(Optional queFiltro As EFiltro = EFiltro.TODOS) As List(Of TreeNode)
        Dim lUni As New List(Of TreeNode)
        LClsUnion = New List(Of ClsUnion)
        '
        If clsA Is Nothing Then clsA = New a2.A2acad(Eventos.COMApp(), cfg._appFullPath, regAPPCliente)
        Dim arrTodos As List(Of String) = clsA.SeleccionaDameHandle_PorNombreBloque(NombreBloqueUNION)
        If arrTodos Is Nothing OrElse arrTodos.Count = 0 Then
            Return lUni
            Exit Function
        End If
        '
        ' Filtrar lista de grupo. Sacar nombres únicos.
        For Each queHandle As String In arrTodos
            Dim acadObj As AcadObject = Eventos.COMDoc().HandleToObject(queHandle)
            If TypeOf acadObj Is AcadBlockReference = False Then Continue For
            Dim oBl As AcadBlockReference = CType(acadObj, AcadBlockReference)
            If oBl.EffectiveName <> NombreBloqueUNION Then Continue For
            '
            Dim union As String = clsA.XLeeDato(oBl.Handle, "UNION")
            Dim oNode As New TreeNode
            oNode.Text = queHandle
            oNode.Name = queHandle
            oNode.Tag = queHandle
            oNode.ToolTipText = "Unión = " & oBl.EffectiveName & "·" & queHandle
            'oNode = Nothing

            '
            If queFiltro.ToString = EFiltro.TODOS.ToString Then
                lUni.Add(oNode)
                LClsUnion.Add(New ClsUnion(oBl.Handle))
            ElseIf queFiltro.ToString = EFiltro.XXX.ToString Then
                'Dim t1 As String = clsA.XLeeDato(oBl.Handle, "T1HANDLE")
                'Dim t2 As String = clsA.XLeeDato(oBl.Handle, "T2HANDLE")
                'If t1 <> "" AndAlso t2 <> "" Then
                '    Continue For
                'Else
                '    lUni.Add(oNode)
                'End If
                Dim Uni As String = clsA.XLeeDato(oBl.Handle, "UNION")
                If Uni = "" OrElse Uni = "XXX" Then
                    lUni.Add(oNode)
                    LClsUnion.Add(New ClsUnion(oBl.Handle))
                Else
                    Continue For
                End If
            End If
            oBl = Nothing
            oNode = Nothing
        Next
        '
        Return lUni
    End Function

    Public Enum EFiltro
        TODOS
        XXX
    End Enum
    Public Enum INCLINATION
        FLAT
        DOWN
        UP
    End Enum
End Module