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

    Public Function Grupo_SeleccionarObjetos(grupo As String, conzoom As Boolean) As ArrayList
        Dim resultado As New ArrayList
        Dim lGrupos As List(Of Long) = clsA.SeleccionaTodosObjetosXData(cGRUPO, grupo)
        If lGrupos IsNot Nothing AndAlso lGrupos.Count > 0 Then
            Dim arrSeleccion As New ArrayList
            For Each queId As Long In lGrupos
                arrSeleccion.Add(Eventos.COMDoc().ObjectIdToObject(queId))
            Next
            If arrSeleccion.Count > 0 Then
                Eventos.COMDoc.ActiveSelectionSet.Clear()
                clsA.SeleccionCreaResalta(arrSeleccion, 2000, conzoom)
            End If
        End If
        Return resultado
    End Function

    Public Sub Report_Blocks(lh As List(Of String), sufijo As String)
        Dim columnas() As String = {"BLOCK", "CODE", "COUNT", "UNION", "UNITS"}
        Dim fiOut As String = IO.Path.ChangeExtension(Eventos.COMDoc.Path, sufijo & ".csv")
        If IO.File.Exists(fiOut) Then IO.File.Delete(fiOut)
        Dim lineas As New Dictionary(Of String, bloque)
        '
        Dim texto As String = String.Join(";", columnas) & vbCrLf
        For Each h As String In lh
            Dim obj As AcadObject = Eventos.COMDoc.HandleToObject(h)
            If TypeOf obj Is AcadBlockReference = False Then
                Continue For
            End If
            '
            Dim oBl As AcadBlockReference = CType(obj, AcadBlockReference)
            Dim name As String = oBl.EffectiveName
            If lineas.ContainsKey(name) Then
                lineas(name).count += 1
            Else
                Dim linea As New bloque
                linea.name = oBl.EffectiveName
                linea.code = ""
                'linea.count = 1
                lineas.Add(name, linea)
                linea = Nothing
            End If
        Next
        For Each n As String In lineas.Keys
            texto &= lineas(n).valores & vbCrLf
        Next
        texto = texto.Substring(0, texto.Length - 2)
        IO.File.WriteAllText(fiOut, texto, Text.Encoding.UTF8)
        If IO.File.Exists(fiOut) Then Process.Start(fiOut)
    End Sub
    Public Class bloque
        Public name As String = ""
        Public code As String = ""
        Public count As Integer = 1
        Public Function valores() As String
            Return name & ";" & code & ";" & count
        End Function
    End Class

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