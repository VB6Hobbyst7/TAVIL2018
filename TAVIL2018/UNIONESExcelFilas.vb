﻿Imports System.Linq
Imports ce = ClosedXML.Excel
Imports System.Windows.Forms

Public Class UNIONESExcelFilas
    Public nHoja As String = HojaUniones
    Public Campos As List(Of String)
    Public filas As List(Of UNIONESExcelFila)        ' Key = ITEM_NUMBER, Value = clsPTItem
    Public Sub New()
        If cfg Is Nothing Then cfg = New UtilesAlberto.Conf(System.Reflection.Assembly.GetExecutingAssembly)
        'If cXML Is Nothing Then cXML = New ClosedXML2acad.ClosedXML2acad
        Campos = New List(Of String)
        filas = New List(Of UNIONESExcelFila)
        '
        For Each fila As ce.IXLRow In modClosedXMLTavil.Excel_LeeFilar(LAYOUTDB, nHoja, concabeceras:=True).AsParallel
            If fila.RowNumber = 1 Then
                For Each oCe As ce.IXLCell In fila.CellsUsed
                    Campos.Add(Convert.ToString(oCe.Value))
                Next
            Else
                filas.Add(New UNIONESExcelFila(fila))
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
    End Sub
    ' inC y outC siempre enviarlos con 8 carácteres (Pueden ser TRD300XX o TRD30015)
    ' Debemos buscar el que sea único, con las XX y sin las XX
    'Public Function Fila_BuscaDame(inC As String, inInc As String, outC As String, outInc As String, angle As String) As UNIONESExcelFila
    '    Dim resultado As UNIONESExcelFila = Nothing
    '    If angle = "0" Then angle = ""
    '    Dim fila = From x In filas
    '               Where x.INFEED_CONVEYOR.Trim.StartsWith(inC.Trim) AndAlso
    '                                x.INFEED_INCLINATION.Trim = inInc.Trim AndAlso
    '                                x.OUTFEED_CONVEYOR.Trim.StartsWith(outC.Trim) AndAlso
    '                                x.OUTFEED_INCLINATION.Trim = outInc.Trim AndAlso
    '                                x.ANGLE.Trim = angle.Trim
    '    '
    '    If fila.Count > 0 Then
    '        resultado = fila.FirstOrDefault
    '    End If
    '    '
    '    Return resultado
    'End Function

    ' inC y outC siempre enviarlos con 8 carácteres (Pueden ser TRD300XX o TRD30015)
    ' Debemos buscar el que sea único, con las XX y sin las XX
    Public Function Fila_BuscaDame(inC As String, inInc As String, outC As String, outInc As String, angle As String) As UNIONESExcelFila
        Dim resultado As UNIONESExcelFila = Nothing
        Dim inCFin As String = ""
        Dim outCFin As String = ""
        '
        ' Filtro 1 (para ver si existe inC con 8 carácteres o con 6 y terminado en XX)
        Dim filtros() As String = {inC, inC.Substring(0, 6) & "XX"}
        For Each filtro As String In filtros
            Dim f1 As IEnumerable(Of UNIONESExcelFila) = From x In filas
                                                         Where x.INFEED_CONVEYOR.ToUpper.Trim = filtro
            If f1 IsNot Nothing AndAlso f1.Count > 0 Then
                inCFin = filtro
                Exit For
            End If
        Next
        ' Filtro 2 (para ver si existe outC con 8 carácteres o con 6 y terminado en XX)
        filtros = {outC, outC.Substring(0, 6) & "XX"}
        For Each filtro As String In filtros
            Dim f1 As IEnumerable(Of UNIONESExcelFila) = From x In filas
                                                         Where x.INFEED_CONVEYOR.ToUpper.Trim = filtro
            If f1 IsNot Nothing AndAlso f1.Count > 0 Then
                outCFin = filtro
                Exit For
            End If
        Next
        '
        If inCFin = "" OrElse outCFin = "" Then
            ' Salimos con Nothing
            Return resultado
            'Exit Function
        End If
        '
        '
        If angle = "0" Then angle = ""
        Dim fila = From x In filas
                   Where x.INFEED_CONVEYOR.Trim.StartsWith(inCFin) AndAlso
                                    x.INFEED_INCLINATION.Trim = inInc.Trim AndAlso
                                    x.OUTFEED_CONVEYOR.Trim.StartsWith(outCFin.Trim) AndAlso
                                    x.OUTFEED_INCLINATION.Trim = outInc.Trim AndAlso
                                    x.ANGLE.Trim = angle.Trim
        '
        If fila.Count > 0 Then
            resultado = fila.FirstOrDefault
        End If
        '
        Return resultado
    End Function
End Class


Public Class UNIONESExcelFila
    ' TRD300XX	FLAT	132353 o 132348; 120291	1; 1	TRD30305	FLAT		RR     Primer transportador/union Recto. Segundo transportador/union Recto
    Public nFila As Integer = -1
    Public INFEED_CONVEYOR As String = ""           ' TRD300XX
    Public INFEED_INCLINATION As String = ""        ' FLAT
    Private _UNION As String = ""                   ' 132353 o 132348; 120291
    Private _UNITS As String = ""                   ' 1; 1
    Public OUTFEED_CONVEYOR As String = ""          ' TRD30305
    Public OUTFEED_INCLINATION As String = ""       ' FLAT
    Public ANGLE As String = ""                     ' "" o 90
    Public INFORMACION As String = ""               ' RR     Primer transportador/union Recto. Segundo transportador/union Recto
    Public Rows As List(Of DataGridViewRow)
    Public hayerror As Boolean = False

    Public Property UNION As String
        Get
            Return _UNION
        End Get
        Set(value As String)
            _UNION = value.Replace(" ", "").Trim    ' value.Replace("o", ";").Replace(" ", "").Trim
        End Set
    End Property

    Public Property UNITS As String
        Get
            Return _UNITS
        End Get
        Set(value As String)
            _UNITS = value.Replace(" ", "").Trim
        End Set
    End Property

    Public Sub New(fila As ce.IXLRow)
        'If cXML Is Nothing Then cXML = New ClosedXML2acad.ClosedXML2acad
        nFila = fila.RowNumber
        For Each oCell As ce.IXLCell In fila.Cells.AsParallel
            Dim cabecera As String = oCell.WorksheetColumn.FirstCell.Value.ToString.Trim
            Dim valor As String = oCell.Value
            Select Case cabecera.ToUpper
                Case "INFEED_CONVEYOR" : INFEED_CONVEYOR = Convert.ToString(valor).Trim
                Case "INFEED_INCLINATION" : Me.INFEED_INCLINATION = Convert.ToString(valor).Trim
                Case "UNION" : Me.UNION = Convert.ToString(valor).Trim.Replace(" ", "")
                Case "UNITS" : Me.UNITS = Convert.ToString(valor).Trim.Replace(" ", "")
                Case "OUTFEED_CONVEYOR" : Me.OUTFEED_CONVEYOR = Convert.ToString(valor).Trim
                Case "OUTFEED_INCLINATION" : Me.OUTFEED_INCLINATION = Convert.ToString(valor).Trim
                Case "ANGLE" : Me.ANGLE = Convert.ToString(valor).Trim
                Case "INFORMACION" : Me.INFORMACION = Convert.ToString(valor).Trim
            End Select
        Next
        FilasDataGridView_Crea()
    End Sub
    Public Sub FilasDataGridView_Crea()
        Dim unionesP() As String = Me.UNION.Split(";")
        Dim unidadesP() As String = Me.UNITS.Split(";")
        If unionesP.Count <> unidadesP.Count Then
            MsgBox("Error in Excel. Columns UNION and UNITS")
            hayerror = True
            Exit Sub
        Else
            hayerror = False
        End If
        '
        Rows = New List(Of DataGridViewRow)
        '
        For x As Integer = 0 To unionesP.Count - 1
            Dim Fila As New DataGridViewRow
            Dim Ctext As New DataGridViewComboBoxCell
            Dim Ttext As New DataGridViewTextBoxCell
            Dim TUnits As New DataGridViewTextBoxCell
            '
            TUnits.Value = unidadesP(x)
            If unionesP(x).Contains("o") Then
                Dim partes() As String = unionesP(x).Split("o")
                Ctext.Items.AddRange(partes)
                Ctext.Value = partes(0)
                Fila.Cells.Add(Ctext)
            Else
                Ttext.Value = unionesP(x)
                Fila.Cells.Add(Ttext)
            End If
            Fila.Cells.Add(TUnits)
            Rows.Add(Fila)
            'TUnits = Nothing
            'Ctext = Nothing
            'Ttext = Nothing
            'Fila = Nothing
        Next
    End Sub
End Class

Public Enum nombreColumnaUNIONES
    INFEED_CONVEYOR
    INFEED_INCLINATION
    UNION
    UNITS
    OUTFEED_CONVEYOR
    OUTFEED_INCLINATION
    ANGLE
    INFORMACION
End Enum
