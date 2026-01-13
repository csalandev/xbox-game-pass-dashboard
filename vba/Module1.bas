'-----------------------------------------------
' Módulo1 - Period Control and PivotTable Management
'-----------------------------------------------

Option Explicit

' Updates the Dashboard state and applies filters to all PivotTables
Sub SelecionarPeriodo(periodo As String)
    Dim shp As Shape
 
 
    ' 1. Update control cell with the selected period
    Sheets("Dashboard").Range("Z1").Value = periodo
    
    ' 2. Update button visual states (selected vs. unselected)
    For Each shp In Sheets("Dashboard").Shapes
        Select Case shp.Name
            Case "btnTotal", "btnMonthly", "btnQuarterly", "btnAnnual"
                If shp.Name = "btn" & periodo Then
                    ' Active button style
                    shp.Fill.ForeColor.RGB = RGB(42, 230, 177)
                Else
                    ' Inactive button style
                    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    shp.Line.Visible = msoTrue
                    shp.Line.ForeColor.RGB = RGB(42, 230, 177)
                    shp.Line.Weight = 0.2
                End If
        End Select
    Next shp
    
    ' 3. Apply period filter to all PivotTables
    FiltrarTabelasPorPeriodo periodo
    
    ' 4. Force full recalculation of KPI values
    Application.CalculateFull
    DoEvents ' Allows the UI thread to complete screen updates
End Sub

' Applies the selected period filter to all PivotTables on the "Cálculos" worksheet
Sub FiltrarTabelasPorPeriodo(periodo As String)

    Dim wsTD As Worksheet
    Dim pt As PivotTable

    Set wsTD = ThisWorkbook.Sheets("Cálculos")

    ' 1st PivotTable: tbl_autorenewal_total
    Set pt = wsTD.PivotTables("tbl_autorenewal_total")
    If periodo = "Total" Then
        pt.PivotFields("Subscription Type").ClearAllFilters
    Else
        pt.PivotFields("Subscription Type").CurrentPage = periodo
    End If
    
    pt.RefreshTable

    ' 2nd PivotTable: tbl_easeasonpass_total
    Set pt = wsTD.PivotTables("tbl_easeasonpass_total")
    If periodo = "Total" Then
        pt.PivotFields("Subscription Type").ClearAllFilters
    Else
        pt.PivotFields("Subscription Type").CurrentPage = periodo
    End If
    
    pt.RefreshTable

    ' 3rd PivotTable: tbl_minecraftpass_total
    Set pt = wsTD.PivotTables("tbl_minecraftpass_total")
    If periodo = "Total" Then
        pt.PivotFields("Subscription Type").ClearAllFilters
    Else
        pt.PivotFields("Subscription Type").CurrentPage = periodo
    End If
    
    pt.RefreshTable

    ' 4th PivotTable: tbl_xboxgamepass_total
    Set pt = wsTD.PivotTables("tbl_xboxgamepass_total")
    If periodo = "Total" Then
        pt.PivotFields("Subscription Type").ClearAllFilters
    Else
        pt.PivotFields("Subscription Type").CurrentPage = periodo
    End If
    
    pt.RefreshTable

End Sub


'-----------------------------------------------
' Dashboard button callback macros
'-----------------------------------------------

Sub btnMonthly_Click()
    SelecionarPeriodo "Monthly"
End Sub

Sub btnQuarterly_Click()
    SelecionarPeriodo "Quarterly"
End Sub

Sub btnAnnual_Click()
    SelecionarPeriodo "Annual"
End Sub
Sub btnTotal_Click()
    SelecionarPeriodo "Total"
End Sub