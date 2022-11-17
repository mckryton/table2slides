Attribute VB_Name = "TSupport"
'this module is for sharing helper functions between your step implementation classes

Option Explicit

Public Sub write_data_table_to_sheet(data_table As TDataTable, target_sheet As Worksheet)

    Dim column_name As Variant
    Dim data_row As Variant
    Dim current_range As Range
    Dim table_top_left As Range
    
    Set table_top_left = target_sheet.Range("B2")
    Set current_range = table_top_left
    For Each column_name In data_table.column_names
        current_range.Value = "'" & CStr(column_name)
        For Each data_row In data_table.table_rows
            Set current_range = current_range.Offset(1)
            current_range.Value = "'" & data_row(column_name)
        Next
        Set current_range = table_top_left.Offset(, current_range.Row - 1)
    Next
End Sub
