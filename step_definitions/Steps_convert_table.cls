VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_convert_table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_excel_sheet_with_a_table_9C2B567D5280(example_context As TContext, data_table As TDataTable)
    'Given an Excel sheet with a table
    Dim test_workbook As Workbook
    Dim xl_sheet As Worksheet
    
    Set test_workbook = Application.Workbooks.Add
    Do While test_workbook.Sheets.Count > 1
        test_workbook.Sheets(test_workbook.Sheets.Count).Delete
    Loop
    Set xl_sheet = test_workbook.Sheets(1)
    example_context.set_value xl_sheet, "xl_sheet"
    TSupport.write_data_table_to_sheet data_table, xl_sheet
End Sub

Public Sub When_the_table_is_converted_into_slides_5C8852845EF9(example_context As TContext)
    'When the table is converted into slides
    pending
End Sub

Public Sub Then_a_new_presentations_is_created_31D8BE63F567(example_context As TContext)
    'Then a new presentations is created
    pending
End Sub

Public Sub Then_the_presentation_has_INT_new_slides_923EE0B0FEBA(example_context As TContext, step_expressions As Collection)
    'And the presentation has {integer} new slides
    pending
End Sub

Public Sub Then_slide_INT_has_the_title_STR_439A0A32FA64(example_context As TContext, step_expressions As Collection)
    'And slide {integer} has the title {string}
    pending
End Sub

