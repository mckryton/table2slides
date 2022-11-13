Attribute VB_Name = "TRun"
'This is the starting point for all test runs. TRun.test will allow you to optional add tags and
' report formats while TRun.wip will call test with the @wip (work in progress) tag using the
' the verbose report format by default.

Option Explicit

Private m_step_implementations As Collection

Public Sub test(Optional filter_tag, Optional feature_filter, Optional report_format)

    Dim session As TSession
    
    Set session = THelper.new_TSession
    session.run_test StepImplementations, filter_tag, feature_filter, report_format, application_dir:=table2slides_steps_workbook.path
    Set session = Nothing
    Set m_step_implementations = Nothing
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip", report_format:="verbose"
End Sub

Public Sub progress(Optional filter_tag)
    test filter_tag, report_format:="progress"
End Sub

Private Property Get StepImplementations() As Collection
    
    Dim step_implementations As Variant
    Dim step_implementation_class As Variant

    Set m_step_implementations = New Collection
    'REGISTER all classes with STEP IMPLEMENTATIONS HERE >>>
    '-------------------------------------------------------
    step_implementations = Array()
    Set m_step_implementations = New Collection
    For Each step_implementation_class In step_implementations
        m_step_implementations.Add step_implementation_class
    Next
    Set StepImplementations = m_step_implementations
End Property
