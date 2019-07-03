'-----------------------------------------------------------[date: 2019.07.03]
Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' 2019.07.03(êÖ).new
'***********************************************
' https://github.com/acr002/vba_collect_files.git
' git@github.com:acr002/vba_collect_files.git
'***********************************************
Private Enum numbers
  zero = 0
  N1 = 1
  N2 = 2
  N3 = 3
  N4 = 4
  N5 = 5
  N6 = 6
  N7 = 7
  N8 = 8
  N9 = 9
End Enum
'***********************************************
Private Enum header_pj
  hpj_id           = 1
  hpj_Sheet_name   = 2
  hpj_y_excel      = 3
  hpj_x_excel      = 4
  hpj_title        = 5
  hpj_title_sub    = 6
  hpj_type         = 7
  hpj_not_use      = 8
  hpj_RA_num       = 9
  hpj_key          = 10
  hpj_x_text       = 11
  hpj_size         = 12
  hpj_cts          = 13
  hpj_limit        = 14
  hpj_APF_comments = 15
End Enum
'***********************************************

Public Sub main()
  Dim out_line_size As Long
  Dim ff_out       As Long
  Dim filename_out As String
  Dim col_in       As Collection
  Dim ws           As Worksheet
  Dim pj           As Variant
  Dim ys_pj        As Long
  Dim ii           As Variant
  Dim fc           As C0path
  Set ws = ThisWorkbook.Worksheets("section")
  pj = ws.UsedRange.Value
  ys_pj = UBound(pj, N1)
  out_line_size = text_size(pj, 10, 11, 13)
  Set fc = New C0path
  Set col_in = New Collection
  For Each ii In fc.col_fn_in_sub
    If ii Like "*xls*" Then
      col_in.Add data_text(CStr(ii), pj, ys_pj, out_line_size)
    End If
  Next ii
  If col_in.Count Then
    ff_out = FreeFile()
    filename_out = fc.cur & fc.jobcode & ".in"
    Open filename_out For Output As #ff_out
    For Each ii In col_in
      Print #ff_out, ii
    Next ii
    Close #ff_out
  End If
  MsgBox "end of run" & vbCrLf & "in sample count: " & CStr(col_in.Count)
End Sub
'-----------------------------------------------------------------------------

Private Function data_text(a_filename As String, pj As Variant, ys_pj As Long, out_line_size As Long) As String
  Dim ar()      As Boolean
  Dim t_type    As String
  Dim buf       As String
  Dim wb        As Workbook
  Dim ws        As Worksheet
  Dim ol        As Variant
  Dim j         As Long
  Dim i         As Long
  Dim y         As Long
  Dim now_sheet As String
  Set wb = Workbooks.Open(filename:=a_filename, password:="r01htsny")
  buf = Space(out_line_size)
  For y = N2 To ys_pj
    If Val(pj(y, hpj_x_text)) Then
      If now_sheet <> pj(y, hpj_sheet_name) Then
        now_sheet = pj(y, hpj_sheet_name)
        Set ws = wb.Worksheets(now_sheet)
        ol = ws.UsedRange.Value
      End If
      t_type = pj(y, hpj_type)
      Select Case t_type
        Case "M", "L", "S"
          ReDim ar(N1 To pj(y, hpj_cts))
        Case Else
      End Select
      ty = pj(y, hpj_y_excel)
      tx = pj(y, hpj_x_excel)
      select case t_type
        case "S", "M", "L", "R"
          For i = N1 To pj(y, hpj_limit)
            ' typeÇÃëIï Ç©ÇÁÅH
            tt_buf = cstr(ol(ty, tx + i - N1))
            tt = val(tt_buf)
            if tt then
              if tt > pj(y, hpj_cts) then
                for j = 
                '
              else
                ar(tt) = true
              end if
            else
              if ucase(tt_buf) = "TRUE" then
                ar(i) = true
              end if
            end if

            if len(tt_buf) > pj(y, hpj_size) then
            if len(tt_buf) mod pj(y, hpj_size) then
              tt_buf = "0" & tt_buf
            tt = val(ol(ty, tx + i - N1))
            if tt then
              if 


    End If
  Next y
End Function
'-----------------------------------------------------------------------------

Private Function text_size(pj As Variant, ys_pj As Long, x As Long, size As Long, limit As Long) As Long
  Dim max_position As Long
  Dim t_end As Long
  Dim y As Long
  For y = N2 To ys_pj
    t_end = pj(y, x) + (pj(y, size) * pj(y, limit)) - N1
    If max_position < t_end Then
      max_position = t_end
    End If
  Next y
  text_size = (Int((max_position - N1) / 100) + N1) * 100
End Function
'-----------------------------------------------------------------------------

