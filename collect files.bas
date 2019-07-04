'-----------------------------------------------------------[date: 2019.07.04]
Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' 2019.07.03(水).new
' 2019.07.04(木)
'***********************************************
Public fc As C0path
'***********************************************
Private Enum numbers
  ZERO = 0
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
  hpj_id = 1
  hpj_sheet_name = 2
  hpj_y_excel = 3
  hpj_x_excel = 4
  hpj_title = 5
  hpj_title_sub = 6
  hpj_type = 7
  hpj_not_use = 8
  hpj_RA_num = 9
  hpj_key = 10
  hpj_x_text = 11
  hpj_size = 12
  hpj_cts = 13
  hpj_limit = 14
  hpj_APF_comments = 15
End Enum
'***********************************************

Public Sub main()
  Dim py_fa         As Long
  Dim ws_fa         As Worksheet
  Dim wb_fa         As Workbook
  Dim sample_id     As Long
  Dim elog          As c0_Log
  Dim out_line_size As Long
  Dim ff_out        As Long
  Dim filename_out  As String
  Dim buf           As String
  Dim col_in        As Collection
  Dim ws            As Worksheet
  Dim pj            As Variant
  Dim ys_pj         As Long
  Dim ii            As Variant
  Set elog = New c0_Log
  Set ws = ThisWorkbook.Worksheets("section")
  pj = ws.UsedRange.Value
  ys_pj = UBound(pj, N1)
  out_line_size = text_size(pj, ys_pj)
  Set fc = New C0path
  Set col_in = New Collection
  Set wb_fa = Workbooks.Add
  Set ws_fa = wb_fa.Worksheets(N1)
  Call setup_ws_fa(ws_fa, pj)
  ws_fa.Name = "自由記述"
  For Each ii In fc.col_fn_in_sub
    If ii Like "*xls*" Then
      py_fa = col_in.Count + N2
      buf = data_text(CStr(ii), pj, ys_pj, out_line_size, ws_fa, py_fa)
      sample_id = 10001 + col_in.Count
      Mid(buf, N1, N5) = CStr(sample_id)
      col_in.Add buf
      ws_fa.Cells(py_fa, N1).Value = sample_id
      elog.pp sample_id, ii
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
  elog.putf ("sample id")
  wb_fa.SaveAs filename:=fc.cur & "result FA.xlsx"
  wb_fa.Close savechanges:=False
  MsgBox "end of run" & vbCrLf & "in sample count: " & CStr(col_in.Count)
End Sub
'-----------------------------------------------------------------------------

Private Function data_text(a_filename As String, pj As Variant, ys_pj As Long, out_line_size As Long, ws_fa As Worksheet, py_fa As Long) As String
  Dim ar()      As Boolean
  Dim t_type    As String
  Dim buf       As String
  Dim wb        As Workbook
  Dim ws        As Worksheet
  Dim ol        As Variant
  Dim tt_buf    As String
  Dim tt_size   As Long
  Dim tt        As Long
  Dim ty        As Long
  Dim tx        As Long
  Dim j         As Long
  Dim i         As Long
  Dim y         As Long
  Dim now_sheet As String
  Set wb = Workbooks.Open(filename:=a_filename, Password:="r01htsny")
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
      Select Case t_type
        Case "S", "M", "L"
          For i = N1 To pj(y, hpj_limit)
            ' typeの選別から？
            tt_buf = Trim(CStr(ol(ty, tx + i - N1)))
            If Len(tt_buf) Then
              tt = Val(tt_buf)
              If tt Then
                tt_size = (Len(tt_buf) - N1) \ pj(y, hpj_size)
                If tt_size Then
                  If Len(tt_buf) Mod pj(y, hpj_size) Then
                    tt_buf = "0" & tt_buf
                  End If
                  For j = N1 To tt_size + N1
                    tt = Val(Mid(tt_buf, (j - N1) * pj(y, hpj_size) + N1, pj(y, hpj_size)))
                    If tt > pj(y, hpj_cts) Then
                      MsgBox "range over"
                    Else
                      If tt Then
                        ar(tt) = True
                      End If
                    End If
                  Next j
                  ' 出力しなきゃ
                Else
                  If tt Then
                    ar(tt) = True
                  End If
                End If
              Else
                If UCase(tt_buf) = "TRUE" Then
                  ar(i) = True
                End If
              End If
            End If
          Next i
          Mid(buf, pj(y, hpj_x_text), pj(y, hpj_size) * pj(y, hpj_limit)) = to_s(ar)
        Case "R"
          tt_buf = Trim(CStr(ol(ty, tx)))
          If Len(tt_buf) Then
            tt = Val(ol(ty, tx))
            Mid(buf, pj(y, hpj_x_text), pj(y, hpj_size)) = rbuf(tt, Val(pj(y, hpj_size)))
          End If
        Case "F"
          ws_fa.Cells(py_fa, pj(y, hpj_x_text) + N1).Value = ol(ty, tx)
        Case Else
      End Select
    End If
  Next y
  wb.Close savechanges:=False
  data_text = buf
End Function
'-----------------------------------------------------------------------------

Private Function text_size(pj As Variant, ys_pj As Long) As Long
  Dim max_position As Long
  Dim t_end        As Long
  Dim y            As Long
  For y = N2 To ys_pj
    t_end = pj(y, hpj_x_text) + (pj(y, hpj_size) * pj(y, hpj_limit)) - N1
    If max_position < t_end Then
      max_position = t_end
    End If
  Next y
  text_size = (Int((max_position - N1) / 100) + N1) * 100
End Function
'-----------------------------------------------------------------------------

Private Function rbuf(ByVal a_buf As Variant, Optional a_size As Long = 5) As String
  If Len(Trim(a_buf)) Then
    rbuf = Right(Space(a_size) & a_buf, a_size)
  End If
End Function
'-----------------------------------------------------------------------------

Private Function to_s(ar As Variant) As String
  Dim buf  As String
  Dim i    As Long
  Dim cts  As Long
  Dim size As Long
  If IsArray(ar) Then
    cts = UBound(ar)
    If cts >= 10 Then
      size = N2
    Else
      size = N1
    End If
    For i = LBound(ar) To cts
      If ar(i) Then
        buf = buf & Format(i, String(size, "0"))
      End If
    Next i
  End If
  to_s = buf
End Function
'-----------------------------------------------------------------------------

Private Sub setup_ws_fa(ws_fa As Worksheet, pj As Variant)
  Dim y  As Long
  Dim cn As Long
  ws_fa.Cells(N1, N1).Value = "SampleNo."
  For y = N2 To UBound(pj, N1)
    If pj(y, hpj_type) = "F" Then
      cn = cn + N1
      ws_fa.Cells(N1, cn + N1).Value = pj(y, hpj_title)
    End If
  Next y
End Sub
'-----------------------------------------------------------------------------

