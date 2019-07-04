VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c0_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------[date: 2018.11.07]
Option Explicit

'-----------------------------------------------------------------------------
' log用のクラスです。
' 使用例
'   dim jj as log              ' オブジェクト変数の宣言
'   set jj = new log           ' インスタンス化
'   jj.pp a, b, c, d, e, f     ' 出力したいものをカンマ区切りで送る
'   jj.putf                    ' ファイルに出力する場合
' [注意] putfの引数a_dirを省略する場合、vbaではコメントアウト箇所を変更してください。
'-----------------------------------------------------------------------------

'***********************************************
Private Const ZERO As Long = 0
Private Const N1 As Long = 1
'***********************************************
Private m_key           As String
Private m_put_type      As Long
Private m_log_report    As Collection
Private m_log_el_count  As Long
Private had_put         As Boolean
Private m_start_time    As Single
Private m_log_file_name As String
'***********************************************

Public Property Let log_file_name(afile_name As String)
  m_log_file_name = afile_name
End Property
'-----------------------------------------------------------------------------
Public Property Get log_file_name() As String
  log_file_name = m_log_file_name
End Property
'-----------------------------------------------------------------------------

Public Property Let put_type(ByVal a_put_type As Long)
  m_put_type = a_put_type
End Property
'-----------------------------------------------------------------------------
Public Property Get put_type() As Long
  put_type = m_put_type
End Property
'-----------------------------------------------------------------------------
Public Property Let key(ByVal a_key As String)
  m_key = a_key
End Property
'-----------------------------------------------------------------------------
Public Property Get key() As String
  key = m_key
End Property
'-----------------------------------------------------------------------------

Public Sub pp(ParamArray ar() As Variant)
  Dim ii As Variant
  Dim buf As String
  For Each ii In ar
    buf = buf & "," & ii
  Next ii
  m_log_report.Add buf
  If m_log_el_count < UBound(ar) Then
    m_log_el_count = UBound(ar)
  End If
End Sub
'-----------------------------------------------------------------------------

' 2018.09.21(金) works用にファイル名を変更
Public Sub putf(Optional a_name As String, Optional add_info As Boolean = False)
  Const FILENAME_TEMP As String = "Log(works) "
  Dim cn        As Long
  Dim ii        As Variant
  Dim i         As Long
  Dim file_name As String
  Dim fn        As Long
  Dim buf       As String
  If Len(Trim(a_name)) Then
    buf = Trim(a_name)
  Else
    If Len(m_key) Then
      buf = m_key
    Else
      If Len(m_log_file_name) Then
        buf = FILENAME_TEMP & m_log_file_name
      Else
        buf = FILENAME_TEMP & "auto put"
      End If
    End If
  End If
  If buf Like "*.*" Then
    file_name = buf
  Else
    file_name = buf & ".csv"
  End If
  If file_name Like "*\*" Then
  Else
    file_name = fc.cur & file_name
  End If
  If m_log_report.Count Then
    fn = FreeFile()
    Open file_name For Output As #fn
    ' header------------------------------------
    If add_info Then
      Write #fn, "date"; Format(Date, "mm.dd(ddd)") '; clng(date)
      Write #fn, "time"; Format(Now, "hh:mmampm") '; clng(time)
      Write #fn, "run"; Format(Timer - m_start_time, "#,0.00")
      Write #fn,
    End If
    Write #fn, "seq";
    For i = N1 To m_log_el_count + N1
      Write #fn, "part_" & Format(i, "00");
    Next i
    Write #fn,
    ' putf log_report---------------------------
    For Each ii In m_log_report
      cn = cn + N1
      Print #fn, cn, ii
    Next ii
    Close #fn
    Debug.Print "putf", file_name
    'lpp "ログを出力しました(" & file_name & ")"
  End If
  had_put = True
End Sub
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
  m_start_time = Timer
  Set m_log_report = New Collection
End Sub
'-----------------------------------------------------------------------------

Private Sub Class_Terminate()
  Dim file_name As String
  ' 2018.11.07(水)  コメントアウト
  'If Not had_put Then
  '  If Len(m_log_file_name) Then
  '    file_name = m_log_file_name & "(terminate)"
  '  Else
  '    file_name = "log(terminate)"
  '  End If
  '  ' 2018.10.09(火) なんとなくinfoはいらないと思い修正。
  '  'Call putf(a_name:=file_name, add_info:=True)
  '  Call putf(a_name:=file_name, add_info:=False)
  'End If
  Set m_log_report = Nothing
End Sub
'-----------------------------------------------------------------------------

Public Function Count() As Long
  Count = m_log_report.Count
End Function
'-----------------------------------------------------------------------------

