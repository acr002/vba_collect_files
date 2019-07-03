VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C0path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------[date: 2019.04.04]
Option Explicit

'***********************************************
' 2017.12.04(月) revision01
' 2017.12.14(木) revision02
' 2018.07.24(火) revision03 プロパティの追加、ファイル名コレクションの追加
' 2018.08.02(木) col_xxのプロパティが抜けていたので追加
' 2018.09.14(金) col_cur, fd, dataを追加(できればすべてのファイルは取得してもいいかも)
'                1.key, 2.path, 3.filenames(col)のセットであるべき？
'                これならpathと個別のファイル名を取得しやすい？ ただこれだともう一階層深いものになる。。。
'***********************************************
Private m_cur          As String
Private m_ar           As Variant
Private m_drive        As String
Private m_jobcode      As String
Private m_root         As String
Private m_data         As String
Private m_fd           As String
Private m_in_          As String
Private m_out_         As String
Private m_files        As String
Private m_col_fn_in_sub As Collection
Private m_col_fn_in_   As Collection
Private m_col_fn_files As Collection
Private m_col_fn_cur   As Collection
Private m_col_fn_data  As Collection
Private m_col_fn_fd    As Collection
'***********************************************

Public Property Set col_fn_in_sub(ByVal vData As collection)
  set m_col_fn_in_sub = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get col_fn_in_sub() As collection
  set col_fn_in_sub = m_col_fn_in_sub
End Property
'-----------------------------------------------------------------------------

Public Property Set col_fn_in_(ByVal vData As collection)
  set m_col_fn_in_ = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get col_fn_in_() As collection
  set col_fn_in_ = m_col_fn_in_
End Property
'-----------------------------------------------------------------------------

Public Property Set col_fn_files(ByVal vData As collection)
  set m_col_fn_files = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get col_fn_files() As collection
  set col_fn_files = m_col_fn_files
End Property
'-----------------------------------------------------------------------------

Public Property Let in_(ByVal vData As String)
  m_in_ = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get in_() As String
  in_ = m_in_
End Property
'-----------------------------------------------------------------------------

Public Property Let out_(ByVal vData As String)
  m_out_ = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get out_() As String
  out_ = m_out_
End Property
'-----------------------------------------------------------------------------

Public Property Let files(ByVal vData As String)
  m_files = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get files() As String
  files = m_files
End Property
'-----------------------------------------------------------------------------

Public Property Let fd(ByVal vData As String)
  m_fd = vData
End Property
'-----------------------------------------------------------------------------
Public Property Get fd() As String
  fd = m_fd
End Property
'-----------------------------------------------------------------------------

Public Property Let data(ByVal vData As String)
  m_data = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get data() As String
  data = m_data
End Property
'-----------------------------------------------------------------------------

Public Property Let jobcode(ByVal vData As String)
  m_jobcode = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get jobcode() As String
  jobcode = m_jobcode
End Property
'-----------------------------------------------------------------------------

Public Property Let ar(ByVal vData As Variant)
  m_ar = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get ar() As Variant
  ar = m_ar
End Property
'-----------------------------------------------------------------------------

Public Property Let cur(ByVal vData As String)
  m_cur = vData
End Property
'-----------------------------------------------------------------------------

Public Property Get cur() As String
  cur = m_cur
End Property
'-----------------------------------------------------------------------------

Private Sub Class_Initialize()
  Const KEY    As String = "\"
  Dim filename As String
  'm_cur = CurDir() & KEY
  m_cur = thisworkbook.path & KEY
  m_ar = Split(m_cur, KEY)
  If m_ar(0) = "H:" Then
    m_drive = m_ar(0)
    m_jobcode = m_ar(1)
    m_root = m_drive & KEY & m_jobcode & KEY
    m_data = m_root & "data\"
    m_fd = m_root & "fd\"
  End If
  m_files = m_cur & "files\"
  m_in_ = m_cur & "in\"
  m_out_ = m_cur & "out\"
  ' in ///////////////////////////////////////////////////////////////////////
  Set m_col_fn_in_ = New Collection
  filename = Dir(m_in_ & "*.*")
  Do While Len(filename)
    m_col_fn_in_.Add filename
    filename = Dir()
  Loop
  ' files ////////////////////////////////////////////////////////////////////
  Set m_col_fn_files = New Collection
  filename = Dir(m_files & "*.*")
  Do While Len(filename)
    m_col_fn_files.Add filename
    filename = Dir()
  Loop
  ' cur //////////////////////////////////////////////////////////////////////
  Set m_col_fn_cur = New Collection
  filename = Dir(m_cur & "*.*")
  Do While Len(filename)
    m_col_fn_cur.Add filename
    filename = Dir()
  Loop
  ' fd ///////////////////////////////////////////////////////////////////////
  Set m_col_fn_fd = New Collection
  filename = Dir(m_fd & "*.*")
  Do While Len(filename)
    m_col_fn_fd.Add filename
    filename = Dir()
  Loop
  ' data /////////////////////////////////////////////////////////////////////
  Set m_col_fn_data = New Collection
  filename = Dir(m_data & "*.*")
  Do While Len(filename)
    m_col_fn_data.Add filename
    filename = Dir()
  Loop
  'Call chk_dir_mkdir(m_data)
  'Call chk_dir_mkdir(m_fd)
  ' in to subfolders /////////////////////////////////////////////////////////
  set m_col_fn_in_sub = get_filenames_sub(m_in_)
End Sub
'-----------------------------------------------------------------------------

Private Sub chk_dir_mkdir(a_path As String)
  If Len(Dir(a_path, vbDirectory)) = 0 Then
    MkDir a_path
  End If
End Sub
'-----------------------------------------------------------------------------

' 2019.04.04(木).new
' サブフォルダも含めてファイル名を取得します。
' フルパスを保持します。
Private Function get_filenames_sub(ByVal a_path As String) As Collection
  Dim r_cc     As Collection
  Dim cc       As Collection
  Dim ii       As Variant
  Dim fso      As object
  Dim b_file   As object
  Dim b_folder As object
  Set cc = New Collection
  Set fso = Createobject("Scripting.FileSystemObject")
  For Each b_file In fso.getfolder(a_path).Files
    cc.Add b_file.Path
  Next b_file
  For Each b_folder In fso.getfolder(a_path).subfolders
    Set r_cc = get_filenames_sub(b_folder.Path)
    For Each ii In r_cc
      cc.Add ii
    Next ii
  Next b_folder
  Set get_filenames_sub = cc
End Function
'-----------------------------------------------------------------------------

