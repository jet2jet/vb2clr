Attribute VB_Name = "ExitHandler"
' Copyright (C) 2018 jet
' For more information about license, see LICENSE.
'
' Helper module for handling on application-exit
Option Explicit

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' インスタンスのデータ
Private Type MyClassData
    vtblPtr As LongPtr
    RefCount As Long
#If Win64 Then ' 64ビット版かどうか
    Padding As Long
#End If
End Type

' 仮想関数テーブルのデータ
Private Type IUnknownVtbl
    QueryInterface As LongPtr
    AddRef As LongPtr
    Release As LongPtr
End Type

Private Const S_OK As Long = 0
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003

Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
    (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function GetProcessHeap Lib "kernel32.dll" () As LongPtr
Public Declare PtrSafe Function HeapAlloc Lib "kernel32.dll" _
    (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function HeapFree Lib "kernel32.dll" _
    (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Boolean

Public Declare PtrSafe Function CoTaskMemAlloc Lib "ole32.dll" _
    (ByVal cb As LongPtr) As LongPtr
Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal pv As LongPtr)

' VBA実行中は自前インスタンスが入り続ける変数
Dim m_unk As IUnknown
Dim m_collHandlers As Collection

' 変数に関数アドレスを代入するために用いる関数
Private Function GetAddressOf(ByVal func As LongPtr) As LongPtr
    GetAddressOf = func
End Function

' MyClassData と IUnknownVtbl のサイズを合わせたデータを指すポインターを返す
Private Function CreateInstanceMemory() As LongPtr
    Dim p As LongPtr, d As MyClassData, v As IUnknownVtbl
    ' MyClassData と IUnknownVtbl のサイズを合わせたデータを作成
    p = CoTaskMemAlloc(Len(d) + Len(v))
    If p <> 0 Then
        ' 最初の参照カウントは必ず 1 とする
        d.RefCount = 1
        ' MyClassData の直後に IUnknownVtbl を置くので p に MyClassData のサイズを加えたアドレスをセットする
        d.vtblPtr = p + Len(d)
        ' 割り当てたメモリブロックの先頭を MyClassData のデータで埋める
        Call CopyMemory(ByVal p, d, Len(d))
        ' 仮想関数テーブルの作成
        v.QueryInterface = GetAddressOf(AddressOf My_QueryInterface)
        v.AddRef = GetAddressOf(AddressOf My_AddRef)
        v.Release = GetAddressOf(AddressOf My_Release)
        ' 仮想関数テーブルを p + Len(d) の部分にコピー
        Call CopyMemory(ByVal d.vtblPtr, v, Len(v))
    End If
    CreateInstanceMemory = p
End Function

' HRESULT STDMETHODCALLTYPE QueryInterface(THIS_ REFIID refiid, LPVOID FAR* ppv)
' 別のインターフェイスへ変換するのをリクエストするときに呼び出される関数
' (ppv は念のため NULL チェックを入れるため ByVal で定義)
Private Function My_QueryInterface(ByVal This As LongPtr, ByRef refiid As IID, ByVal ppv As LongPtr) As Long
    Debug.Print "My_QueryInterface"
    If ppv = 0 Then
        Debug.Print "  E_POINTER"
        My_QueryInterface = E_POINTER
        Exit Function
    End If
    ' IID_IUnknown: {00000000-0000-0000-C000-000000000046} かどうか確認
    If refiid.Data1 = 0 And refiid.Data2 = 0 And refiid.Data3 = 0 And _
        refiid.Data4(0) = &HC0 And refiid.Data4(1) = 0 And _
        refiid.Data4(2) = 0 And refiid.Data4(3) = 0 And _
        refiid.Data4(4) = 0 And refiid.Data4(5) = 0 And _
        refiid.Data4(6) = 0 And refiid.Data4(7) = 0 Then
        ' IID_IUnknown の場合は ppv が指すポインターの先に This のアドレス(This の値)をコピー
        Debug.Print "  IID_IUnknown"
        Call CopyMemory(ByVal ppv, This, Len(This))
        ' さらに参照カウントを増やす
        Call My_AddRef(This)
        My_QueryInterface = S_OK
        Exit Function
    End If
    ' IID_IUnknown 以外はサポートしない
    Debug.Print "  E_NOINTERFACE"
    My_QueryInterface = E_NOINTERFACE
End Function

' ULONG STDMETHODCALLTYPE AddRef(THIS)
' 参照カウントを増やす際に呼び出される関数
Private Function My_AddRef(ByVal This As LongPtr) As Long
    Dim d As MyClassData
    ' インスタンスのデータを一旦 d にコピーし、
    ' 参照カウントを増やしたら書き戻す
    Call CopyMemory(d, ByVal This, Len(d))
    d.RefCount = d.RefCount + 1
    Debug.Print "My_AddRef: new RefCount ="; d.RefCount
    Call CopyMemory(ByVal This, d, Len(d))
    ' 戻り値は参照カウント
    My_AddRef = d.RefCount
End Function

' ULONG STDMETHODCALLTYPE Release(THIS)
' 参照カウントを減らす際に呼び出される関数(0 になったら破棄)
Private Function My_Release(ByVal This As LongPtr) As Long
    Dim d As MyClassData
    ' インスタンスのデータを一旦 d にコピーし、
    ' 参照カウントを減らしたら書き戻す
    Call CopyMemory(d, ByVal This, Len(d))
    d.RefCount = d.RefCount - 1
    Debug.Print "My_Release: new RefCount ="; d.RefCount
    Call CopyMemory(ByVal This, d, Len(d))
    ' 参照カウントが 0 になったら CoTaskMemFree で破棄する
    If d.RefCount = 0 Then
        Call CoTaskMemFree(This)
        ' 終了関数を呼び出す
        Call OnExit
    End If
    ' 戻り値は参照カウント
    My_Release = d.RefCount
End Function

' 終了時に Handler.OnExit() が呼び出されるように
' Handler オブジェクトを登録
Public Function AddExitHandler(ByVal Handler As Object, Optional ByVal Key As String) As Object
    Dim ptr As LongPtr
    If Not m_collHandlers Is Nothing Then
        On Error Resume Next
        Dim o As Object
        ptr = 0^
        ptr = m_collHandlers.Item(Key)
        On Error GoTo 0
        If ptr <> 0^ Then
            Call CopyMemory(o, ptr, Len(ptr))
            Set AddExitHandler = o
            ptr = 0^
            Call CopyMemory(o, ptr, Len(ptr))
            Exit Function
        End If
    End If
    If m_unk Is Nothing Then
        Dim p As LongPtr
        ' インスタンスを作成
        p = CreateInstanceMemory()
        If p = 0 Then Exit Function
        Dim unk As IUnknown
        ' unk を p が指すインスタンスに設定
        Call CopyMemory(unk, p, Len(p))
        ' m_unk にセット(内部で My_AddRef が呼び出される)
        Set m_unk = unk
        Set m_collHandlers = New Collection
    End If
    Call CopyMemory(ptr, Handler, Len(ptr))
    Call m_collHandlers.Add(ptr, Key)
    Set AddExitHandler = Handler
End Function

Public Sub RemoveExitHandler(ByVal Handler As Variant)
    If m_collHandlers Is Nothing Then Exit Sub
    If VarType(Handler) = vbString Then
        On Error Resume Next
        Call m_collHandlers.Remove(Handler)
        Exit Sub
    End If
    If VarType(Handler) <> vbObject And VarType(Handler) <> 13 Then
        Call Err.Raise(13)
    End If
    Dim ptr As LongPtr, i As Long
    On Error Resume Next
    For i = 1 To m_collHandlers.Count
        ptr = m_collHandlers.Item(i)
        If ptr = ObjPtr(Handler) Then
            Call m_collHandlers.Remove(i)
            Exit For
        End If
    Next i
End Sub

' VBA終了時の処理を記述
Private Sub OnExit()
    Dim o As Object
    On Error Resume Next
    For Each o In m_collHandlers
        Call o.OnExit
    Next o
End Sub
