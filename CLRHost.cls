VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLRHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright (C) 2018 jet
' For more information about license, see LICENSE.
'
' Defines CLRHost class
Option Explicit

Private Declare PtrSafe Function VariantCopy Lib "oleaut32.dll" _
    (ByRef pvargDest As Variant, ByRef pvargSrc As Variant) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" _
    (ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Long, _
    ByVal cActuals As Long, _
    ByRef prgvt As Long, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

Private Declare PtrSafe Function CLRCreateInstance Lib "mscoree.dll" _
    (ByRef rclsid As Any, ByRef riid As Any, ByRef ppvInterface As IUnknown) As Long

Dim m_host As mscoree.CorRuntimeHost
Dim m_domain As mscorlib.AppDomain
Dim m_asmCore As mscorlib.Assembly
Dim m_typeObject As mscorlib.Type
Dim m_objManagedSupport As Object
Dim m_collTypeCache As Collection

Private Sub Class_Terminate()
    Call CleanupImpl
End Sub

Private Sub CleanupImpl()
    If Not m_host Is Nothing Then
        Set m_collTypeCache = Nothing
        Set m_objManagedSupport = Nothing
        Set m_typeObject = Nothing
        Set m_asmCore = Nothing
        Call m_host.UnloadDomain(m_domain)
        Set m_domain = Nothing
        Call m_host.Stop
        Set m_host = Nothing
        Call RemoveExitHandler(Me)
    End If
End Sub

' Parses GUID in string format into an array of Long.
' The array can be used as a pointer to GUID structure data.
Public Sub ParseGUID(ByRef lnGUID() As Long, ByVal str As String)
    If Left$(str, 1) = "{" Then
        If Right$(str, 1) <> "}" Then Call Err.Raise(5)
        str = Mid$(str, 2, Len(str) - 2)
    End If
    If Len(str) <> 8 + 4 + 4 + 4 + 12 + 4 Then Call Err.Raise(5)
    Dim parts() As String
    parts = Split(str, "-")
    If LBound(parts) <> 0 Or UBound(parts) <> 4 Then Call Err.Raise(5)
    If Len(parts(0)) <> 8 Then Call Err.Raise(5)
    If Len(parts(1)) <> 4 Then Call Err.Raise(5)
    If Len(parts(2)) <> 4 Then Call Err.Raise(5)
    If Len(parts(3)) <> 4 Then Call Err.Raise(5)
    If Len(parts(4)) <> 12 Then Call Err.Raise(5)
    lnGUID(0) = CLng("&H" + parts(0))
    lnGUID(1) = CLng("&H" + parts(2) + parts(1))
    lnGUID(2) = CLng("&H" + Mid$(parts(4), 3, 2) + Mid$(parts(4), 1, 2) + Right$(parts(3), 2) + Left$(parts(3), 2))
    lnGUID(3) = CLng("&H" + Mid$(parts(4), 11, 2) + Mid$(parts(4), 9, 2) + Mid$(parts(4), 7, 2) + Mid$(parts(4), 5, 2))
End Sub

' Call object's method with index of vftable
Private Function VBCallAbsoluteObject(ByVal Object As IUnknown, _
    ByVal IndexForVftable As Integer, _
    ByVal RetType As VbVarType, _
    ParamArray Arguments() As Variant) As Variant
    If Object Is Nothing Then
        Call Err.Raise(5)
    End If
    Dim hr As Long
    Dim argVt() As Long
    Dim argsPtr() As LongPtr
    Dim i As Long, c As Long
    Dim lb As Long, ub As Long
    lb = LBound(Arguments)
    ub = UBound(Arguments)
    c = ub - lb + 1
    If c > 0 Then
        ReDim argVt(lb To ub)
        ReDim argsPtr(lb To ub)
        For i = lb To ub
            argVt(i) = VarType(Arguments(i))
            argsPtr(i) = VarPtr(Arguments(i))
        Next i
        hr = DispCallFunc(ObjPtr(Object), _
            CLngPtr(IndexForVftable) * Len(argsPtr(0)), _
            4, _
            CLng(RetType), _
            c, _
            argVt(lb), _
            argsPtr(lb), _
            VBCallAbsoluteObject)
    Else
        ReDim argVt(0)
        ReDim argsPtr(0)
        hr = DispCallFunc(ObjPtr(Object), _
            CLngPtr(IndexForVftable) * Len(argsPtr(0)), _
            4, _
            CLng(RetType), _
            0, _
            argVt(0), _
            argsPtr(0), _
            VBCallAbsoluteObject)
    End If
    If hr < 0 Then Call Err.Raise(hr)
End Function

Private Function CreateCorRuntimeHost(ByVal Version As String) As mscoree.CorRuntimeHost
    Dim g(0 To 3) As Long
    Dim g2(0 To 3) As Long
    Dim pMetaHost As IUnknown, hr As Long
    Call ParseGUID(g, "{9280188D-0E8E-4867-B30C-7FA83884E8DE}") ' CLSID_CLRMetaHost
    Call ParseGUID(g2, "{D332DB9E-B9B3-4125-8207-A14884F53216}") ' IID_ICLRMetaHost
    hr = CLRCreateInstance(g(0), g2(0), pMetaHost)
    If hr < 0 Then Call Err.Raise(hr)

    Dim pRuntimeInfo As IUnknown
    Call ParseGUID(g, "{BD39D1D2-BA2F-486A-89B0-B4B0CB466891}") ' IID_ICLRRuntimeInfo
    ' ICLRMetaHost::GetRuntime(LPCWSTR, REFIID, void**) [vftable index = 3]
    hr = VBCallAbsoluteObject(pMetaHost, 3, vbLong, _
        StrPtr(Version), VarPtr(g(0)), VarPtr(pRuntimeInfo))
    Set pMetaHost = Nothing
    If hr < 0 Then Call Err.Raise(hr)

    Dim pCorRuntimeHost As IUnknown
    ' ICLRRuntimeInfo::GetInterface(REFCLSID, REFIID, void**) [vftable index = 9]
    Call ParseGUID(g, "{CB2F6723-AB3A-11D2-9C40-00C04FA30A3E}") ' CLSID_CorRuntimeHost
    Call ParseGUID(g2, "{CB2F6722-AB3A-11D2-9C40-00C04FA30A3E}") ' IID_ICorRuntimeHost
    hr = VBCallAbsoluteObject(pRuntimeInfo, 9, vbLong, _
        VarPtr(g(0)), VarPtr(g2(0)), VarPtr(pCorRuntimeHost))
    Set pRuntimeInfo = Nothing
    If hr < 0 Then Call Err.Raise(hr)
    Set CreateCorRuntimeHost = pCorRuntimeHost
End Function

Private Function GetCLRTypeFromInheritancesByMemberName(ByVal t As mscorlib.Type, ByVal MemberName As String, ByVal MemberType As mscorlib.MemberTypes) As mscorlib.Type
    Dim arrM() As mscorlib.MemberInfo
    Dim mi As mscorlib.MemberInfo, lb As Long, ub As Long, i As Long
    arrM = t.GetMembers_2()
    lb = LBound(arrM)
    ub = UBound(arrM)
    For i = lb To ub
        Set mi = arrM(i)
        If mi.Name = MemberName And (mi.MemberType And MemberTypes_All) <> 0 Then
            Set GetCLRTypeFromInheritancesByMemberName = t
            Exit Function
        End If
    Next i
    If t.BaseType Is Nothing Then
        Set GetCLRTypeFromInheritancesByMemberName = Nothing
        Exit Function
    End If
    Set GetCLRTypeFromInheritancesByMemberName = GetCLRTypeFromInheritancesByMemberName(t.BaseType, MemberName, MemberType)
End Function

Private Function ExecuteCSharpCode(ByVal domain As mscorlib.AppDomain, ByVal code As String, _
    ParamArray RefAssemblyName() As Variant) As mscorlib.Assembly
    Dim asmSys As mscorlib.Assembly, o As Object
    'Set asmSys = domain.Load_2("System.dll")
    Set asmSys = domain.Load_2("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    Dim cobjParams As mscorlib.Object
    Set cobjParams = asmSys.CreateInstance("System.CodeDom.Compiler.CompilerParameters")
    Set o = cobjParams
    o.GenerateInMemory = True
    Dim v As Variant
    Dim oRefAsms As Object
    Set oRefAsms = ToObject(o.ReferencedAssemblies)
    For Each v In RefAssemblyName
        Call oRefAsms.Add(CStr(v))
    Next v
    Dim cobjProvider As mscorlib.Object
    Dim vCodes(0) As String
    Set cobjProvider = asmSys.CreateInstance("Microsoft.CSharp.CSharpCodeProvider")
    Set o = cobjProvider
    vCodes(0) = code
    Dim cobjResults As mscorlib.Object
    Set cobjResults = o.CompileAssemblyFromSource(cobjParams, vCodes)
    Set o = cobjResults
    Dim oErrors As Object
    Set oErrors = ToObject(o.Errors)
    If oErrors.HasErrors Then
        Dim oError As Object
        Dim c As Long, i As Long
        c = oErrors.Count - 1
        For i = 0 To c
            Set oError = ToObject(oErrors.Item(i))
            Debug.Print oError.ErrorText; " [Line="; oError.Line; "]"
        Next i
        Exit Function
    End If
    Set ExecuteCSharpCode = o.CompiledAssembly
End Function

Private Function IsType(ByRef v As Variant) As Boolean
    IsType = (VarType(v) = 13 And TypeOf v Is mscorlib.Type)
End Function

' PickupMethodByParamTypeImpl: 型「t」に含まれる名前「Name」のメソッドを検索します。
'   BindingFlags: 絞り込み条件
'   Types: メソッドの引数に対応する型を表す mscorlib.Type のインスタンス。
'          対応する引数の数だけ指定します。
'   戻り値: mscorlib.MethodInfo のインスタンス、または見つからない場合 Nothing
Private Function PickupMethodByParamTypeImpl(ByVal t As mscorlib.Type, ByVal Name As String, ByVal BindingFlags As mscorlib.BindingFlags, ByRef Types() As Variant) As mscorlib.MethodInfo
    Dim tlb As Long, tub As Long
    Dim actualTypeCount As Long
    Dim i As Long
    tlb = LBound(Types)
    tub = UBound(Types)
    actualTypeCount = 0
    ' 可変長引数で指定された Type インスタンスの数を計算
    For i = tlb To tub
        ' Types(i) が Type インスタンスかどうかを判定
        ' (13 は VT_UNKNOWN、mscorlib.Type は IUnknown ベースの型のため念のためチェック)
        If IsType(Types(i)) Then
            actualTypeCount = actualTypeCount + 1
        End If
    Next i

    Dim mis() As mscorlib.MethodInfo, mi As mscorlib.MethodInfo
    Dim j As Long, k As Long, Matched As Boolean
    ' BindingFlags で対応するメソッドの一覧を取得
    mis = t.GetMethods(BindingFlags)
    ' それぞれのメソッドに対し、メソッド名と引数の型をチェック
    For i = LBound(mis) To UBound(mis)
        Set mi = mis(i)
        ' メソッド名を Option Compare ステートメント設定に応じた比較方法で判定
        ' (大文字・小文字の区別を制御したい場合は StrComp 関数を用います。)
        If mi.Name = Name Then
            Dim p() As mscorlib.ParameterInfo
            ' 引数情報を取得(配列が返ります)
            p = mi.GetParameters()
            ' 引数の数の一致性を確認
            If UBound(p) - LBound(p) + 1 = actualTypeCount Then
                Matched = True
                k = LBound(p)
                For j = tlb To tub
                    If IsType(Types(j)) Then
                        Dim objPI As mscorlib.Object
                        Dim o As Object, tP As mscorlib.Type
                        ' ParameterInfo クラスの ParameterType プロパティーを見るため、
                        ' mscorlib.Object 経由で VB の Object 型に変換
                        Set objPI = p(k)
                        Set o = objPI
                        Set tP = o.ParameterType
                        ' 型が一致しない引数が一つでもあった場合は不一致とする
                        If Not tP.Equals(Types(j)) Then
                            Matched = False
                            Exit For
                        End If
                        k = k + 1
                    End If
                Next j
                ' 一致した場合は見つかったものとしてループを抜ける
                If Matched Then Exit For
            End If
        End If
        Set mi = Nothing
    Next i
    Set PickupMethodByParamTypeImpl = mi
End Function

Private Function PickupMethodByParamType(ByVal t As mscorlib.Type, ByVal Name As String, ByVal BindingFlags As mscorlib.BindingFlags, ParamArray Types() As Variant) As mscorlib.MethodInfo
    Dim v() As Variant
    v = Types
    Set PickupMethodByParamType = PickupMethodByParamTypeImpl(t, Name, BindingFlags, v)
End Function

Private Function CreateManagedSupportObject(ByVal domain As mscorlib.AppDomain) As Object
    Dim asm As mscorlib.Assembly
    Dim rName As String, rIName As String
    Dim code As String
    rIName = "VBVariableDelegateWrapper_" + CStr(Int(Rnd() * 16384))
    rName = "VBManagedSupport_" + CStr(Int(Rnd() * 16384))
    code = "" + _
        "using System;" + vbCrLf + _
        "using System.Reflection;" + vbCrLf + _
        "using System.Runtime.InteropServices;" + vbCrLf
    ' class VBVariableDelegateWrapper
    '   handles any delegate types using methods with different parameter count
    code = code + "class " + rIName + vbCrLf + _
        "{" + vbCrLf + _
        "    private object _target;" + vbCrLf + _
        "    private Type _typeTarget;" + vbCrLf + _
        "    private string _methodName;" + vbCrLf + _
        "" + vbCrLf + _
        "    public static Delegate CreateDelegate(Type delegateType, object target, string methodName)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        return CreateDelegate(delegateType, target, target.GetType(), methodName);" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf + _
        "    public static Delegate CreateDelegate(Type delegateType, object target, Type targetType, string methodName)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        if (!delegateType.IsSubclassOf(typeof(Delegate)))" + vbCrLf + _
        "        {" + vbCrLf + _
        "            throw new ArgumentException(""Invalid 'delegateType'"");" + vbCrLf
    code = code + "        }" + vbCrLf + _
        "        var miInvoke = delegateType.GetMethod(""Invoke"");" + vbCrLf + _
        "        if (miInvoke == null)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            throw new ArgumentException(""Invalid 'delegateType'"");" + vbCrLf + _
        "        }" + vbCrLf + _
        "        var c = miInvoke.GetParameters().Length;" + vbCrLf + _
        "        if (c > 19)" + vbCrLf + _
        "        {" + vbCrLf + _
        "            throw new NotSupportedException(""Parameter count of delegate is too large (maximum support = 19)"");" + vbCrLf + _
        "        }" + vbCrLf + _
        "        var wrapper = new " + rIName + "(target, targetType, methodName);" + vbCrLf + _
        "        var miWrapper = typeof(" + rIName + ").GetMethod(" + vbCrLf + _
        "            (miInvoke.ReturnType == typeof(void) ? ""VMethod"" : ""Method"") + c.ToString()," + vbCrLf + _
        "            BindingFlags.NonPublic | BindingFlags.Instance" + vbCrLf + _
        "            );" + vbCrLf + _
        "        return Delegate.CreateDelegate(delegateType, wrapper, miWrapper);" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf + _
        "    private " + rIName + "(object target, Type targetType, string methodName)" + vbCrLf
    code = code + "    {" + vbCrLf + _
        "        _target = target;" + vbCrLf + _
        "        _typeTarget = targetType;" + vbCrLf + _
        "        _methodName = methodName;" + vbCrLf + _
        "    }" + vbCrLf + _
        "" + vbCrLf + _
        "    private object Method(object[] args)" + vbCrLf + _
        "    {" + vbCrLf + _
        "        return _typeTarget.InvokeMember(_methodName," + vbCrLf + _
        "            BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding," + vbCrLf + _
        "            null, _target, args);" + vbCrLf + _
        "    }" + vbCrLf + _
        "    private object Method0() { return Method(new object[] { }); }" + vbCrLf + _
        "    private object Method1(object p1) { return Method(new object[] { p1 }); }" + vbCrLf + _
        "    private object Method2(object p1, object p2) { return Method(new object[] { p1, p2 }); }" + vbCrLf + _
        "    private object Method3(object p1, object p2, object p3) { return Method(new object[] { p1, p2, p3 }); }" + vbCrLf + _
        "    private object Method4(object p1, object p2, object p3, object p4) { return Method(new object[] { p1, p2, p3, p4 }); }" + vbCrLf + _
        "    private object Method5(object p1, object p2, object p3, object p4, object p5) { return Method(new object[] { p1, p2, p3, p4, p5 }); }" + vbCrLf + _
        "    private object Method6(object p1, object p2, object p3, object p4, object p5, object p6) { return Method(new object[] { p1, p2, p3, p4, p5, p6 }); }" + vbCrLf + _
        "    private object Method7(object p1, object p2, object p3, object p4, object p5, object p6, object p7) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7 }); }" + vbCrLf
    code = code + "    private object Method8(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8 }); }" + vbCrLf + _
        "    private object Method9(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9 }); }" + vbCrLf + _
        "    private object Method10(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10 }); }" + vbCrLf + _
        "    private object Method11(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11 }); }" + vbCrLf + _
        "    private object Method12(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12 }); }" + vbCrLf + _
        "    private object Method13(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13 }); }" + vbCrLf + _
        "    private object Method14(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14 }); }" + vbCrLf + _
        "    private object Method15(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15 }); }" + vbCrLf + _
        "    private object Method16(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16 }); }" + vbCrLf + _
        "    private object Method17(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16, object p17) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17 }); }" + vbCrLf + _
        "    private object Method18(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16, object p17, object p18) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17, p18 }); }" + vbCrLf + _
        "    private object Method19(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16, object p17, object p18, object p19) { return Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17, p18, p19 }); }" + vbCrLf + _
        "    private void VMethod0() { Method(new object[] { }); }" + vbCrLf + _
        "    private void VMethod1(object p1) { Method(new object[] { p1 }); }" + vbCrLf + _
        "    private void VMethod2(object p1, object p2) { Method(new object[] { p1, p2 }); }" + vbCrLf + _
        "    private void VMethod3(object p1, object p2, object p3) { Method(new object[] { p1, p2, p3 }); }" + vbCrLf + _
        "    private void VMethod4(object p1, object p2, object p3, object p4) { Method(new object[] { p1, p2, p3, p4 }); }" + vbCrLf + _
        "    private void VMethod5(object p1, object p2, object p3, object p4, object p5) { Method(new object[] { p1, p2, p3, p4, p5 }); }" + vbCrLf + _
        "    private void VMethod6(object p1, object p2, object p3, object p4, object p5, object p6) { Method(new object[] { p1, p2, p3, p4, p5, p6 }); }" + vbCrLf + _
        "    private void VMethod7(object p1, object p2, object p3, object p4, object p5, object p6, object p7) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7 }); }" + vbCrLf
    code = code + "    private void VMethod8(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8 }); }" + vbCrLf + _
        "    private void VMethod9(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9 }); }" + vbCrLf + _
        "    private void VMethod10(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10 }); }" + vbCrLf + _
        "    private void VMethod11(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11 }); }" + vbCrLf + _
        "    private void VMethod12(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12 }); }" + vbCrLf + _
        "    private void VMethod13(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13 }); }" + vbCrLf + _
        "    private void VMethod14(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14 }); }" + vbCrLf + _
        "    private void VMethod15(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15 }); }" + vbCrLf + _
        "    private void VMethod16(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16 }); }" + vbCrLf + _
        "    private void VMethod17(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16, object p17) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17 }); }" + vbCrLf + _
        "    private void VMethod18(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16, object p17, object p18) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17, p18 }); }" + vbCrLf + _
        "    private void VMethod19(object p1, object p2, object p3, object p4, object p5, object p6, object p7, object p8, object p9, object p10, object p11, object p12, object p13, object p14, object p15, object p16, object p17, object p18, object p19) { Method(new object[] { p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, p16, p17, p18, p19 }); }" + vbCrLf + _
        "}" + vbCrLf + _
        "" + vbCrLf
    ' end of class VBVariableDelegateWrapper

    ' class VBManagedSupport
    '   provides following methods (with VB-style declaration):
    '     Function MyInvoke(mb As MethodBase, ins As Variant, pms() As Variant, isIntPtr() As Variant) As Variant
    '     Function MyGetType(obj As Variant) As Type
    '     Function MyInvokeMember(targetType As Type, methodName As String, bindingFlags As BindingFlags, obj As Variant, methodArgs() As Variant) As Variant
    '     Function MyInvokeInstanceMember(obj As Variant, methodName As String, bindingFlags As BindingFlags, methodArgs() As Variant) As Variant
    '     Function MyCreateDelegate(delegateType As Type, obj As Variant, methodName As String) As Delegate
    code = code + _
        "[ComVisible(true)]" + vbCrLf + _
        "public class " + rName + " {" + vbCrLf + _
        "  public object MyInvoke(MethodBase mb, object ins, object[] pms, object[] isIntPtr) {" + vbCrLf + _
        "    object[] newParams = new object[pms.Length];" + vbCrLf + _
        "    for (int i = 0; i < pms.Length; ++i) {" + vbCrLf + _
        "      if (isIntPtr[i] is bool && (bool)isIntPtr[i]) {" + vbCrLf + _
        "        object o = pms[i];" + vbCrLf + _
        "        if (o is int) {" + vbCrLf + _
        "          newParams[i] = new IntPtr((int)o);" + vbCrLf + _
        "        } else if (o is Int64) {" + vbCrLf + _
        "          newParams[i] = new IntPtr((Int64)o);" + vbCrLf + _
        "        } else {" + vbCrLf + _
        "          newParams[i] = (IntPtr)o;" + vbCrLf + _
        "        }" + vbCrLf + _
        "      } else {" + vbCrLf + _
        "        newParams[i] = pms[i];" + vbCrLf + _
        "      }" + vbCrLf + _
        "    }" + vbCrLf + _
        "    return mb.Invoke(ins, newParams);" + vbCrLf + _
        "  }" + vbCrLf
    code = code + _
        "  public Type MyGetType(object obj) {" + vbCrLf + _
        "    return obj.GetType();" + vbCrLf + _
        "  }" + vbCrLf
    code = code + _
        "  public object MyInvokeMember(Type targetType, string methodName, BindingFlags bindingFlags, object obj, object[] methodArgs) {" + vbCrLf + _
        "    return targetType.InvokeMember(methodName, bindingFlags, null, obj, methodArgs);" + vbCrLf + _
        "  }" + vbCrLf
    code = code + _
        "  public object MyInvokeInstanceMember(object obj, string methodName, object[] methodArgs) {" + vbCrLf + _
        "    return obj.GetType().InvokeMember(methodName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.InvokeMethod, null, obj, methodArgs);" + vbCrLf + _
        "  }" + vbCrLf
    code = code + _
        "  public Delegate MyCreateDelegate(Type delegateType, object obj, string methodName) {" + vbCrLf + _
        "    return " + rIName + ".CreateDelegate(delegateType, obj, methodName);" + vbCrLf + _
        "  }" + vbCrLf
    code = code + _
        "  public Delegate MyCreateDelegateWithFunction(Type delegateType, object objPtr) {" + vbCrLf + _
        "    IntPtr x;" + vbCrLf + _
        "    if (objPtr is int) {" + vbCrLf + _
        "      x = new IntPtr((int)objPtr);" + vbCrLf + _
        "    } else if (objPtr is Int64) {" + vbCrLf + _
        "      x = new IntPtr((Int64)objPtr);" + vbCrLf + _
        "    } else {" + vbCrLf + _
        "      x = (IntPtr)objPtr;" + vbCrLf + _
        "    }" + vbCrLf + _
        "    return Marshal.GetDelegateForFunctionPointer(x, delegateType);" + vbCrLf + _
        "  }" + vbCrLf + _
        "}" + vbCrLf
    ' end of class VBManagedSupport

    'Set asm = domain.Load_2("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
    Set asm = ExecuteCSharpCode(domain, code)
    Set CreateManagedSupportObject = ToObject(asm.CreateInstance(rName))
End Function

'-------------------------------------------------------------------------------
' Public properties/methods
'-------------------------------------------------------------------------------

' Initializes this object.
'   TerminateOnExit: Terminates CLR when Visual Basic application finishes
'     WARNING: When True, you must not Stop debug when breaking.
'              When False, you should not Stop debug without releasing this instance.
'   Version: CLR version to use
Public Sub Initialize(ByVal TerminateOnExit As Boolean, Optional ByVal Version As String = "v4.0.30319")
    If Not m_host Is Nothing Then Exit Sub
    If TerminateOnExit Then
        Dim o As Object
        Set o = AddExitHandler(Me, "VBCLRHost")
        If Not o Is Me Then
            Call Err.Raise(31027, , "Another CLRHost instance is running")
        End If
    End If
    On Error GoTo Handler
    Set m_host = CreateCorRuntimeHost(Version)
    Call m_host.Start
    On Error GoTo Handler2
    Call m_host.CreateDomain("VBCLRHostDomain", Nothing, m_domain)
    Set m_asmCore = m_domain.Load_2("mscorlib")
    Set m_typeObject = m_asmCore.GetType_2("System.Object")
    Set m_objManagedSupport = CreateManagedSupportObject(m_domain)
    Set m_collTypeCache = New Collection
    Exit Sub
Handler2:
    Set m_typeObject = Nothing
    Set m_asmCore = Nothing
    If Not m_domain Is Nothing Then Call m_host.UnloadDomain(m_domain)
    Set m_domain = Nothing
    Call m_host.Stop
Handler:
    Set m_host = Nothing
    Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

' Terminates the object.
Public Sub Terminate()
    Call CleanupImpl
End Sub

' (Used by ExitHandler module)
Public Sub OnExit()
Attribute OnExit.VB_MemberFlags = "40"
    Call CleanupImpl
End Sub

' Converts CLR's object to VB Object
Public Function ToObject(ByVal obj As mscorlib.Object) As Object
    Set ToObject = obj
End Function

' Converts CLR's object to IEnumerable
Public Function ToEnumerable(ByVal obj As mscorlib.Object) As mscorlib.IEnumerable
    Set ToEnumerable = obj
End Function

' Gets the actual CorRuntimeHost object
Public Property Get RuntimeHost() As mscoree.CorRuntimeHost
    Set RuntimeHost = m_host
End Property

' Gets the AppDomain object for CLRHost
Public Property Get AppDomain() As mscorlib.AppDomain
    Set AppDomain = m_domain
End Property

' Loads the assembly (same as AppDomain.Load_2)
Public Function CLRLoadAssembly(ByVal AssemblyName As String) As mscorlib.Assembly
    Set CLRLoadAssembly = m_domain.Load_2(AssemblyName)
End Function

' Returns the result of obj.GetType().
' Some of objects cannot call obj.GetType() directory;
' using this method you can retrieve the Type instance of the object.
Public Function CLRGetType(ByVal obj As mscorlib.Object) As mscorlib.Type
    Set CLRGetType = m_objManagedSupport.MyGetType(obj)
End Function

' Resolves the type name to Type instance
' This method searches the type from all loaded assemblies.
' This method also supports the cache system;
' once the type is resolved, next time this method will return
' the cached type for the same type (will be faster).
Public Function CLRResolveType(ByVal TypeName As String) As mscorlib.Type
    Dim strActualTypeName As String
    strActualTypeName = TypeName
    Set CLRResolveType = Nothing
    On Error Resume Next
    Set CLRResolveType = m_collTypeCache.Item(strActualTypeName)
    On Error GoTo 0
    If Not CLRResolveType Is Nothing Then
        Exit Function
    End If

    Dim asms() As mscorlib.Assembly
    Dim v As Variant, asm As Assembly
    Set CLRResolveType = Nothing
    asms = m_domain.GetAssemblies()
    For Each v In asms
        Set asm = v
        Set CLRResolveType = asm.GetType_2(strActualTypeName)
        If Not CLRResolveType Is Nothing Then
            Call m_collTypeCache.Add(CLRResolveType, strActualTypeName)
            Exit Function
        End If
    Next v
End Function

' Creates the instance of specified type
Public Function CLRCreateObject(ByVal TypeName As String) As mscorlib.Object
    Dim t As mscorlib.Type
    Set t = CLRResolveType(TypeName)
    If t Is Nothing Then Call Err.Raise(419)
    Set CLRCreateObject = t.Assembly.CreateInstance_2(TypeName, False)
End Function

' Creates the instance of specified type with constructor parameters
Public Function CLRCreateObjectWithParams(ByVal TypeName As String, ParamArray Arguments() As Variant) As mscorlib.Object
    Dim t As mscorlib.Type
    Set t = CLRResolveType(TypeName)
    If t Is Nothing Then Call Err.Raise(419)
    Dim v() As Variant
    v = Arguments
    Set CLRCreateObjectWithParams = ToObject(t.Assembly).CreateInstance_3(TypeName, False, _
        BindingFlags_Public Or BindingFlags_Instance Or BindingFlags_CreateInstance, _
        Nothing, v, Nothing, Array())
End Function

' Returns the property value of specified name in the object with additional parameters
Public Function CLRPropertyGet(ByVal obj As mscorlib.Object, ByVal PropName As String, ParamArray Arguments() As Variant) As Variant
    Dim hr As Long, v() As Variant
    v = Arguments
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRPropertyGet, m_objManagedSupport.MyInvokeMember( _
        CLRGetType(obj), _
        PropName, _
        BindingFlags_GetProperty Or BindingFlags_Instance Or BindingFlags_Public, _
        obj, v))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Changes the property value of specified name in the object with additional parameters
Public Sub CLRPropertyPut(ByVal obj As mscorlib.Object, ByVal PropName As String, ByVal Value As Variant, ParamArray Arguments() As Variant)
    Dim v() As Variant
    Dim lb As Long, ub As Long, c As Long, i As Long, j As Long
    lb = LBound(Arguments)
    ub = UBound(Arguments)
    c = ub - lb + 1
    ReDim v(0 To c)
    j = lb
    Call VariantCopy(v(0), Value)
    For i = 1 To c
        Call VariantCopy(v(i), Arguments(j))
        j = j + 1
    Next i
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    Call m_objManagedSupport.MyInvokeMember( _
        CLRGetType(obj), _
        PropName, _
        BindingFlags_SetProperty Or BindingFlags_Instance Or BindingFlags_Public, _
        obj, v)
End Sub

' Accessor of CLRPropertyGet for using property-style
Public Property Get CLRProperty(ByVal obj As mscorlib.Object, ByVal PropName As String) As Variant
    Call VariantCopy(CLRProperty, CLRPropertyGet(obj, PropName))
End Property

' Accessor of CLRPropertyPut for using property-style
Public Property Let CLRProperty(ByVal obj As mscorlib.Object, ByVal PropName As String, ByVal Value As Variant)
    Call CLRPropertyPut(obj, PropName, Value)
End Property

' Accessor of CLRPropertyPut for using property-style (with Set)
Public Property Set CLRProperty(ByVal obj As mscorlib.Object, ByVal PropName As String, ByVal Value As Variant)
    Call CLRPropertyPut(obj, PropName, Value)
End Property

' Calls the instance method
Public Function CLRInvokeMethod(ByVal obj As mscorlib.Object, ByVal MethodName As String, ParamArray Arguments() As Variant) As Variant
    Dim v() As Variant
    v = Arguments
    Dim hr As Long
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRInvokeMethod, m_objManagedSupport.MyInvokeInstanceMember( _
        obj, _
        MethodName, _
        v _
    ))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Calls the instance method with specified type.
'   ArgTypesAndArgs must be the following values:
'     index[0, 2, 4, ...]: Type instance(s)
'     index[1, 3, 5, ...]: parameter value(s)
Public Function CLRInvokeMethodWithTypes(ByVal obj As mscorlib.Object, ByVal MethodName As String, ParamArray ArgTypesAndArgs() As Variant) As Variant
    Dim ArgTypes() As Variant, Args() As Variant
    Dim c As Long, lb As Long, ub As Long, i As Long
    lb = LBound(ArgTypesAndArgs)
    ub = UBound(ArgTypesAndArgs)
    c = ub - lb + 1
    If c Mod 2 <> 0 Then
        Call Err.Raise(5)
    End If
    ReDim ArgTypes(0 To (c / 2) - 1), Args(0 To (c / 2) - 1)
    Dim j As Long
    j = lb
    For i = 0 To c - 1 Step 2
        If Not IsType(ArgTypesAndArgs(j)) Then Call Err.Raise(5)
        Set ArgTypes(i) = ArgTypesAndArgs(j)
        Call VariantCopy(Args(i), ArgTypesAndArgs(j + 1))
        j = j + 2
    Next i

    Dim cmi As mscorlib.MethodInfo
    Set cmi = PickupMethodByParamTypeImpl(CLRGetType(obj), MethodName, BindingFlags_Public Or BindingFlags_Instance, ArgTypes)
    If cmi Is Nothing Then Call Err.Raise(438)

    Dim hr As Long
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRInvokeMethodWithTypes, ToObject(cmi).Invoke_3(obj, Args))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Calls the instance method with specified type.
Public Function CLRInvokeMethodWithTypes2(ByVal obj As mscorlib.Object, ByVal MethodName As String, ByRef ArgTypes() As Variant, ParamArray Arguments() As Variant) As Variant
    Dim cmi As mscorlib.MethodInfo
    Set cmi = PickupMethodByParamTypeImpl(CLRGetType(obj), MethodName, BindingFlags_Public Or BindingFlags_Instance, ArgTypes)
    If cmi Is Nothing Then Call Err.Raise(438)

    Dim hr As Long
    Dim v() As Variant
    v = Arguments
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRInvokeMethodWithTypes2, ToObject(cmi).Invoke_3(obj, v))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Calls the static method
Public Function CLRInvokeStaticMethod(ByVal t As mscorlib.Type, ByVal MethodName As String, ParamArray Arguments() As Variant) As Variant
    Dim hr As Long
    Dim v() As Variant
    v = Arguments
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRInvokeStaticMethod, m_objManagedSupport.MyInvokeMember( _
        t, _
        MethodName, _
        BindingFlags_Public Or BindingFlags_Static Or BindingFlags_InvokeMethod, _
        Nothing, v _
    ))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Calls the static method with specified type.
'   ArgTypesAndArgs must be the following values:
'     index[0, 2, 4, ...]: Type instance(s)
'     index[1, 3, 5, ...]: parameter value(s)
Public Function CLRInvokeStaticMethodWithTypes(ByVal t As mscorlib.Type, ByVal MethodName As String, ParamArray ArgTypesAndArgs() As Variant) As Variant
    Dim ArgTypes() As Variant, Args() As Variant
    Dim c As Long, lb As Long, ub As Long, i As Long
    lb = LBound(ArgTypesAndArgs)
    ub = UBound(ArgTypesAndArgs)
    c = ub - lb + 1
    If c Mod 2 <> 0 Then
        Call Err.Raise(5)
    End If
    c = c / 2
    ReDim ArgTypes(0 To c - 1), Args(0 To c - 1)
    Dim j As Long
    j = lb
    For i = 0 To c - 1
        If Not IsType(ArgTypesAndArgs(j)) Then Call Err.Raise(5)
        Set ArgTypes(i) = ArgTypesAndArgs(j)
        Call VariantCopy(Args(i), ArgTypesAndArgs(j + 1))
        j = j + 2
    Next i

    Dim cmi As mscorlib.MethodInfo
    Set cmi = PickupMethodByParamTypeImpl(t, MethodName, BindingFlags_Public Or BindingFlags_Static, ArgTypes)
    If cmi Is Nothing Then Call Err.Raise(438)

    Dim hr As Long
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRInvokeStaticMethodWithTypes, ToObject(cmi).Invoke_3(Nothing, Args))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Calls the static method with specified type.
Public Function CLRInvokeStaticMethodWithTypes2(ByVal t As mscorlib.Type, ByVal MethodName As String, ByRef ArgTypes() As Variant, ParamArray Arguments() As Variant) As Variant
    Dim cmi As mscorlib.MethodInfo
    Set cmi = PickupMethodByParamTypeImpl(t, MethodName, BindingFlags_Public Or BindingFlags_Static, ArgTypes)
    If cmi Is Nothing Then Call Err.Raise(438)

    Dim hr As Long, v() As Variant
    v = Arguments
    ' 「Set」なしで値をコピー(ムーブ)できるように VariantCopy を用いる
    hr = VariantCopy(CLRInvokeStaticMethodWithTypes2, ToObject(cmi).Invoke_3(Nothing, v))
    If hr < 0 Then Call Err.Raise(hr)
End Function

' Parses the value into actual enum value
Public Function CLRParseEnum(ByVal EnumTypeName As String, ByVal Value As Variant) As Variant
    Dim tEnum As mscorlib.Type
    Set tEnum = CLRResolveType("System.Enum")
    Dim tTargetEnum As mscorlib.Type
    Set tTargetEnum = CLRResolveType(EnumTypeName)
    If VarType(Value) = vbString Then
        CLRParseEnum = CLRInvokeStaticMethod(tEnum, "Parse", tTargetEnum, Value)
    Else
        CLRParseEnum = CLRInvokeStaticMethod(tEnum, "ToObject", tTargetEnum, CLng(Value))
    End If
End Function

' Parses the value into actual enum value with Type instance
Public Function CLRParseEnumByType(ByVal tTargetEnum As mscorlib.Type, ByVal Value As Variant) As mscorlib.Object
    Dim tEnum As mscorlib.Type
    Set tEnum = CLRResolveType("System.Enum")
    If VarType(Value) = vbString Then
        Set CLRParseEnumByType = CLRInvokeStaticMethod(tEnum, "Parse", tTargetEnum, Value)
    Else
        Set CLRParseEnumByType = CLRInvokeStaticMethod(tEnum, "ToObject", tTargetEnum, CLng(Value))
    End If
End Function

' Creates the delegate instance which calls the method of specified instance
Public Function CLRCreateDelegate(ByVal typeDelegate As mscorlib.Type, ByVal Target As Object, ByVal MethodName As String) As mscorlib.Delegate
    Set CLRCreateDelegate = m_objManagedSupport.MyCreateDelegate(typeDelegate, Target, MethodName)
End Function

' Creates the delegate instance which calls the function
'   FuncPtr must be specified with AddressOf operator
' CAUTION: the signature of the function represented by FuncPtr must match
'          with the specified delegate type; inappropriate parameter types or
'          return type may cause the program crash.
Public Function CLRCreateDelegateWithFunction(ByVal typeDelegate As mscorlib.Type, ByVal FuncPtr As LongPtr) As mscorlib.Delegate
    Set CLRCreateDelegateWithFunction = m_objManagedSupport.MyCreateDelegateWithFunction(typeDelegate, FuncPtr)
End Function
