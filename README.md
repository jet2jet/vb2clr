# vb2clr

The helper class `CLRHost` for Visual Basic for Applications (VBA) 7.0, providing access to CLR (.NET Framework) assemblies and classes.

## Requirements

Visual Basic for Application 7.0 (included in Microsoft Office 2010 or higher)

* ***(not tested)*** To use on Visual Basic 6.0, rewrite `LongPtr` to `Long` and remove all `PtrSafe` specifiers.

## Usage

1. Import [CLRHost.cls](./CLRHost.cls) and [ExitHandler.bas](./ExitHandler.bas) into your VB/VBA project
2. Add Type Library reference 'Common Runtime Language Execution Engine' (maybe version 2.4) and 'mscorlib.dll'
3. Write your code using `CLRHost` class

## Notes and Warnings

* You should release the `CLRHost` instance or call `Terminate` method when you finish using CLR.
  * Unexpected behavior may occur due to living CLR instances if you don't release or terminate them.
* If you pass `True` to `TerminateOnExit` parameter of `CLRHost.Initialize`, you must not stop the debugger when breaking or pausing the application.
  * The code in `ExitHandler` module cannot be run when stopped during pausing, and the application (including VBA host such as Excel) may cause crash.
* Encodings of VB files are Shift-JIS; if you have problem with encodings, check `.utf8.*` files, remove Japanese comments, and import them.

## Example

```
Public Sub RegexSample()
    Dim host As New CLRHost
    Call host.Initialize(False)

    On Error Resume Next
    Dim asmSys As mscorlib.Assembly
    Set asmSys = host.CLRLoadAssembly("System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")

    Dim cobjRegex As mscorlib.Object
    Set cobjRegex = host.CLRCreateObjectWithParams("System.Text.RegularExpressions.Regex", _
        "([0-9])+")

    Dim cobjColl As mscorlib.Object
    Set cobjColl = host.CLRInvokeMethod(cobjRegex, "Matches", "10 20 50 1234 98765")

    Dim vMatch As Variant
    For Each vMatch In host.ToEnumerable(cobjColl)
        Dim cobjMatch As mscorlib.Object
        Set cobjMatch = vMatch
        Debug.Print "Matches: "; host.CLRProperty(cobjMatch, "Value")
        Set cobjMatch = Nothing
    Next vMatch
    vMatch = Empty
    Set cobjColl = Nothing
    Set cobjRegex = Nothing

    'Call host.Terminate
    Set host = Nothing
End Sub
```

## More details

- (In Japanese) [Using CLR(.NET) from VB](https://www.pg-fl.jp/program/tips/vb2clr1.htm)

## Author

jet (@jet2jet)

## License

[New BSD License (or The 3-Clause BSD License)](./LICENSE)
