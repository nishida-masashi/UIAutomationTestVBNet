Imports System.Text

<TestClass()>
Public Class UnitTest1

    Private _process As Process
    Private _form As System.Windows.Automation.AutomationElement = Nothing

    Private _button1 As System.Windows.Automation.AutomationElement
    Private _button1_click As System.Windows.Automation.InvokePattern

    <TestInitialize()>
    Public Sub TestInit()
        _process = Process.Start("WindowsApplication1.exe")
        While (_form Is Nothing)
            Try
                _form = System.Windows.Automation.AutomationElement.FromHandle(_process.MainWindowHandle)
            Catch ex As Exception
                _form = Nothing
            End Try
        End While

        _button1 = Class_FindElement.Find_AutomationID_ALL(_form, "Button1")
        _button1_click = _button1.GetCurrentPattern(System.Windows.Automation.InvokePattern.Pattern)
    End Sub

    <TestCleanup()>
    Public Sub TestFin()
        _process.Kill()
    End Sub

    <TestMethod()>
    Public Sub Test_1()
        _button1_click.Invoke()

        Dim _dialog As System.Windows.Automation.AutomationElement
        _dialog = Class_FindElement.Find_Dialog(_form)

        Dim _label As System.Windows.Automation.AutomationElement
        _label = Class_FindElement.Find_ClassName_ALL(_dialog, "Static")

        StringAssert.Contains(_label.Current.Name, ClassLibrary1.Class1.GetDLLString)
        Debug.Print(ClassLibrary1.Class1.GetDLLString)

        Dim _ok As System.Windows.Automation.AutomationElement
        _ok = Class_FindElement.Find_ClassName_ALL(_dialog, "Button")
        Dim _ok_click As System.Windows.Automation.InvokePattern
        _ok_click = _ok.GetCurrentPattern(System.Windows.Automation.InvokePattern.Pattern)

        _ok_click.Invoke()
    End Sub

   

End Class
