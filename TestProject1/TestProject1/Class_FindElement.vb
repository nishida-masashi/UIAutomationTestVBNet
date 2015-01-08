Public Class Class_FindElement

    Private Sub New()
    End Sub

#Region "Find ID"

    ' AutomationID All
    Public Shared Function Find_AutomationID_ALL( _
        ByVal elem As System.Windows.Automation.AutomationElement,
        ByVal id As String
        ) As System.Windows.Automation.AutomationElement

        Dim retElem As System.Windows.Automation.AutomationElement
        Dim scope As System.Windows.Automation.TreeScope
        Dim condition As System.Windows.Automation.PropertyCondition
        Dim propert As System.Windows.Automation.AutomationProperty =
            System.Windows.Automation.AutomationElement.AutomationIdProperty

        scope = System.Windows.Automation.TreeScope.Element +
                System.Windows.Automation.TreeScope.Descendants

        condition = New System.Windows.Automation.PropertyCondition(propert, id)

        retElem = elem.FindFirst(scope, condition)

        Return retElem
    End Function

    ' AutomationID Children Only
    Public Shared Function Find_AutomationID_Children( _
        ByVal elem As System.Windows.Automation.AutomationElement,
        ByVal id As String
        ) As System.Windows.Automation.AutomationElement

        Dim retElem As System.Windows.Automation.AutomationElement
        Dim scope As System.Windows.Automation.TreeScope
        Dim condition As System.Windows.Automation.PropertyCondition
        Dim propert As System.Windows.Automation.AutomationProperty =
            System.Windows.Automation.AutomationElement.AutomationIdProperty

        scope = System.Windows.Automation.TreeScope.Children

        condition = New System.Windows.Automation.PropertyCondition(propert, id)

        retElem = elem.FindFirst(scope, condition)

        Return retElem
    End Function

    ' ClassName All
    Public Shared Function Find_ClassName_ALL( _
        ByVal elem As System.Windows.Automation.AutomationElement,
        ByVal id As String
        ) As System.Windows.Automation.AutomationElement

        Dim retElem As System.Windows.Automation.AutomationElement
        Dim scope As System.Windows.Automation.TreeScope
        Dim condition As System.Windows.Automation.PropertyCondition
        Dim propert As System.Windows.Automation.AutomationProperty =
            System.Windows.Automation.AutomationElement.ClassNameProperty

        scope = System.Windows.Automation.TreeScope.Element +
                System.Windows.Automation.TreeScope.Descendants

        condition = New System.Windows.Automation.PropertyCondition(propert, id)

        retElem = elem.FindFirst(scope, condition)

        Return retElem
    End Function

    ' ClassName Children Only
    Public Shared Function Find_ClassName_Children( _
        ByVal elem As System.Windows.Automation.AutomationElement,
        ByVal id As String
        ) As System.Windows.Automation.AutomationElement

        Dim retElem As System.Windows.Automation.AutomationElement
        Dim scope As System.Windows.Automation.TreeScope
        Dim condition As System.Windows.Automation.PropertyCondition
        Dim propert As System.Windows.Automation.AutomationProperty =
            System.Windows.Automation.AutomationElement.ClassNameProperty

        scope = System.Windows.Automation.TreeScope.Children

        condition = New System.Windows.Automation.PropertyCondition(propert, id)

        retElem = elem.FindFirst(scope, condition)

        Return retElem
    End Function

#End Region

#Region "Find Other"

    ' Find dialog 
    Public Shared Function Find_Dialog( _
        ByVal elem As System.Windows.Automation.AutomationElement
        ) As System.Windows.Automation.AutomationElement

        Dim retElem As System.Windows.Automation.AutomationElement = Nothing

        While retElem Is Nothing
            retElem = Find_ClassName_Children(elem, "#32770")
        End While

        Return retElem
    End Function


#End Region

End Class
