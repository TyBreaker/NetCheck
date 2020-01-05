Imports System.ComponentModel

Public Class PingBGW
    Inherits BackgroundWorker

    ' Just wanted to be able to address a thread by name!
    ' Hence, this class which inherits the BackgroundWorker and adds a string identifier.

    Private ID As String = String.Empty

    Public Property pingID() As String
        Get
            Return ID
        End Get
        Set(value As String)
            ID = value
        End Set
    End Property

    Sub New(ByVal siteID As String)
        ID = siteID
    End Sub

End Class
