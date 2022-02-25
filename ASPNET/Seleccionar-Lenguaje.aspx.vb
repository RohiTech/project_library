Imports System.Threading
Imports System.Globalization

Partial Class Seleccionar_Lenguaje
    Inherits System.Web.UI.Page


    Public Const LanguageDropDownID As String = "ddlLan"
    Public Const PostBackEventTarget As String = "__EVENTTARGET"

    Protected Overrides Sub InitializeCulture()

        If Request(PostBackEventTarget) IsNot Nothing Then
            Dim controlID As String = Request(PostBackEventTarget)

            If controlID.Equals(LanguageDropDownID) Then
                Dim selectedValue As String = Request.Form(Request(PostBackEventTarget)).ToString()

                Select Case selectedValue
                    Case "es-NI"
                        SetCulture("es-NI", "es-NI")
                        Exit Select
                    Case "en-US"
                        SetCulture("en-US", "en-US")
                        Exit Select
                    Case "fr"
                        SetCulture("fr-FR", "fr-FR")
                        Exit Select
                    Case Else
                        Exit Select
                End Select
            End If
        End If


        If Session("MyUICulture") IsNot Nothing AndAlso Session("MyCulture") IsNot Nothing Then
            Thread.CurrentThread.CurrentUICulture = DirectCast(Session("MyUICulture"), CultureInfo)
            Thread.CurrentThread.CurrentCulture = DirectCast(Session("MyCulture"), CultureInfo)
        End If
        MyBase.InitializeCulture()
    End Sub


    Protected Sub SetCulture(ByVal name As String, ByVal locale As String)
        Thread.CurrentThread.CurrentUICulture = New CultureInfo(name)
        Thread.CurrentThread.CurrentCulture = New CultureInfo(locale)

        Session("MyUICulture") = Thread.CurrentThread.CurrentUICulture
        Session("MyCulture") = Thread.CurrentThread.CurrentCulture
    End Sub



End Class
