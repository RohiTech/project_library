Imports Microsoft.VisualBasic

Namespace NombreSistema.Estructura
    Public Class Categories
        Dim vCategoryID As Integer

        Public Property CategoryID() As Integer
            Get
                Return vCategoryID
            End Get
            Set(ByVal value As Integer)
                vCategoryID = value
            End Set
        End Property

        Dim vCategoryName As String

        Public Property CategoryName() As String
            Get
                Return vCategoryName
            End Get
            Set(ByVal value As String)
                vCategoryName = value
            End Set
        End Property

        Dim vDescription As String

        Public Property Description() As String
            Get
                Return vDescription
            End Get
            Set(ByVal value As String)
                vDescription = value
            End Set
        End Property

        Dim vPicture As String

        Public Property Picture() As String
            Get
                Return vPicture
            End Get
            Set(ByVal value As String)
                vPicture = value
            End Set
        End Property

    End Class
End Namespace
