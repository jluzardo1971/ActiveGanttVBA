Option Explicit On 

Friend Class clsDictionary
    Inherits DictionaryBase

    Private mp_lKey As Integer = 1

    Public Sub Add(ByVal Value As Integer, ByVal Key As String)
        MyBase.Dictionary.Add(Key, Value)
    End Sub

    Public Sub Add(ByVal Value As String)
        MyBase.Dictionary.Add(mp_lKey, Value)
        mp_lKey = mp_lKey + 1
    End Sub

    Public Function Contains(ByVal Key As String) As Boolean
        Return MyBase.Dictionary.Contains(Key)
    End Function

    Public ReadOnly Property StrItem(ByVal Index As Integer) As String
        Get
            Return MyBase.Dictionary.Item(Index)
        End Get
    End Property

    Default Public Property Item(ByVal Key As String) As Integer
        Get
            Return MyBase.Dictionary.Item(Key)
        End Get
        Set(ByVal Value As Integer)
            MyBase.Dictionary.Item(Key) = Value
        End Set
    End Property

End Class

