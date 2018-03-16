Public Class AppObjectBase(Of EntityType As {CNRLBusinessEntity, New})
    Inherits CNRLBusinessObject(Of EntityType)

#Region "Constructor/Destructor"

    Public Sub New()
        Me.DatabaseKey = Constants.Values.ApplicationKey
    End Sub

#End Region

#Region "Properties"

#End Region

#Region "Public Methods"

#End Region

#Region "Private/Protected Methods"

#End Region

End Class
