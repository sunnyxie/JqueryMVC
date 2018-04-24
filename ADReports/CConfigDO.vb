<Serializable()>
Public Class CConfigDO : Implements IEquatable(Of CConfigDO)
    Private mDefaultOptionIndex As Integer
    Private mADInfoUserListFile As String
    Private mADInfoUserMembersFile As String
    Private mADReportExcelFile As String


    ' Constructor
    Sub New()
        mDefaultOptionIndex = 2
        mADInfoUserListFile = "AD Info User List export.csv"
        mADInfoUserMembersFile = "AD Info Members All.csv"
        mADReportExcelFile = "ad filter.xls"
    End Sub

    Public Property ADInfoUserListFile As String
        Get
            Return mADInfoUserListFile
        End Get
        Set(value As String)
            mADInfoUserListFile = value
        End Set
    End Property

    Public Property ADInfoUserMembersFile As String
        Get
            Return mADInfoUserMembersFile
        End Get
        Set(value As String)
            mADInfoUserMembersFile = value
        End Set
    End Property

    Public Property ADReportExcelFile As String
        Get
            Return mADReportExcelFile
        End Get
        Set(value As String)
            mADReportExcelFile = value
        End Set
    End Property

    Protected Property DefaultOptionIndex As Integer
        Get
            Return mDefaultOptionIndex
        End Get
        Set(value As Integer)
            mDefaultOptionIndex = value
        End Set
    End Property

    Public Overloads Function Equals(other As CConfigDO) As Boolean _
        Implements IEquatable(Of CConfigDO).Equals
        If other Is Nothing Then Return False

        If Me.mADInfoUserListFile = other.mADInfoUserListFile AndAlso mADInfoUserMembersFile _
            = other.mADInfoUserMembersFile AndAlso mADReportExcelFile = other.mADReportExcelFile Then
            Return True
        End If

        Return False
    End Function

    Public Overloads Shared Operator =(a As CConfigDO, b As CConfigDO) As Boolean
        If a Is Nothing AndAlso b Is Nothing Then
            Return True
        End If

        Return a.Equals(b)
    End Operator

    Public Overloads Shared Operator <>(a As CConfigDO, b As CConfigDO) As Boolean
        Return Not a = b
    End Operator
End Class
