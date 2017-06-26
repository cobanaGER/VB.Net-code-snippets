Public Class NeueKlasse
    ''' <summary>
    ''' <seealso cref="TestFunction"/>
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub TestSub()

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sDatum"></param>
    ''' <para>use <see cref="TestSub"/></para>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TestFunction(ByRef sDatum As Date) As String
        Dim sname As String = Nothing
        Try
            Return sname
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' <para></para>
    ''' <seealso> cref=" TestFunction"/></seealso>
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub xTest()

    End Sub
    ''' <param name="id">The ID of the record to update.</param>
    ''' <remarks>Updates the record <paramref name="id"/>.
    ''' <para>Use <see cref="DoesRecordExist"/> to verify that
    ''' the record exists before calling this method.</para>
    ''' </remarks>
    Public Sub UpdateRecord(ByVal id As Integer)
        ' Code goes here.
    End Sub
    ''' <param name="id">The ID of the record to check.</param>
    ''' <returns><c>True</c> if <paramref name="id"/> exists,
    ''' <c>False</c> otherwise.</returns>
    ''' <remarks><seealso cref="UpdateRecord"/></remarks>
    Public Function DoesRecordExist(ByVal id As Integer) As Boolean
        ' Code goes here.
    End Function
    ''' <summary>
    '''  Gets the collection of source XML files.  See <see cref="DataTable"/>
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ttest()

    End Sub
End Class
