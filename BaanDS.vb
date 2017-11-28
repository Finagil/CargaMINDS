Partial Class BaanDS
   

    Partial Class SugerenciaBaanDataTable

        Private Sub SugerenciaBaanDataTable_SugerenciaBaanRowChanging(ByVal sender As System.Object, ByVal e As SugerenciaBaanRowChangeEvent) Handles Me.SugerenciaBaanRowChanging

        End Sub

    End Class

    Partial Class PagosBaanDataTable


        Private Sub PagosBaanDataTable_ColumnChanging(ByVal sender As System.Object, ByVal e As System.Data.DataColumnChangeEventArgs) Handles Me.ColumnChanging
            If (e.Column.ColumnName = Me.serieColumn.ColumnName) Then
                'Add user code here
            End If

        End Sub

    End Class

End Class
