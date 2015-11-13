Public Class which_form


    Public Sub New(ByVal whichform As Form, ByVal id As String)
        Dim str As String = whichform.Name.ToString


        Select Case str
            Case "Confirming"
                Dim x As New ConfirmingData
                x.PullCustomerINFO(Confirming.Tab, id)
                Confirming.BringToFront()
                Confirming.Focus()


            Case "Sales"
                Sales.BringToFront()
                Sales.Focus()

            Case "MarketingManager"
                MarketingManager.BringToFront()
                MarketingManager.Focus()

            Case "Installation"
                Installation.BringToFront()
                Installation.Focus()

            Case "WCaller"
                WCaller.BringToFront()
                WCaller.Focus()

            Case "Finance"
                Finance.BringToFront()
                Finance.Focus()

            Case "ConfirmingSingleRecord"
                ConfirmingSingleRecord.BringToFront()
                ConfirmingSingleRecord.Focus()

            Case "Administration"
                Administration.BringToFront()
                Administration.Focus()

            Case "Recovery"
                Recovery.BringToFront()
                Recovery.Focus()

            Case "PreviousCustomer"
                PreviousCustomer.BringToFront()
                PreviousCustomer.Focus()

            Case "ColdCalling"
                ColdCalling.BringToFront()
                ColdCalling.Focus()

            Case "SecondSource"
                SecondSource.BringToFront()
                SecondSource.Focus()



        End Select
    End Sub
End Class
