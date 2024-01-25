Attribute VB_Name = "Module1"
Sub Calculate_Loan_Schedule()
Attribute Calculate_Loan_Schedule.VB_ProcData.VB_Invoke_Func = "P\n14"
    Dim P As Currency
    Dim years As Integer
    Dim interest As Double
    Dim pmt As Currency
    Dim pDateObj As Date
    Dim startN As Integer
    Dim endN As Integer
    Dim rangeStr As String
    Dim monthNo As Integer
    Dim monthAddress As String
    Dim interestAddress As String
    Dim principleAddress As String
    Dim balanceAddress As String
    
    Dim interestPaid As Currency
    Dim principlePaid As Currency
    Dim balance As Currency
    

    P = Range("E2").Value
    interest = Range("E6").Value
    years = Range("E4").Value
    
    pmt = Calculate_Payment(P, interest, years)
    
    Range("E8").Value = pmt
    
    Make_Decision (pmt)
    
    Range("A19:I2000").ClearContents
    
    monthNo = 1
    Range("A19").Value = Range("I2").Value
    Range("C19").Value = monthNo
    
    interestPaid = Range("E2").Value * Range("I6").Value
    principlePaid = pmt - interestPaid
    balance = Range("E2").Value - principlePaid
    
    Range("E19").Value = Round(interestPaid, 1)
    Range("G19").Value = Round(principlePaid, 1)
    Range("I19").Value = Round(balance, 1)
    
    
    startN = 20
    endN = startN + Range("I4").Value - 2
    
    rangeStr = "A" & startN & ":A" & endN
    
    pDateObj = Range("A19").Value
    

    
    For Each dateObj In Range(rangeStr)
    
        pDateObj = DateAdd("m", 1, pDateObj)
        
        dateObj.Value = pDateObj
        
        monthNo = monthNo + 1
        
        monthAddress = "C" & (18 + monthNo)
        interestAddress = "E" & (18 + monthNo)
        principleAddress = "G" & (18 + monthNo)
        balanceAddress = "I" & (18 + monthNo)
        
        interestPaid = balance * Range("I6").Value
        principlePaid = pmt - interestPaid
        balance = balance - principlePaid
        
        Range(monthAddress).Value = monthNo
        Range(interestAddress).Value = Round(interestPaid, 1)
        Range(principleAddress).Value = Round(principlePaid, 1)
        Range(balanceAddress).Value = Round(balance, 1)
    
    Next
End Sub

Private Sub Make_Decision(payment As Integer)
    If payment > Range("I8").Value Then
        Range("E13").Value = "No Go"
    Else
        Range("E13").Value = "Go"
    End If
End Sub

Private Function Calculate_Payment(principle As Currency, interest As Double, years As Integer) As Currency
    Dim r As Double
    Dim n As Integer
    
    r = interest / 12
    n = years * 12
    
    Range("I6").Value = r
    Range("I4").Value = n
    
    Calculate_Payment = (principle * r * (1 + r) ^ n) / (((1 + r) ^ n) - 1)
End Function




