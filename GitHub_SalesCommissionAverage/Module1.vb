Module Module1
    '   Employee Information
    Private EmpName As String
    Private EmpNumber As String

    Private EmpSales As Decimal
    Private EmpCommission As Decimal

    '   Total Sales & Commission
    Private TotalSales As Decimal
    Private TotalEmployees As Decimal
    Private TotalCommission As Decimal

    '   Average Sales & Commission
    Private AverageSales As Decimal
    Private AverageCommission As Decimal

    '   File Variable
    Private CurrentRecord() As String
    Private Const CommissionRate As Decimal = 0.03
    Private CommissionFile As New Microsoft.VisualBasic.FileIO.TextFieldParser("COMSALES.txt")
    Sub Main()
        Call HouseKeeping()
        Do While Not (CommissionFile).EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Sub HouseKeeping()
        Call SetFileDelimiters()
        Call WriteHeadings()
    End Sub

    Sub SetFileDelimiters()
        CommissionFile.TextFieldType = FileIO.FieldType.Delimited
        CommissionFile.SetDelimiters(",")
    End Sub

    Sub WriteHeadings()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "Sales Commission Report")
        Console.WriteLine()
        Console.WriteLine(Space(5) & "Emp Number" & Space(10) & "Sales Person" & Space(12) & "Sales" & Space(6) & "Commission")
        Console.WriteLine()
    End Sub

    Sub ProcessRecords()
        Call ReadFile()
        Call DetailCalculation()
        Call Accumulation()
        Call WriteDetailLine()
    End Sub

    Sub ReadFile()

        CurrentRecord = CommissionFile.ReadFields()

        EmpName = CurrentRecord(1)
        EmpSales = CurrentRecord(2)
        EmpNumber = CurrentRecord(0)

    End Sub

    Sub DetailCalculation()
        EmpCommission = EmpSales * CommissionRate
    End Sub

    Sub Accumulation()
        TotalSales += EmpSales
        TotalEmployees += 1
        TotalCommission += EmpCommission
    End Sub

    Sub WriteDetailLine()
        Console.WriteLine(Space(5) &
                          EmpNumber.PadRight(11) &
                          Space(9) &
                          EmpName.PadRight(15) &
                          Space(5) &
                          EmpSales.ToString("c").PadLeft(9) &
                          Space(10) &
                          EmpCommission.ToString("N2").PadLeft(6))
    End Sub

    Sub EndOfJob()
        Call SummaryCalculation()
        Call SummaryOutput()
        Call CloseFile()
    End Sub

    Sub SummaryCalculation()
        AverageSales = TotalSales / TotalEmployees
        AverageCommission = TotalCommission / TotalEmployees
    End Sub

    Sub SummaryOutput()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(5) & "Totals:" & Space(32) & TotalSales.ToString("c").PadLeft(10) & Space(8) & TotalCommission.ToString("n").PadLeft(8))
        Console.WriteLine(Space(5) & "Averages for " & TotalEmployees.ToString().PadLeft(2) & " Employees:" & Space(14) & AverageSales.ToString.PadLeft(9) & Space(10) & AverageCommission.ToString("N2").PadLeft(6))
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(31) & "Press -ENTER- To Exit")
    End Sub

    Sub CloseFile()
        Console.ReadLine()
        CommissionFile.Close()
    End Sub

End Module
