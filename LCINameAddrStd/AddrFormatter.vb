Imports System.IO
Public Class AddrFormatter
    Public Structure Address_Formatted
        Public Address_Formatted_Line_1 As String
        Public Address_Formatted_Line_2 As String
    End Structure


    Public Function Format_AddrStd(ByVal Consumer_Address_Range As String, ByVal Consumer_Address_Pre_Direction As String, _
                              ByVal Consumer_Address_Street_Name As String, ByVal Consumer_Address_Suffix As String, _
                              ByVal Consumer_Address_Post_Direction As String, ByVal Consumer_Address_Suite_Name As String, _
                              ByVal Consumer_Address_PMB_Name As String, ByVal Consumer_Address_PMB_Number As String, _
                              ByVal Consumer_Address_Suite_Number As String, ByVal Consumer_Address_City As String, _
                              ByVal Consumer_Address_State As String) As DataSet

        Dim dt As New DataTable
        Dim drow As DataRow
        Dim formataddr_ds As New DataSet
        Dim Consumer_Address_Line_1 As String = ""
        Dim Consumer_Address_Line_2 As String = ""
        Dim Consumer_Address_Format_Code As Integer
        Dim Postal_Format_Code As Integer
        ' Dim ret_Address_Formatted As Address_Formatted
        '********Code to remove extra white spaces
        Dim rsRegEx As System.Text.RegularExpressions.Regex

        'Format the Address
        If Consumer_Address_State = "PR" Then

            Consumer_Address_Format_Code = 2402 'PR Address

            If (InStr(Consumer_Address_Street_Name, "PO")) = 1 Or (InStr(Consumer_Address_Street_Name, "Caller")) = 1 Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name + " " + Consumer_Address_Range
                Consumer_Address_Line_2 = " "
                Postal_Format_Code = 3006 ' PO Box
            ElseIf (InStr(Consumer_Address_Street_Name, "General Delivery")) = 1 Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name
                Consumer_Address_Line_2 = " "
                Postal_Format_Code = 3005 ' General Delivery
            ElseIf (InStr(Consumer_Address_Street_Name, "RR")) = 1 Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name + " " + Consumer_Address_Suite_Name + " " + Consumer_Address_Range
                Consumer_Address_Line_2 = " "
                Postal_Format_Code = 3004 ' Rural Address
            Else
                Consumer_Address_Line_1 = Consumer_Address_Range + " " + Consumer_Address_Pre_Direction + " " + Consumer_Address_Street_Name + " " & _
                                               Consumer_Address_Suffix + " " + Consumer_Address_Post_Direction
                Consumer_Address_Line_2 = Consumer_Address_Suite_Name + " " + Consumer_Address_Suite_Number
                Postal_Format_Code = 3001 'Standard address

            End If

        ElseIf Consumer_Address_State = "AE" Or Consumer_Address_State = "AP" Or Consumer_Address_State = "AA" Then

            Consumer_Address_Format_Code = 2401 'U.S. military Address

            If (Consumer_Address_City = "APO" Or Consumer_Address_City = "FPO") And (InStr(Consumer_Address_Street_Name, "PSC") Or InStr(Consumer_Address_Street_Name, "CMR")) Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name + ", BOX" + Consumer_Address_Range
                'Consumer_Address_Line_1 = Consumer_Address_Range + " " + Consumer_Address_Pre_Direction + " " + Consumer_Address_Street_Name + " " & _
                '                            Consumer_Address_Suffix + " " + Consumer_Address_Post_Direction
                Consumer_Address_Line_2 = Consumer_Address_Suite_Name + " " + Consumer_Address_Suite_Number
                Postal_Format_Code = 3002 ' Military International
            Else
                Consumer_Address_Line_1 = Consumer_Address_Range + " " + Consumer_Address_Pre_Direction + " " + Consumer_Address_Street_Name + " " & _
                                               Consumer_Address_Suffix + " " + Consumer_Address_Post_Direction
                Consumer_Address_Line_2 = Consumer_Address_Suite_Name + " " + Consumer_Address_Suite_Number
                Postal_Format_Code = 3001 'Standard address

            End If
        ElseIf Len(Consumer_Address_State) = 2 And Consumer_Address_State <> "PR" Then
            'Non-PR, non-military U.S. Address

            Consumer_Address_Format_Code = 2401 'U.S. Address

            If (InStr(Consumer_Address_Street_Name, "PO")) = 1 Or (InStr(Consumer_Address_Street_Name, "Caller")) = 1 Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name + " " + Consumer_Address_Range
                Consumer_Address_Line_2 = " "
                Postal_Format_Code = 3006 ' PO Box
            ElseIf (InStr(Consumer_Address_Street_Name, "General Delivery")) = 1 Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name
                Consumer_Address_Line_2 = " "
                Postal_Format_Code = 3005 ' General Delivery
            ElseIf (InStr(Consumer_Address_Street_Name, "RR")) = 1 Then
                Consumer_Address_Line_1 = Consumer_Address_Street_Name + " " + Consumer_Address_Suite_Name + " " + Consumer_Address_Range
                Consumer_Address_Line_2 = " "
                Postal_Format_Code = 3004 ' Rural Address
            ElseIf Consumer_Address_PMB_Name <> " " And Consumer_Address_PMB_Name <> "" Then
                Consumer_Address_Line_1 = Consumer_Address_Range + " " + Consumer_Address_Pre_Direction + " " + Consumer_Address_Street_Name + " " & _
                                            Consumer_Address_Suffix + " " + Consumer_Address_Post_Direction
                Consumer_Address_Line_2 = Consumer_Address_Suite_Name + " " + Consumer_Address_Suite_Number + " " + Consumer_Address_PMB_Name + " " & _
                                            Consumer_Address_PMB_Number
                Postal_Format_Code = 3007 'PMB Address
            Else
                Consumer_Address_Line_1 = Consumer_Address_Range + " " + Consumer_Address_Pre_Direction + " " + Consumer_Address_Street_Name + " " & _
                                            Consumer_Address_Suffix + " " + Consumer_Address_Post_Direction
                Consumer_Address_Line_2 = Consumer_Address_Suite_Name + " " + Consumer_Address_Suite_Number

                Postal_Format_Code = 3001 ' Standard Address
            End If

        Else
            Consumer_Address_Line_1 = Consumer_Address_Range + " " + Consumer_Address_Pre_Direction + " " + Consumer_Address_Street_Name + " " & _
                                          Consumer_Address_Suffix + " " + Consumer_Address_Post_Direction
            Consumer_Address_Line_2 = Consumer_Address_Suite_Name + " " + Consumer_Address_Suite_Number + " " + Consumer_Address_PMB_Name + " " + Consumer_Address_PMB_Number
            Postal_Format_Code = 3008 'Foreign
            Consumer_Address_Format_Code = 2403 'Foreign

        End If

        If Consumer_Address_Line_1 = "" Or Consumer_Address_Line_1 Is Nothing Then
            Consumer_Address_Line_1 = " "
        End If

        If Consumer_Address_Line_2 = "" Or Consumer_Address_Line_2 Is Nothing Then
            Consumer_Address_Line_2 = " "
        End If

        '***********************Removing extra white spaces
        rsRegEx = New System.Text.RegularExpressions.Regex("\s+")

        'If Consumer_Address_Line_1 <> "" Or Not (Consumer_Address_Line_1 Is Nothing) Then
        Consumer_Address_Line_1 = rsRegEx.Replace(Consumer_Address_Line_1, " ").Trim()
        'End If

        'If Consumer_Address_Line_2 <> "" Or Not (Consumer_Address_Line_2 Is Nothing) Then
        Consumer_Address_Line_2 = rsRegEx.Replace(Consumer_Address_Line_2, " ").Trim()
        'End If


        dt.Columns.Add("Consumer_Address_Line_1")
        dt.Columns.Add("Consumer_Address_Line_2")
        dt.Columns.Add("Postal_Format_Code")
        dt.Columns.Add("Consumer_Address_Format_Code")

        drow = dt.NewRow()



        drow.Item("Consumer_Address_Line_1") = Consumer_Address_Line_1
        drow.Item("Consumer_Address_Line_2") = Consumer_Address_Line_2
        drow.Item("Postal_Format_Code") = Postal_Format_Code
        drow.Item("Consumer_Address_Format_Code") = Consumer_Address_Format_Code

        dt.Rows.Add(drow)
        formataddr_ds.Tables.Add(dt)

        dt.Dispose()


        Return formataddr_ds

    End Function







End Class
