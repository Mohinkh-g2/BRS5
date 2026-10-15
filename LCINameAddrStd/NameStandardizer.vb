Imports System.Data.SqlClient
Imports System.IO
Public Class NameStandardizer

    '  Public Function Name_Standardize(ByVal nameDataRow As DataRow, ByVal GUID As String) As DataSet
    Public Function Name_Standardize(ByVal Name_Raw As String, ByVal Name_Input_Type_Code As String) As DataSet

        Dim namestdize_ds As New DataSet
        Dim TempName1 As String
        Dim TempName2 As String
        Dim TempName3 As String

        Dim TempStringHolder As String
        Dim Melissa_Data_Name_Std_Code As String
        Dim Net_Gender_Std_Code As String
        Dim Net_Gender_Quality_Code As String
        Dim Consumer_Name_Raw As String
        Dim Consumer_Name_Full As String
        Dim Consumer_Name_Raw_Input_Melissa As String
        Dim Consumer_Name_First_Raw As String
        Dim Consumer_Name_Middle_Raw As String
        Dim Consumer_Name_Last_Raw As String
        Dim Consumer_Name_Suffix_Raw As String
        Dim Consumer_Name_Format_Code As Integer
        Dim Consumer_Name_Prefix As String
        Dim Consumer_Name_First As String
        Dim Consumer_Name_Middle As String
        Dim Consumer_Name_Last As String
        Dim Consumer_Name_Suffix As String

        Dim dt As New DataTable
        Dim drow As DataRow

        Dim Input_Name_Type_Code As Integer

        Dim MDN_Flag As String = 0
        Dim Net_Gender_Flag As String = 0

        Dim rsRegExName As System.Text.RegularExpressions.Regex

        Dim namePtr As New NAMEOBJECTLib.NameParse

        Dim melissaLicenseString As String = MelissaLicense.MELISSA_LICENSE_STRING
        'namePtr.SetLicenseString("28-112-505")
        'namePtr.SetLicenseString("ITtSw9AAQ1iKULHiZDDgej==eUFeY5b7RdDMUglfI51QUJ==")
        namePtr.SetLicenseString(melissaLicenseString)
        'namePtr.Initialize("C:\Program Files\Melissa DATA\COM Objects\Data Files\Mp4fname.dat")
        namePtr.Initialize(MelissaLicense.MELISSA_NAMEEXE_FILEPATH)
        'Create an instance of the Net Gender Object
        Dim NetGender As New NetGender.NetGender
        NetGender.Company_Check = True
        NetGender.Name_Style = "COMPANY"
        NetGender.Gender_Confidence = "70"
        NetGender.Output_Case = "Mixed"
        NetGender.Static_Key_Name = "Lundquist Consulting, Inc. - nle@lundquistconsulting.com[1-Dev]"
        NetGender.Static_Key = "1B61-9F61-6A1D-FC4D"

        TempName1 = ""
        TempName2 = ""
        TempName3 = ""

        TempStringHolder = ""
        Melissa_Data_Name_Std_Code = ""
        Net_Gender_Std_Code = ""
        Net_Gender_Quality_Code = ""
        Consumer_Name_Raw = ""
        Consumer_Name_Full = ""
        Consumer_Name_Raw_Input_Melissa = ""
        Consumer_Name_First_Raw = ""
        Consumer_Name_Middle_Raw = ""
        Consumer_Name_Last_Raw = ""
        Consumer_Name_Suffix_Raw = ""

        Consumer_Name_Prefix = ""
        Consumer_Name_First = ""
        Consumer_Name_Middle = ""
        Consumer_Name_Last = ""
        Consumer_Name_Suffix = ""
        MDN_Flag = 0
        Net_Gender_Flag = 0

        Consumer_Name_Raw = Name_Raw
        Input_Name_Type_Code = Name_Input_Type_Code

        If Input_Name_Type_Code = 2201 Then 'Raw: F,M,L,S '1.2.1
            If (InStr(Consumer_Name_Raw, ",") > 0) Then
                Dim NameComponents() As String = Consumer_Name_Raw.Split(",")
                'Fill the values only if present in the input
                Dim _Count1 As Integer = NameComponents.Length

                If (_Count1 = 5) Then
                    Consumer_Name_Prefix = NameComponents(0)
                    Consumer_Name_Last = NameComponents(1)
                    Consumer_Name_First = NameComponents(2)
                    Consumer_Name_Middle = NameComponents(3)
                    Consumer_Name_Suffix = NameComponents(4)

                ElseIf (_Count1 = 4) Then
                    Consumer_Name_Prefix = ""
                    Consumer_Name_Last = NameComponents(0)
                    Consumer_Name_First = NameComponents(1)
                    Consumer_Name_Middle = NameComponents(2)
                    Consumer_Name_Suffix = NameComponents(3)

                ElseIf (_Count1 = 3) Then
                    Consumer_Name_Prefix = ""
                    Consumer_Name_Last = NameComponents(0)
                    Consumer_Name_First = NameComponents(1)
                    Consumer_Name_Middle = NameComponents(2)
                    Consumer_Name_Suffix = ""
                ElseIf (_Count1 = 2) Then
                    Consumer_Name_Prefix = ""
                    Consumer_Name_Last = NameComponents(0)
                    Consumer_Name_First = NameComponents(1)
                    Consumer_Name_Middle = ""
                    Consumer_Name_Suffix = ""
                End If

                Consumer_Name_Format_Code = 2502
                Consumer_Name_Full = Consumer_Name_Prefix & " " & Consumer_Name_First & " " & Consumer_Name_Middle & " " & Consumer_Name_Last & " " & _
                                        Consumer_Name_Suffix
                MDN_Flag = 0
                Net_Gender_Std_Code = ""
                Net_Gender_Quality_Code = ""
                If (namePtr.Parse(Trim(Consumer_Name_Full)) = 0) Then
                    'Melissa Name Failed
                    MDN_Flag += 1
                    Consumer_Name_Prefix = ""
                    Consumer_Name_First = ""
                    Consumer_Name_Middle = ""
                    Consumer_Name_Last = ""
                    Consumer_Name_Suffix = ""
                    Consumer_Name_Full = ""
                    Melissa_Data_Name_Std_Code = namePtr.StatusCode
                Else
                    MDN_Flag += 1
                    If namePtr.StatusCode = " " Or namePtr.StatusCode = "V" Then
                        Consumer_Name_Prefix = Trim(UCase(namePtr.Prefix))
                        Consumer_Name_First = Trim(UCase(namePtr.FirstName))
                        Consumer_Name_Middle = Trim(UCase(namePtr.MiddleName))
                        Consumer_Name_Last = Trim(UCase(namePtr.LastName))
                        Consumer_Name_Suffix = Trim(UCase(namePtr.Suffix))
                        Melissa_Data_Name_Std_Code = Trim(UCase(namePtr.StatusCode))
                        '    MDN_Error_Code = namePtr.ErrorCode
                        TempStringHolder = Consumer_Name_Prefix & " " & Consumer_Name_First & " " & Consumer_Name_Middle & " " & Consumer_Name_Last & " " & _
                                           Consumer_Name_Suffix
                        TempName1 = TempStringHolder.Replace(".", "")
                        TempName2 = TempName1.Replace(",", "")
                        TempName3 = TempName2.Replace(";", "")
                        Consumer_Name_Full = Trim(UCase(TempName3))

                        If Melissa_Data_Name_Std_Code = "" Then
                            Melissa_Data_Name_Std_Code = "G"
                        End If

                    Else
                        Consumer_Name_Prefix = ""
                        Consumer_Name_First = ""
                        Consumer_Name_Middle = ""
                        Consumer_Name_Last = ""
                        Consumer_Name_Suffix = ""
                        Consumer_Name_Full = ""
                        Melissa_Data_Name_Std_Code = Trim(UCase(namePtr.StatusCode))
                    End If
                End If
            End If
        End If


        If Input_Name_Type_Code = 2203 Then 'Broken Out: F M L S  '1.2.5

            MDN_Flag = 0
            If (InStr(Consumer_Name_Raw, " ") > 0) Then
                Dim NameComponents() As String = Consumer_Name_Raw.Split(" ")
                Dim _Count2 As Integer = NameComponents.Length
                'Fill the values only if present in the input
                If (_Count2 = 5) Then
                    Consumer_Name_Prefix = NameComponents(0)
                    Consumer_Name_Last = NameComponents(1)
                    Consumer_Name_First = NameComponents(2)
                    Consumer_Name_Middle = NameComponents(3)
                    Consumer_Name_Suffix = NameComponents(4)
                ElseIf (_Count2 = 4) Then
                    Consumer_Name_Prefix = ""
                    Consumer_Name_Last = NameComponents(0)
                    Consumer_Name_First = NameComponents(1)
                    Consumer_Name_Middle = NameComponents(2)
                    Consumer_Name_Suffix = NameComponents(3)

                ElseIf (_Count2 = 3) Then
                    Consumer_Name_Prefix = ""
                    Consumer_Name_First = NameComponents(0)
                    Consumer_Name_Middle = NameComponents(1)
                    Consumer_Name_Last = NameComponents(2)
                    Consumer_Name_Suffix = ""
                ElseIf (_Count2 = 2) Then
                    Consumer_Name_Prefix = ""
                    Consumer_Name_Last = NameComponents(0)
                    Consumer_Name_First = NameComponents(1)
                    Consumer_Name_Middle = ""
                    Consumer_Name_Suffix = ""
                End If

                If Consumer_Name_First = "" And Consumer_Name_Middle = "" And Consumer_Name_Suffix = "" And Consumer_Name_Last <> "" Then

                    Consumer_Name_Format_Code = 2503 ' Company

                Else
                    'Standardize using Melissa Data Name
                    Consumer_Name_Format_Code = 2501 '2501: Broken Out Name Components
                    Consumer_Name_Full = Consumer_Name_Prefix & " " & Consumer_Name_First & " " & Consumer_Name_Middle & " " & Consumer_Name_Last & " " & _
                                       Consumer_Name_Suffix
                    If (namePtr.Parse(Trim(Consumer_Name_Full)) = 0) Then
                        'Melissa Name Failed
                        MDN_Flag += 1
                        Consumer_Name_Prefix = ""
                        Consumer_Name_First = ""
                        Consumer_Name_Middle = ""
                        Consumer_Name_Last = ""
                        Consumer_Name_Suffix = ""
                        Consumer_Name_Full = ""
                        Melissa_Data_Name_Std_Code = Trim(UCase(namePtr.StatusCode))
                    Else
                        MDN_Flag += 1
                        If namePtr.StatusCode = " " Or namePtr.StatusCode = "V" Then
                            Consumer_Name_Prefix = Trim(UCase(namePtr.Prefix))
                            Consumer_Name_First = Trim(UCase(namePtr.FirstName))
                            Consumer_Name_Middle = Trim(UCase(namePtr.MiddleName))
                            Consumer_Name_Last = Trim(UCase(namePtr.LastName))
                            Consumer_Name_Suffix = Trim(UCase(namePtr.Suffix))
                            Melissa_Data_Name_Std_Code = Trim(UCase(namePtr.StatusCode))
                            TempStringHolder = Consumer_Name_Prefix & " " & Consumer_Name_First & " " & Consumer_Name_Middle & " " & Consumer_Name_Last & " " & _
                                               Consumer_Name_Suffix
                            TempName1 = TempStringHolder.Replace(".", "")
                            TempName2 = TempName1.Replace(",", "")
                            TempName3 = TempName2.Replace(";", "")
                            Consumer_Name_Full = Trim(UCase(TempName3))

                            If Melissa_Data_Name_Std_Code = "" Then
                                Melissa_Data_Name_Std_Code = "G"
                            End If
                        Else
                            Consumer_Name_Prefix = ""
                            Consumer_Name_First = ""
                            Consumer_Name_Middle = ""
                            Consumer_Name_Last = ""
                            Consumer_Name_Suffix = ""
                            Consumer_Name_Full = ""
                            Melissa_Data_Name_Std_Code = Trim(UCase(namePtr.StatusCode))
                        End If
                    End If
                End If
            Else
                'Company
                'Format the Input Name without commas,colons and periods
                'MDN_Flag += 1
                'TempName1 = Consumer_Name_Raw.Replace(".", " ")
                'TempName2 = TempName1.Replace(",", " ")
                'TempName3 = TempName2.Replace(";", " ")
                Consumer_Name_Format_Code = 2503
                Consumer_Name_Full = ""
                Melissa_Data_Name_Std_Code = "X"
            End If

            If Melissa_Data_Name_Std_Code = "" Then
                Melissa_Data_Name_Std_Code = "X"
            End If

        End If

        If Input_Name_Type_Code = 2204 Then 'Broken Out Company Name

            Net_Gender_Flag = 0
            'Check this
            Consumer_Name_Full = Consumer_Name_Raw

            NetGender.Name_In = Consumer_Name_Full
            NetGender.Parse()
            If NetGender.Gender = "C" And NetGender.Name_Quality = "High" Then 'Company
                Net_Gender_Flag += 1
                Consumer_Name_Format_Code = 2503
                Consumer_Name_Prefix = ""
                Consumer_Name_First = ""
                Consumer_Name_Middle = ""
                Consumer_Name_Last = ""
                Consumer_Name_Suffix = ""
                Consumer_Name_Full = Trim(UCase(NetGender.Company))
                Net_Gender_Std_Code = Trim(UCase(NetGender.Gender))
                Net_Gender_Quality_Code = Trim(UCase(NetGender.Name_Quality))
            Else
                'Unknown
                Net_Gender_Flag += 1
                NetGender.Company_Check = False
                NetGender.Name_Style = "FML"
                NetGender.Gender_Confidence = "70"
                NetGender.Output_Case = "Mixed"
                NetGender.Static_Key_Name = "Lundquist Consulting, Inc. - nle@lundquistconsulting.com[1-Dev]"
                NetGender.Static_Key = "1B61-9F61-6A1D-FC4D"
                Consumer_Name_Format_Code = 2502
                Consumer_Name_Prefix = ""
                Consumer_Name_First = ""
                Consumer_Name_Middle = ""
                Consumer_Name_Last = ""
                Consumer_Name_Suffix = ""
                Consumer_Name_Full = Consumer_Name_Full
                Net_Gender_Std_Code = Trim(UCase(NetGender.Gender))
                Net_Gender_Quality_Code = Trim(UCase(NetGender.Name_Quality))
            End If
        End If

        '********Code to remove extra white spaces
        rsRegExName = New System.Text.RegularExpressions.Regex("\s+")
        Consumer_Name_Full = rsRegExName.Replace(Consumer_Name_Full, " ").Trim()

        dt.Columns.Add("Consumer_Name_Prefix")
        dt.Columns.Add("Consumer_Name_First")
        dt.Columns.Add("Consumer_Name_Middle")
        dt.Columns.Add("Consumer_Name_Last")
        dt.Columns.Add("Consumer_Name_Suffix")
        dt.Columns.Add("Consumer_Name_Full")
        dt.Columns.Add("Consumer_Name_Format_Code")
        dt.Columns.Add("MDN_Flag")
        dt.Columns.Add("Melissa_Data_Name_Std_Code")
        dt.Columns.Add("Net_Gender_Flag")
        dt.Columns.Add("Net_Gender_Std_Code")
        dt.Columns.Add("Net_Gender_Quality_Code")

        drow = dt.NewRow()

        drow.Item("Consumer_Name_Prefix") = Consumer_Name_Prefix
        drow.Item("Consumer_Name_First") = Consumer_Name_First
        drow.Item("Consumer_Name_Middle") = Consumer_Name_Middle
        drow.Item("Consumer_Name_Last") = Consumer_Name_Last
        drow.Item("Consumer_Name_Suffix") = Consumer_Name_Suffix
        drow.Item("Consumer_Name_Full") = Consumer_Name_Full
        drow.Item("Consumer_Name_Format_Code") = Consumer_Name_Format_Code
        drow.Item("MDN_Flag") = MDN_Flag
        drow.Item("Melissa_Data_Name_Std_Code") = Melissa_Data_Name_Std_Code
        drow.Item("Net_Gender_Flag") = Net_Gender_Flag
        drow.Item("Net_Gender_Std_Code") = Net_Gender_Std_Code
        drow.Item("Net_Gender_Quality_Code") = Net_Gender_Quality_Code

        dt.Rows.Add(drow)
        namestdize_ds.Tables.Add(dt)


        Return namestdize_ds
        ' Next
    End Function
End Class
