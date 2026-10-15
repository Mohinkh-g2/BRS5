Imports System.Data.SqlClient
Imports System.IO
Public Class AddrStandardizer

    ' Public Function Addrress_Standardize(ByVal addrDataRow As DataRow, ByVal GUID As String) As DataSet
    Dim addrptr As ADDRESSOBJECTLib.AddressCheck = Nothing

    Dim melissaLicenseString As String = ""

    Public Sub New()

        addrptr = New ADDRESSOBJECTLib.AddressCheck

        melissaLicenseString = MelissaLicense.MELISSA_LICENSE_STRING
        'Create an instance of the Melissa Address Object
        ' addrptr.SetLicenseString("2m-23m-8w9")
        ' addrptr.SetLicenseString("ITtSw9AAQ1iKULHiZDDgej==eUFeY5b7RdDMUglfI51QUJ==")

        addrptr.SetLicenseString(melissaLicenseString)
        'addrptr.PathToUSFiles = "C:\Program Files\Melissa DATA\COM Objects\Data Files\mdAddr.nat"
        addrptr.PathToUSFiles = MelissaLicense.MELISSA_ADDRESSEXE_FILEPATH
        addrptr.InitializeDataFiles()

    End Sub

    Public Sub Deinitialize()
        addrptr = Nothing

    End Sub

    Public Function Addrress_Standardize(ByVal _Address_Line_1 As String, ByVal _Address_Line_2 As String, _
                                        ByVal _City As String, ByVal _State As String, ByVal _Zip As String) As DataSet

        Dim addrstdize_ds As New DataSet
        Dim Addr_Formatter As New AddrFormatter
        Dim ds As New DataSet


        Dim Melissa_Data_Address_Std_Code As String
        Dim Melissa_Data_Address_Error_Code As String
        Dim Net_Address_Std_Code As String = ""

        Dim Consumer_Address_Line_1 As String
        Dim Consumer_Address_Line_2 As String
        Dim Consumer_Address_City As String
        Dim Consumer_Address_State As String
        Dim Consumer_Address_Zip As String

        Dim TempName1 As String
        Dim TempName2 As String
        Dim TempName3 As String


        Dim Original_Address As String
        Dim Consumer_Address_Raw_Line_1 As String
        Dim Consumer_Address_Raw_Line_2 As String
        Dim Consumer_Address_Raw_City As String
        Dim Consumer_Address_Raw_State As String
        Dim Consumer_Address_Raw_Zip As String
        Dim Consumer_Address_Range As String
        Dim Consumer_Address_Pre_Direction As String
        Dim Consumer_Address_Street_Name As String
        Dim Consumer_Address_Post_Direction As String
        Dim Consumer_Address_Suffix As String
        Dim Consumer_Address_PMB_Name As String
        Dim Consumer_Address_PMB_Number As String
        Dim Consumer_Address_Suite_Name As String
        Dim Consumer_Address_Suite_Number As String
        Dim TempAddress As String
        Dim Consumer_Address_County_Name As String

        Dim Consumer_Address_Format_Code As Integer
        Dim Postal_Format_Code As Integer

        Dim MDA_Flag As String
        Dim NA_Flag As String


        Dim dt As New DataTable
        Dim drow As DataRow

        ''Create an instance of the Melissa Address Object
        'Dim addrptr As New ADDRESSOBJECTLib.AddressCheck
        'addrptr.SetLicenseString("2m-23m-8w9")
        'addrptr.PathToUSFiles = "D:\Program Files\Melissa DATA\COM Objects\Data Files\mdAddr.nat"
        'addrptr.InitializeDataFiles()




        '*******Clearing all  values
        Melissa_Data_Address_Std_Code = ""
        Melissa_Data_Address_Error_Code = ""
        Net_Address_Std_Code = ""

        Consumer_Address_Line_1 = ""
        Consumer_Address_Line_2 = ""
        Consumer_Address_City = ""
        Consumer_Address_State = ""
        Consumer_Address_Zip = ""

        TempName1 = ""
        TempName2 = ""
        TempName3 = ""


        Original_Address = ""
        Consumer_Address_Raw_Line_1 = ""
        Consumer_Address_Raw_Line_2 = ""
        Consumer_Address_Raw_City = ""
        Consumer_Address_Raw_State = ""
        Consumer_Address_Raw_Zip = ""
        Consumer_Address_Range = ""
        Consumer_Address_Pre_Direction = ""
        Consumer_Address_Street_Name = ""
        Consumer_Address_Post_Direction = ""
        Consumer_Address_Suffix = ""
        Consumer_Address_PMB_Name = ""
        Consumer_Address_PMB_Number = ""
        Consumer_Address_Suite_Name = ""
        Consumer_Address_Suite_Number = ""
        TempAddress = ""
        Consumer_Address_County_Name = ""
        '*****************

        'Get Address values


        Consumer_Address_Raw_Line_1 = _Address_Line_1
        Consumer_Address_Raw_Line_2 = _Address_Line_2
        Consumer_Address_Raw_City = _City
        Consumer_Address_Raw_State = _State
        Consumer_Address_Raw_Zip = _Zip

        MDA_Flag = 0
        NA_Flag = 0

        Original_Address = Trim(Replace(Consumer_Address_Raw_Line_1 & " " & Consumer_Address_Raw_Line_2, ".", ""))
        'Clear the Address Object
        addrptr.ClearProperties()
        'Set the Address property
        addrptr.Address = Original_Address
        'Set the City property
        addrptr.City = Consumer_Address_Raw_City
        'Set the State property
        addrptr.State = Consumer_Address_Raw_State
        'Set the Zip property
        addrptr.Zip = Consumer_Address_Raw_Zip

        If (addrptr.VerifyAddress() = 1) Then  'Melissa Data Address Passes
            MDA_Flag += 1
            Consumer_Address_Range = Trim(UCase(addrptr.ParsedAddressRange))
            Consumer_Address_Pre_Direction = Trim(UCase(addrptr.ParsedPreDirection))
            Consumer_Address_Street_Name = Trim(UCase(addrptr.ParsedStreetName))
            Consumer_Address_Post_Direction = Trim(UCase(addrptr.ParsedPostDirection))
            Consumer_Address_Suffix = Trim(UCase(addrptr.ParsedSuffix))
            Consumer_Address_PMB_Name = Trim(UCase(addrptr.ParsedPrivateMailboxName))
            Consumer_Address_PMB_Number = Trim(UCase(addrptr.ParsedPrivateMailboxNumber))
            Consumer_Address_Suite_Name = Trim(UCase(addrptr.ParsedSuiteName))
            Consumer_Address_Suite_Number = Trim(UCase(addrptr.ParsedSuiteRange))
            Consumer_Address_City = Trim(UCase(addrptr.City))
            Consumer_Address_State = Trim(UCase(addrptr.State))
            Consumer_Address_Zip = Trim(UCase(addrptr.Zip)) + Trim(UCase(addrptr.Plus4))
            Consumer_Address_County_Name = Trim(UCase(addrptr.CountyName))
            Melissa_Data_Address_Std_Code = Trim(UCase(addrptr.StatusCode))
            Melissa_Data_Address_Error_Code = Trim(UCase(addrptr.ErrorCode))

        Else
            'Check for Secondary Standardization Preference
            '?????????????????????????
            'Initialize Net Address

            'Create an instance of the Net Address Object
            Dim NetAddress As New NetAddress.NetAddress
            NetAddress.Static_Key_Name = "Lundquist Consulting, Inc. - nle@lundquistconsulting.com[1-Dev]"
            NetAddress.Static_Key = "1862-A9AF-CC0D-298D"

            NA_Flag += 1
            NetAddress.Clear()
            NetAddress.Address_In = Original_Address
            NetAddress.CSZ_In = Consumer_Address_Raw_City & ", " & Consumer_Address_Raw_State & " " & Consumer_Address_Raw_Zip
            NetAddress.Output_Case = "Mixed"
            NetAddress.Numeric_Street_Conv = True
            NetAddress.CSZ_InSameField = True
            NetAddress.Parse()

            If NetAddress.Address_Type = "N" Then
                NA_Flag += 1
                'Net Address Failed (Not a Valid Address)
                Consumer_Address_Range = ""
                Consumer_Address_Pre_Direction = ""
                Consumer_Address_Street_Name = ""
                Consumer_Address_Post_Direction = ""
                Consumer_Address_Suffix = ""
                Consumer_Address_PMB_Name = ""
                Consumer_Address_PMB_Number = ""
                Consumer_Address_Suite_Name = ""
                Consumer_Address_Suite_Number = ""
                Consumer_Address_City = Trim(UCase(Consumer_Address_Raw_City))
                Consumer_Address_State = Trim(UCase(Consumer_Address_Raw_State))
                Consumer_Address_Zip = Trim(UCase(Consumer_Address_Raw_Zip))
                Consumer_Address_County_Name = ""
                Net_Address_Std_Code = Trim(UCase(NetAddress.Address_Type))
                Melissa_Data_Address_Std_Code = Trim(UCase(addrptr.StatusCode))
                Melissa_Data_Address_Error_Code = Trim(UCase(addrptr.ErrorCode))
            Else
                'Net Address Passes
                NA_Flag += 1
                TempAddress = NetAddress.Address_Out
                Original_Address = TempAddress
                'Clear the Address Object

                addrptr.ClearProperties()
                'Set the Address property
                addrptr.Address = Original_Address
                'Set the City property
                addrptr.City = Consumer_Address_Raw_City
                'Set the State property
                addrptr.State = Consumer_Address_Raw_State
                'Set the Zip property
                addrptr.Zip = Consumer_Address_Raw_Zip
                'Verify the Address
                If (addrptr.VerifyAddress() <> 1) Then
                    MDA_Flag += 1
                    'Melissa Data Address Fails
                    Consumer_Address_Range = Trim(UCase(addrptr.ParsedAddressRange))
                    Consumer_Address_Pre_Direction = Trim(UCase(addrptr.ParsedPreDirection))
                    Consumer_Address_Street_Name = Trim(UCase(addrptr.ParsedStreetName))
                    Consumer_Address_Post_Direction = Trim(UCase(addrptr.ParsedPostDirection))
                    Consumer_Address_Suffix = Trim(UCase(addrptr.ParsedSuffix))
                    Consumer_Address_PMB_Name = Trim(UCase(addrptr.ParsedPrivateMailboxName))
                    Consumer_Address_PMB_Number = Trim(UCase(addrptr.ParsedPrivateMailboxNumber))
                    Consumer_Address_Suite_Name = Trim(UCase(addrptr.ParsedSuiteName))
                    Consumer_Address_Suite_Number = Trim(UCase(addrptr.ParsedSuiteRange))
                    Consumer_Address_City = Trim(UCase(addrptr.City))
                    Consumer_Address_State = Trim(UCase(addrptr.State))
                    Consumer_Address_Zip = Trim(UCase(addrptr.Zip)) + addrptr.Plus4
                    Consumer_Address_County_Name = Trim(UCase(addrptr.CountyName))
                    Net_Address_Std_Code = Trim(UCase(NetAddress.Address_Type))
                    Melissa_Data_Address_Std_Code = Trim(UCase(addrptr.StatusCode))
                    Melissa_Data_Address_Error_Code = Trim(UCase(addrptr.ErrorCode))
                Else
                    MDA_Flag += 1
                    'Melissa Data Address Passes
                    Consumer_Address_Range = Trim(UCase(addrptr.ParsedAddressRange))
                    Consumer_Address_Pre_Direction = Trim(UCase(addrptr.ParsedPreDirection))
                    Consumer_Address_Street_Name = Trim(UCase(addrptr.ParsedStreetName))
                    Consumer_Address_Post_Direction = Trim(UCase(addrptr.ParsedPostDirection))
                    Consumer_Address_Suffix = Trim(UCase(addrptr.ParsedSuffix))
                    Consumer_Address_PMB_Name = Trim(UCase(addrptr.ParsedPrivateMailboxName))
                    Consumer_Address_PMB_Number = Trim(UCase(addrptr.ParsedPrivateMailboxNumber))
                    Consumer_Address_Suite_Name = Trim(UCase(addrptr.ParsedSuiteName))
                    Consumer_Address_Suite_Number = Trim(UCase(addrptr.ParsedSuiteRange))
                    Consumer_Address_City = Trim(UCase(addrptr.City))
                    Consumer_Address_State = Trim(UCase(addrptr.State))
                    Consumer_Address_Zip = Trim(UCase(addrptr.Zip)) + addrptr.Plus4
                    Consumer_Address_County_Name = Trim(UCase(addrptr.CountyName))
                    Net_Address_Std_Code = Trim(UCase(NetAddress.Address_Type))
                    Melissa_Data_Address_Std_Code = Trim(UCase(addrptr.StatusCode))
                    Melissa_Data_Address_Error_Code = Trim(UCase(addrptr.ErrorCode))

                End If
            End If
            NetAddress = Nothing
        End If

        'addrptr = Nothing



        ds = Addr_Formatter.Format_AddrStd(Consumer_Address_Range, Consumer_Address_Pre_Direction, _
                         Consumer_Address_Street_Name, Consumer_Address_Suffix, _
                         Consumer_Address_Post_Direction, Consumer_Address_Suite_Name, _
                           Consumer_Address_PMB_Name, Consumer_Address_PMB_Number, _
                           Consumer_Address_Suite_Number, Consumer_Address_City, _
                          Consumer_Address_State)

        Consumer_Address_Line_1 = ds.Tables(0).Rows(0).Item("Consumer_Address_Line_1")
        Consumer_Address_Line_2 = ds.Tables(0).Rows(0).Item("Consumer_Address_Line_2")
        Consumer_Address_Format_Code = ds.Tables(0).Rows(0).Item("Consumer_Address_Format_Code")
        Postal_Format_Code = ds.Tables(0).Rows(0).Item("Postal_Format_Code")


        dt.Columns.Add("Consumer_Address_Range")
        dt.Columns.Add("Consumer_Address_Pre_Direction")
        dt.Columns.Add("Consumer_Address_Street_Name")
        dt.Columns.Add("Consumer_Address_Suffix")
        dt.Columns.Add("Consumer_Address_Post_Direction")
        dt.Columns.Add("Consumer_Address_Suite_Name")
        dt.Columns.Add("Consumer_Address_Suite_Number")
        dt.Columns.Add("Consumer_Address_PMB_Name")
        dt.Columns.Add("Consumer_Address_PMB_Number")
        dt.Columns.Add("Consumer_Address_City")
        dt.Columns.Add("Consumer_Address_State")
        dt.Columns.Add("Consumer_Address_Zip")
        dt.Columns.Add("Consumer_Address_CountyName")
        dt.Columns.Add("Consumer_Address_Line_1")
        dt.Columns.Add("Consumer_Address_Line_2")
        dt.Columns.Add("Consumer_Address_Format_Code")
        dt.Columns.Add("MDA_Flag")
        dt.Columns.Add("Melissa_Data_Address_Std_Code")
        dt.Columns.Add("Melissa_Data_Address_Error_Code")
        dt.Columns.Add("NA_Flag")
        dt.Columns.Add("Net_Address_Std_Code")

        drow = dt.NewRow()



        drow.Item("Consumer_Address_Range") = Consumer_Address_Range
        drow.Item("Consumer_Address_Pre_Direction") = Consumer_Address_Pre_Direction
        drow.Item("Consumer_Address_Street_Name") = Consumer_Address_Street_Name
        drow.Item("Consumer_Address_Suffix") = Consumer_Address_Suffix
        drow.Item("Consumer_Address_Post_Direction") = Consumer_Address_Post_Direction
        drow.Item("Consumer_Address_Suite_Name") = Consumer_Address_Suite_Name
        drow.Item("Consumer_Address_Suite_Number") = Consumer_Address_Suite_Number
        drow.Item("Consumer_Address_PMB_Name") = Consumer_Address_PMB_Name
        drow.Item("Consumer_Address_PMB_Number") = Consumer_Address_PMB_Number
        drow.Item("Consumer_Address_City") = Consumer_Address_City
        drow.Item("Consumer_Address_State") = Consumer_Address_State
        drow.Item("Consumer_Address_Zip") = Consumer_Address_Zip
        drow.Item("Consumer_Address_CountyName") = Consumer_Address_County_Name
        drow.Item("Consumer_Address_Line_1") = Consumer_Address_Line_1
        drow.Item("Consumer_Address_Line_2") = Consumer_Address_Line_2
        drow.Item("Consumer_Address_Format_Code") = Consumer_Address_Format_Code
        drow.Item("MDA_Flag") = MDA_Flag
        drow.Item("Melissa_Data_Address_Std_Code") = Melissa_Data_Address_Std_Code
        drow.Item("Melissa_Data_Address_Error_Code") = Melissa_Data_Address_Error_Code
        drow.Item("NA_Flag") = NA_Flag
        drow.Item("Net_Address_Std_Code") = Net_Address_Std_Code

        dt.Rows.Add(drow)
        addrstdize_ds.Tables.Add(dt)

        dt.Dispose()
        ds.Clear()

        Return addrstdize_ds



    End Function



End Class
