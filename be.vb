Imports System
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Math
Imports System.String
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlClient.SqlException
Imports System.Data.SqlClient.SqlError
Imports System.Runtime.InteropServices
Imports System.ArgumentException

Public Class Form1
    Public Shared radio_total As Integer
    Public Shared appendd() As String
    Public Shared radio() As String
    Public Shared arr_distinct(10000) As Integer
    Public Shared arr6(10000) As String
    Public init As Integer
    Public Shared csv() As String
    Dim flg As Boolean
    Dim chkflag As Boolean
    Dim count As Integer
    Dim posx As Integer
    Dim posy As Integer
    Dim append As String
    Public Shared cc As Integer
    Dim x As Integer
    Dim sCommand As SqlCommand
    Dim sCommand1 As SqlCommand
    Dim sAdapter As SqlDataAdapter
    Dim sBuilder As SqlCommandBuilder
    Dim sDs As DataSet
    Dim sTable As DataTable
    Public Shared total1 As Integer
    Dim temp1 As String
    Dim f As Integer
    Dim f1 As Integer
    Dim val2 As String
    Public Shadows num_of_cols As Integer
    Dim dr As SqlDataReader
    Dim val3 As Integer
    Dim powset As Integer
    Dim p As Integer
    Dim val As String
    Dim rev As String
    Dim f2 As Integer
    Dim init1 As Integer
    Dim init2 As Integer
    Dim temp6 As String
    Dim k As Integer
    Dim pp As Integer
    Dim pa As Integer
    Dim ap As Integer
    Dim aa As Integer
    Dim check As Boolean
    Dim check1 As Boolean
    Dim size1 As Integer
    Dim w As Integer
    Public Shared tablename As String
    Dim connectionString As String
    Public Shared connection As SqlConnection
    Dim sql33 As String
    Public Shared arr5(10000) As String 'stores all the col names 
    Dim checkBox(10000) As CheckBox

    Public Sub conn()

        connectionString = "Data Source=SAYALI-1B156B51;Initial Catalog=beproject;Integrated Security=true;"
        connection = New SqlConnection(connectionString)
        Try
            connection.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Function sql_grid_execute(ByVal sqlst As String, ByVal tablename As String) As SqlDataReader

        Try
            sCommand = New SqlCommand(sqlst, connection)
            sAdapter = New SqlDataAdapter(sCommand)
            sBuilder = New SqlCommandBuilder(sAdapter)
            sDs = New DataSet()
            sAdapter.Fill(sDs, tablename)
            sTable = sDs.Tables(tablename)
        Catch ex As Exception
            MsgBox(ex.Message & " Error Connecting to database!")
        End Try
        Return dr

    End Function

    Public Sub grid(ByRef tablename1 As String)

        DataGrid1.DataSource = sDs.Tables(tablename1)
        DataGrid1.ReadOnly = True
        Try
            DataGrid1.DataSource = DataGridViewSelectionMode.FullRowSelect
        Catch ex As Exception
            '  MsgBox(ex.Message & " Error Connecting to database!")
        End Try
		' MsgBox("grid")

    End Sub
    
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        flg = False
        chkflag = False
        conn()

    End Sub

    Public Sub sql_execute(ByVal sql33 As String, ByVal tablename As String)
        
        sCommand = New SqlCommand(sql33, connection)
        sAdapter = New SqlDataAdapter(sCommand)
        sBuilder = New SqlCommandBuilder(sAdapter)
        Try
            sCommand.ExecuteNonQuery()
        Catch ex As Exception
        End Try
        
    End Sub
    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        tablename = ComboBox1.SelectedItem.ToString

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        
        If IsNullOrEmpty(tablename) Then
            MsgBox("Select the table name ")
            Exit Sub
        End If
        sql33 = "SELECT * FROM " & tablename & ""
        sql_grid_execute(sql33, tablename)
        grid(tablename)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        
        flg = True
        If IsNullOrEmpty(tablename) Then
            MsgBox("Select the table name ")
            Exit Sub
        End If
        Dim count1 As Integer
        Dim temp5 As String
        count1 = 0
        temp5 = ""
        temp6 = ""
        sql33 = "select COUNT(*) from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" & tablename & "'"
        sCommand = New SqlCommand(sql33, connection)
        Try
            dr = sCommand.ExecuteReader()
            'reading from the datareader
            dr.Read()
            'displaying the data from the table
            num_of_cols = dr.GetValue(0)
            'MsgBox(num_of_cols.ToString)
        Catch ex As Exception
            MsgBox(ex.Message & " Error Connecting to database!")
        End Try
        dr.Close()
        name_array(num_of_cols)

    End Sub

    Private Sub name_array(ByVal num_of_cols As Integer)
        
        ReDim arr5(num_of_cols)
        ReDim arr6(num_of_cols)
        Dim k As Integer
        k = 1
        While k <= num_of_cols
            Dim sql2 As String = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" & tablename & "' and ORDINAL_POSITION = '" & k & "'"
            sCommand = New SqlCommand(sql2, connection)
            Try
                dr = sCommand.ExecuteReader()
                While dr.Read()
                    'reading from the datareader
                    'displaying the data from the table
                    temp6 = dr(0).ToString()
                    arr5(k) = temp6
                End While
                dr.Close()
            Catch ex As Exception
            End Try
            k = k + 1
        End While
        append = ""
        Dim tp(num_of_cols) As String
        
        For i = 2 To k - 2 Step 1
            append += arr5(i) + ","
        Next
        
        Dim i5 As Integer
        i5 = 2
        For j = 0 To num_of_cols - 2
            arr6(j) = arr5(i5)

        i5 += 1
        Next
        
        append = append.Remove(append.LastIndexOf(","), 1)
        chckarray(arr5)
        
    End Sub

    Private Sub MyCheckboxes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim p As New Integer
        chkflag = True
        append = ""
        cc = 0
        p = 2
        While (p < (num_of_cols))

            If checkBox(p).Checked = True Then
                arr6(cc) = checkBox(p).Text
                append = append + arr6(cc) + ","
                cc = cc + 1
            End If
            
            p = p + 1
        End While
        '  MsgBox("Value of cc " + cc.ToString)
        ' MsgBox(append)

    End Sub

    Private Sub csv1(ByRef append As String, ByRef cc As Integer)

        Dim t9 As Integer
        t9 = 0
        Dim countsql As String
        countsql = "select COUNT(*) from  " & tablename & ""
        dr = sql_grid_execute(countsql, tablename)
        sCommand = New SqlCommand(countsql, connection)
        
        Try
            dr = sCommand.ExecuteReader()
            dr.Read()
            t9 = dr.GetValue(0)
            ' MsgBox(animals(f2 - 1))
            dr.Close()
        Catch ex As Exception
        End Try

        total1 = t9
        '   MsgBox("t9 " + t9.ToString)
        Dim len As Integer
        Dim t As Integer
        ' MsgBox("total1 " + total1.ToString)
        Dim animals(total1) As String
        ReDim csv(total1)
        t = 1
        len = 0
        
        If Not chkflag Then
            len = append.Length
            append = append.Substring(0, len)
            cc = num_of_cols - 2
            ' MsgBox(cc.ToString)
            '  MsgBox(append)
        Else
            len = append.Length
            ' MsgBox(len.ToString)
            append = append.Substring(0, len - 1)
            ' MsgBox(append)
        End If
       
        Dim f2 As Integer
        f2 = 0
        Dim buffer As String
        Dim buffer1 As String
        buffer1 = ""
        buffer = ""
        ' MsgBox(append)
        'MsgBox("cc " + cc.ToString)
        Dim sql6p As String = "select " + arr5(1) + " from " & tablename & ""
        'MsgBox(sql6p)
        dr = sql_grid_execute(sql6p, tablename)
        sCommand = New SqlCommand(sql6p, connection)
        
        Try
            dr = sCommand.ExecuteReader()
            
            While dr.Read()
                animals(f2) = dr.GetValue(0).ToString
                'MsgBox(dr.GetString(0))
                f2 = f2 + 1
            End While

            ' MsgBox(animals(f2 - 1))
            f2 = 0
            dr.Close()
        Catch ex As Exception

        End Try
      
        ReDim appendd(total1)
        Dim buf As String
        Dim poin As Integer
        Dim t1 As Integer
        t1 = 0
        Dim fc As Integer
        buf = ""
        poin = 0
        t1 = 0
        While poin < total1
            Dim sqlp As String = "select " + append + " from " & tablename & " where " + arr5(1) + "='" + animals(poin) + "'"
            ' MsgBox(sqlp)
            sCommand = New SqlCommand(sqlp, connection)
            Try
                t1 = 0
                dr = sCommand.ExecuteReader()
                fc = dr.FieldCount
                ' MsgBox("fc " + fc.ToString)
                dr.Read()
                While (fc > 0)

                    buf += dr.GetName(t1) + dr.GetValue(t1).ToString + ","
                    'MsgBox("buf " + buf + " t1 " + t1.ToString)
                    fc = fc - 1
                    t1 += 1

                End While

                ' MsgBox("buf gf  " + t1.ToString)
                appendd(poin) = buf
                ' MsgBox(appendd(poin) + " " + poin.ToString + " fkldjglk ")
                poin = poin + 1
                buf = ""
                dr.Close()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        End While

        Dim tp1 As Integer
        tp1 = 0
        While tp1 < total1
            
            ' MsgBox(appendd(tp1) + " b4 ")
            appendd(tp1) = appendd(tp1).Remove(appendd(tp1).LastIndexOf(","), 1)
            ' MsgBox(" aftre " + appendd(tp1))
            tp1 += 1

        End While
       
        buffer = ""
        f2 = 0
        append = append + "," + arr5(num_of_cols)

        While f2 <= total1
        
            Dim sql6pp As String = "select " + append + " from " & tablename & " where " + arr5(1) + "='" + animals(f2) + "'"
            sCommand = New SqlCommand(sql6pp, connection)
            Try
            
                dr = sCommand.ExecuteReader()
                While dr.Read()
                
                    t = 2
                    While (t <= (cc + 2))
                        
                        If (t = cc + 2) Then
                            buffer = buffer + dr.GetValue(t - 2).ToString
                            t = t + 1
                        Else
                            buffer = buffer + dr.GetName(t - 2) + "" + dr.GetValue(t - 2).ToString + ","
                            t = t + 1
                        End If

                    End While

                End While

                csv(f2) = buffer
                f2 = f2 + 1
                buffer = ""
                dr.Close()
            Catch ex As Exception
                'MsgBox(ex.ToString)
            End Try

        End While
        
        Dim sql1016 As String = "delete from final1"
        sql_execute(sql1016, "final1")
        'MsgBox(sql1016)
        Dim i1 As Integer
        i1 = 0
        Dim sql6ppp As String
        sql6ppp = ""
        
        While i1 < total1
        
            sql6ppp = "insert into final1 values('" + animals(i1) + "','" + csv(i1) + "')"
            sCommand1 = New SqlCommand(sql6ppp, connection)
            Try
                sCommand1.ExecuteNonQuery()
            Catch ex As Exception
            End Try
            
            i1 = i1 + 1

        End While

        sql33 = "select * from final1"
        sql_grid_execute(sql33, "final1")
        grid("final1")

    End Sub

    Public Sub chckarray(ByVal ParamArray arr5() As String)

        posx = 42
        posy = 217
        x = 0
        
        For i = 2 To num_of_cols - 1

            checkBox(i) = New CheckBox()
            Me.Controls.Add(checkBox(i))
            checkBox(i).Location = New Point(posx, posy)
            checkBox(i).Text = arr5(i)
            ' MsgBox(arr5(i))
            checkBox(i).Checked = True
            checkBox(i).Size = New Size(81, 20)
            AddHandler checkBox(i).CheckedChanged, AddressOf MyCheckboxes_CheckedChanged
            posx += 93
            x += 1
            If (x > 3) Then
                posy += 35
                posx = 42
                x = 0
            End If

        Next

    End Sub

    Public Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        For i = 2 To num_of_cols - 1
            ' AddHandler checkBox(i).CheckedChanged, AddressOf MyCheckboxes_CheckedChanged
        Next

        If IsNullOrEmpty(tablename) Then

            MsgBox("Select the table name ")
            Exit Sub

        End If
        
        If Not flg Then
            MsgBox("Error...click on CSV button")
            Exit Sub
        End If

        csv1(append, cc)
        Dim sql64ttt As String
        Dim q As Integer
        q = 0
        Dim q1 As Integer
        q1 = 0
        radio_total = 0

        While (q < (cc))
            
            sql64ttt = "select count (distinct " + arr6(q) + ") from " & tablename & ""
            'MsgBox(sql64ttt)
            sCommand = New SqlCommand(sql64ttt, connection)
            
            Try

                dr = sCommand.ExecuteReader()
            Catch ex As Exception
            End Try
           
            While dr.Read()
                arr_distinct(q1) = dr.GetSqlInt32(0)
            End While
            
            radio_total += arr_distinct(q1)
            q1 += 1
            q += 1
            dr.Close()
            
        End While

        q = 0
        q1 = 0
        dr.Close()
        ReDim radio(radio_total)

        While (q < cc)
            Dim sql64tt As String = "select distinct " + arr6(q) + " from " & tablename & ""
            ' MsgBox("q" + q.ToString + "qry" + sql64tt)
            sCommand = New SqlCommand(sql64tt, connection)

            Try
                dr = sCommand.ExecuteReader()
            Catch ex As Exception

            End Try

            While (dr.Read())
                
                ' If (Not (dr.GetValue(0).ToString.Equals("Nil"))) Then
                radio(q1) = dr.GetName(0) + dr.GetValue(0).ToString
                '   MsgBox(radio(q1))
                q1 += 1
                ' End If
                
            End While
            dr.Close()
            q += 1
        End While

        Form2.Show()

    End Sub
   
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        flg = True
        Me.Dispose()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        tablename = ComboBox1.SelectedItem.ToString

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
