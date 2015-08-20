
Imports System.Data.SqlClient

Public Class ADODBNET


    Public Enum LockTypeEnum
        adLockUnspecified = -1
        adLockReadOnly = 1
        adLockPessimistic = 2
        adLockOptimistic = 3
        adLockBatchOptimistic = 4
    End Enum

    Public Enum CursorTypeEnum
        adOpenUnspecified = -1
        adOpenForwardOnly = 0
        adOpenKeyset = 1
        adOpenDynamic = 2
        adOpenStatic = 3
    End Enum

    Public Enum CursorLocationEnum
        adUseNone = 1
        adUseServer = 2
        adUseClient = 3
    End Enum

    Public Class Connection
        Public ConnectionString As String = String.Empty
        Public CommandTimeout As Integer = 30
        Public CursorLocation As CursorLocationEnum = CursorLocationEnum.adUseClient
        'Marcamos sqlconn como protected porque no está disponible en el ADO original, sino que es
        'un añadido de ADODBNET para operar más fácilmente. Por ello, sólo debería ser disponible
        'para esta clase y las que de su funcionamiento dependan
        Protected sqlconn As New SqlConnection()
        Public ReadOnly Property conn() As SqlConnection
            Get
                Return sqlconn
            End Get
        End Property


        Public Sub New(ByVal connString As String, ByVal commTimeout As Integer, ByVal cursLoc As CursorLocationEnum)
            Me.ConnectionString = connString
            Me.CommandTimeout = commTimeout
            Me.CursorLocation = cursLoc
        End Sub

        Public Sub New(ByVal connString As String)
            Me.New(connString, 30, CursorLocationEnum.adUseClient)
        End Sub

        Public Sub Open()
            Me.sqlconn.ConnectionString = Me.ConnectionString
            Me.sqlconn.Open()
        End Sub

        Public Sub Open(ByVal connString As String)
            Me.ConnectionString = connString
            Me.Open()
        End Sub

        Public Function isOpen() As Boolean
            Return (sqlconn.State = ConnectionState.Open)
        End Function


    End Class




    Public Class Command
        Public ActiveConnection As ADODBNET.Connection
        Public CommandText As String
        Public CommandTimeout As Integer


        Public Sub Execute()
            'En el nostre programa, anem configurant els paràmetros del Command sobre la marxa, així que no sabem que estan 
            'tots assignats fins que l'usuari li dona a Execute. Per això, serà ací quan configurem el objecte SQLCommand.
            Dim comm As New SqlCommand(CommandText)
            comm.Connection = ActiveConnection.conn
            comm.CommandTimeout = CommandTimeout

            'IMPORTANTE:     únicamente se pueden ejecutar sentencias que no devuelvan ningún resultado (INSERT, DELETE, UPDATE…). Esto se 
            'debe a que en ADO.NET el comando “Execute” que tenían los objetos ADODB.Command ha quedado separado en “ExecuteReader”, que nos 
            'devuelve el resultado almacenado en un DataReader; “ExecuteScalar”, que únicamente nos devuelve el primer resultado [primera fila, 
            'primera columna] de todos los obtenidos con la consulta; y “ExecuteNonQuery”, que ejecuta todas aquellas instrucciones que realizan
            'operaciones sobre la DB sin esperar ningún resultado de vuelta. En nuestro caso utilizamos ExecuteNonQuery(), dado que todas las veces
            'que llamábamos a Execute era para ejecutar sentencias de este tipo; no obstante, si se desease ejecutar cualquier otro tipo de instrucción
            'habría que cambiar antes el código de esta función. 
            comm.ExecuteNonQuery()
        End Sub

    End Class


    'Clase que servirá para representar las columnas que contenga nuestra tabla. 
    'En la versió original de ADO hi ha moltíssims més camps, però de moment nosaltres només necessitem eixos.
    'OJOALDATO: Açó soles ho gastem quan llegim o modifiquem els datos, és a dir, en un DataSet. Per tant, probablement la implementació siga més fàcil del que pareix.
    Public Class Field

        Public Name As String
        Public Type As Type
        Public DefinedSize As Integer
        Public rsDataAdapter As SqlDataAdapter
        Public rsDataSet As DataSet
        Public currentTable, currentRow, currentColumn As Integer
        Public ADO_CON As ADODBNET.Connection

        Property Value As Object
            Get
                Return rsDataSet.Tables(currentTable).Rows(currentRow).Item(Name)
            End Get

            Set(value As Object)
                'OJOALDATO: No basta con actualizar el DataSet, hay que aplicar el cambio en la DB.
                Try
                    'Guardamos el valor actual. Lo necesitaremos para la sentencia SQL.
                    Dim currVal As Object = rsDataSet.Tables(currentTable).Rows(currentRow).Item(Name)

                    'Abrimos el modo EDIT de la fila, cambiamos el valor en el DataSet y cerramos el modo EDIT
                    Me.rsDataSet.Tables(currentTable).Rows(currentRow).BeginEdit()
                    rsDataSet.Tables(currentTable).Rows(currentRow).Item(Name) = value
                    Me.rsDataSet.Tables(currentTable).Rows(currentRow).EndEdit()

                    'Variables para simplificar el comando UPDATE
                    Dim tabName As String = rsDataSet.Tables(currentTable).TableName
                    Dim colName As String = Me.Name

                    'Consulta SQL a ejecutar (En la tabla {tabName} fija a {value} el valor de la columna {colName} en las filas en las que {colName} valga actualmente {currVal}
                    Dim updateComm As String = "UPDATE " & tabName & " SET " & colName & "='" & value & "' WHERE " & colName & "='" & currVal & "'"

                    'Aplicamos el cambio
                    rsDataAdapter.UpdateCommand = New SqlCommand(updateComm, ADO_CON.conn)
                    rsDataAdapter.Update(rsDataSet, tabName)

                Catch ex As ReadOnlyException
                    Console.WriteLine("Los valores de esta columna no se pueden editar. La columna es de sólo lectura")
                    Me.rsDataSet.Tables(currentTable).Rows(currentRow).EndEdit()
                End Try
            End Set
        End Property


        Public Sub New(ByVal fieldName As String, ByVal fieldType As Type, ByVal value As Object, ByRef da As SqlDataAdapter, ByRef ds As DataSet, ByRef ADO_CON As ADODBNET.Connection, ByRef tablePosition As Integer, ByRef rowPosition As Integer, ByRef colPosition As Integer)
            Me.Name = fieldName
            Me.Type = fieldType
            Me.rsDataAdapter = da
            Me.rsDataSet = ds
            Me.currentTable = tablePosition
            Me.currentRow = rowPosition
            Me.currentColumn = colPosition

            'OJOALDATO: Value hay que settearlo SIEMPRE SIEMPRE SIEMPRE AL FINAL, porque depende de todos los demás y de no hacerlo así daría error
            Me.Value = value
            Me.DefinedSize = Len(Me.Value)
        End Sub

    End Class


    Public Class RecordSet
        Public rsDataAdapter As SqlDataAdapter
        Public rsDataSet As DataSet

        Public currentTable As Integer = 0 'OJOALDATO: 0: Tabla real. 1: Tabla de pruebas (Costumers)
        Public currentRow As Integer = 0
        Public currentColumn As Integer = 0

        'Les asignamos por defecto los tipos más restrictivos, porsiaca
        Public CursorLocation As CursorLocationEnum = CursorLocationEnum.adUseServer
        Public CursorType As CursorTypeEnum = CursorTypeEnum.adOpenForwardOnly
        Public LockType As LockTypeEnum = LockTypeEnum.adLockReadOnly

        'As it isn't part of the original ADODB Class, we'll define it as protected, so we can only use it internally
        Protected ADO_CON As ADODBNET.Connection

        Public ReadOnly Property FieldsCount As Integer 'Property que substituirà a Fields.Count
            Get
                'Returns the number of columns of the current row of the current table. 
                Return rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray.Length
            End Get
        End Property

        'Variable that indicates if we've reached the end of the table that we're reading
        Public ReadOnly Property EOF() As Boolean
            Get
                Dim numRows As Integer = rsDataSet.Tables(currentTable).Rows.Count
                Return (currentRow = numRows)
            End Get
        End Property



        Public Sub New()
            Me.CursorLocation = CursorLocationEnum.adUseServer
            Me.CursorType = CursorTypeEnum.adOpenDynamic
            Me.LockType = LockTypeEnum.adLockOptimistic
        End Sub


        Public Sub Open(ByVal consulta As String, ByVal ADO_CON As ADODBNET.Connection, ByVal cursLocation As CursorLocationEnum, ByVal cursType As CursorTypeEnum, ByVal lockType As LockTypeEnum)
            Me.CursorLocation = cursLocation
            Me.CursorType = cursType
            Me.LockType = lockType
            Me.ADO_CON = ADO_CON

            If Not Me.ADO_CON.isOpen() Then
                Me.ADO_CON.Open()
            End If
            createDataSet(consulta, ADO_CON)

        End Sub

        Public Sub Open(ByVal consulta As String, ByVal ADO_CON As ADODBNET.Connection)
            Open(consulta, ADO_CON, Me.CursorLocation, Me.CursorType, Me.LockType)
        End Sub

        'Function that creates a DataSet executing the SQL Query {consulta} within the DB specified
        'in the {conn} object.
        Public Sub createDataSet(ByVal consulta As String, ByVal ADO_CON As ADODBNET.Connection)
            Me.rsDataSet = New DataSet()
            Me.rsDataAdapter = New SqlDataAdapter(consulta, ADO_CON.conn)
            Me.rsDataAdapter.FillSchema(rsDataSet, SchemaType.Source, "IFIFPC")
            Me.rsDataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Me.rsDataAdapter.Fill(Me.rsDataSet, "IFIFPC")
        End Sub

        Public Function Fields(ByVal columnName As String) As Field
            Dim item As Object = Me.rsDataSet.Tables(currentTable).Rows(currentRow).Item(columnName)
            Dim f As New Field(columnName, Me.rsDataSet.Tables(currentTable).Rows(currentRow).Item(columnName).GetType, item, rsDataAdapter, rsDataSet, Me.ADO_CON, currentTable, currentRow, currentColumn)
            Return f
        End Function

        Public Sub MoveNext()
            currentRow += 1
        End Sub














        'ALL THE CODE WRITTEN BEYOND THIS LINE HAS BEEN MADE ONLY FOR TESTING PURPOSES, AND
        'DOESN'T REPRESENT ANY FUNCTION OR METHOD IN THE ORIGINAL ADODB LIBRARY

        '-------------------------------------------------------------------------------------------------
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '
        '


        'Reads all the data stored in the DataSet of our ADONET.RecordSet object.
        'It must have been initialized first, otherwise this will throw an Exception.
        Private Sub ReadConsoleDS(ByVal justShowTitles As Boolean)

            'READING VALUES TEST 
            'Writing column title
            For Me.currentColumn = 0 To (Me.rsDataSet.Tables(currentTable).Columns.Count - 1)
                Dim columnTitle As String
                columnTitle = Me.rsDataSet.Tables(currentTable).Columns(Me.currentColumn).Caption
                Console.Write(columnTitle & " ")
            Next
            Console.WriteLine()

            If Not justShowTitles Then
                'Getting data values
                Do Until EOF
                    'Array containing the items in the current row of the table
                    Dim itemsInRow As Object() = Me.rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray
                    For currentColumn = 0 To (itemsInRow.Count - 1)
                        Console.Write(itemsInRow(currentColumn) & " ")
                    Next
                    Console.WriteLine()
                    MoveNext()
                Loop
            End If


            '----------------------------------------------------

            'EDITING VALUES TEST

            ''Getting just the last row
            'Do Until EOF
            '    MoveNext()
            'Loop
            'currentRow = currentRow - 1 'Getting back from EOF
            ' ''Dim itemsInRow As DataRow = Me.rsDataSet.Tables(currentTable).Rows(currentRow)
            ''Console.WriteLine(Me.rsDataSet.Tables(currentTable).Rows(currentRow).Item("ifpcobse"))
            ''Me.rsDataSet.Tables(currentTable).Rows(currentRow).BeginEdit()
            'Me.Fields("ifpcobse").Value = "a017"
            ''Me.rsDataSet.Tables(currentTable).Rows(currentRow).EndEdit()

            ''Console.WriteLine(Me.rsDataSet.Tables(currentTable).Rows(currentRow).Item("ifpcobse"))
            ' ''Me.rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray(Me.rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray.Count - 1) = "a717"
            ' ''Me.rsDataSet.AcceptChanges()
            ''For currentColumn = 0 To (Me.rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray.Count - 1)
            ''    Console.Write(Me.rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray(currentColumn) & " ")
            ''Next
            ''Console.WriteLine()

            'rsDataSet.Clear()
            'Me.Open("SELECT * FROM IFIFPC", ADO_CON)
            'Dim itemsInRow As DataRow = Me.rsDataSet.Tables(currentTable).Rows(currentRow)
            'Console.WriteLine(itemsInRow.Item("ifpcobse"))
            'For currentColumn = 0 To (Me.rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray.Count - 1)
            '    Console.Write(itemsInRow.ItemArray(currentColumn) & " ")
            'Next
            'Console.WriteLine()


        End Sub



        Private Sub CreateCostumers()

            Dim table As New DataTable("Costumers")

            ' Create two columns, ID and Name.
            Dim idColumn As DataColumn = table.Columns.Add("ID", GetType(Integer))
            table.Columns.Add("Name", GetType(String))

            ' Set the ID column as the primary key column.
            table.PrimaryKey = New DataColumn() {idColumn}

            table.Rows.Add(New Object() {1, "Mary"})
            table.Rows.Add(New Object() {2, "Andy"})
            table.Rows.Add(New Object() {3, "Peter"})
            table.Rows.Add(New Object() {4, "Russ"})

            For Each column As DataColumn In table.Columns
                Console.Write(column.Caption & " ")

            Next
            Console.WriteLine()

            'Getting data values
            For Each row As DataRow In table.Rows
                For Each item As Object In row.ItemArray
                    Console.Write(item & " ")
                Next
                Console.WriteLine()
            Next

            Me.rsDataSet.Tables.Add(table)

        End Sub

        Private Sub ShowCostumers()
            Dim table As DataTable = Me.rsDataSet.Tables("Costumers")

            For Each column As DataColumn In table.Columns
                Console.Write(column.Caption & " ")

            Next
            Console.WriteLine()

            'Getting data values
            For Each row As DataRow In table.Rows
                For Each item As Object In row.ItemArray
                    Console.Write(item & " ")
                Next
                Console.WriteLine()
            Next
        End Sub

    End Class


End Class
