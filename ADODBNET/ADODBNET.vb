
Imports System.Data.SqlClient

Public Class ADODBNET

    ''' <summary>
    ''' Lock types available to use on the ADODBNET objects. They're an exact copy of the ADODB ones
    ''' Reference: http://www.w3schools.com/asp/prop_rs_locktype.asp
    ''' </summary>
    ''' <remarks><seealso cref="CursorLocationEnum"/> <seealso cref="CursorTypeEnum"/></remarks>
    Public Enum LockTypeEnum
        adLockUnspecified = -1
        adLockReadOnly = 1
        adLockPessimistic = 2
        adLockOptimistic = 3
        adLockBatchOptimistic = 4
    End Enum

    ''' <summary>
    '''Cursor types available to use on the ADODBNET objects. They're an exact copy of the ADODB ones
    ''' Reference: http://www.w3schools.com/asp/prop_rs_cursortype.asp
    ''' </summary>
    ''' <remarks><seealso cref="CursorLocationEnum"/> <seealso cref="LockTypeEnum"/></remarks>
    Public Enum CursorTypeEnum
        adOpenUnspecified = -1
        adOpenForwardOnly = 0
        adOpenKeyset = 1
        adOpenDynamic = 2
        adOpenStatic = 3
    End Enum
    ''' <summary>
    ''' Cursor locations available to use in ADODBNET objects. They're an exact copy of the ADODB ones
    ''' Reference: http://www.w3schools.com/asp/prop_rs_cursorlocation.asp
    ''' </summary>
    ''' <remarks><seealso cref="CursorTypeEnum"/> <seealso cref="LockTypeEnum"/></remarks>
    Public Enum CursorLocationEnum
        adUseNone = 1
        adUseServer = 2
        adUseClient = 3
    End Enum


    ''' <summary>
    ''' Class whose objects will be used to manage the connections to the DB. It's the upgrade of the
    ''' ADODB.Connection component, and contains all its fields and methods, besides some extra ones
    ''' that let us define more precisely the behaviour of the connection.
    ''' </summary>
    ''' <remarks>
    ''' <field name="ConnectionString" type="String"><see cref="Connection.ConnectionString"/></field>
    ''' <field name="CommandTimeout" type="Integer"><see cref="Connection.CommandTimeout"/></field>
    ''' <field name="sqlconn" type="SqlConnection">VB.NET native object that manages the connection to the DB. It sholudn't
    '''  be public, as it wasn't part of the original ADODB Connection component and shouldn't be modified
    '''  by this way. Nonetheless, it may be accessed from other components of the ADODBNET class, in order
    ''' to simplify the operations we're going to do. </field>
    ''' <field name="conn" type="SqlConnection"><see cref="Connection.conn"/></field>
    ''' </remarks>
    Public Class Connection

        ''' <summary name="ConnectionString" type="String">String that contains the info of the connection 
        ''' to the DB.</summary> <remarks>Must be initialized in order to establish the connection</remarks>
        Public ConnectionString As String = String.Empty
        ''' <summary name="CommandTimeout" type="Integer"> Integer representing the amount of time (in seconds) 
        ''' that we'llbe waiting after executing a Command before we abort the connection</summary>
        Public CommandTimeout As Integer = 30
        Public CursorLocation As CursorLocationEnum = CursorLocationEnum.adUseClient
        Protected sqlconn As New SqlConnection()
        ''' <summary>
        ''' Property that allows us to obtain the connection to the DB. It cannot be modified, so we avoid
        ''' malfunctioning caused by unauthorised changes during runtime.
        ''' </summary>
        ''' <value>Connection to the DB.</value>
        ''' <returns>SqlConnection object that contains the information about the current connection to the DB</returns>
        ''' <remarks>Outside this class, it should only be modified using the Open method. <see cref="Open"></see></remarks>
        Public ReadOnly Property conn() As SqlConnection
            Get
                Return sqlconn
            End Get
        End Property

        ''' <summary>
        ''' Constructor that allows us to define every parameter in the object
        ''' </summary>
        ''' <param name="connString">Value we'll give to the ConnectionString field <see cref="ConnectionString"/></param>
        ''' <param name="commTimeout">Value we'll give to the CommandTimeout field <see cref="CommandTimeout"/></param>
        ''' <param name="cursLoc">Value of the Cursor Location of the ADODBNET.Connection object. <see cref="CursorLocationEnum"/></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal connString As String, ByVal commTimeout As Integer, ByVal cursLoc As CursorLocationEnum)
            Me.ConnectionString = connString
            Me.CommandTimeout = commTimeout
            Me.CursorLocation = cursLoc
        End Sub

        ''' <summary>
        ''' Constructor that only defines the ConnectionString of our ADODBNET.Connection object. It sets CommandTimeout to 30
        ''' and CursorLocation to adUseClient (3)
        ''' </summary>
        ''' <param name="connString">Value we'll give to the ConnectionString field <see cref="ConnectionString"/></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal connString As String)
            Me.New(connString, 30, CursorLocationEnum.adUseClient)
        End Sub

        ''' <summary>
        ''' Opens the connection to the database. It uses the parameters we've define when creating the object.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Open()
            Me.sqlconn.ConnectionString = Me.ConnectionString
            Me.sqlconn.Open()
        End Sub

        ''' <summary>
        ''' Opens the connection to the database specified in the connString parameter, overriding the value we 
        ''' may have in the ConnectionString field.
        ''' </summary>
        ''' <param name="connString">Value we'll give to the ConnectionString field <see cref="ConnectionString"/></param>
        ''' <remarks></remarks>
        Public Sub Open(ByVal connString As String)
            Me.ConnectionString = connString
            Me.Open()
        End Sub

        ''' <summary>
        ''' Checks if the connection is currently opened
        ''' </summary>
        ''' <returns>True if the connection is opened and False otherwise</returns>
        ''' <remarks></remarks>
        Public Function isOpen() As Boolean
            Return (sqlconn.State = ConnectionState.Open)
        End Function
    End Class



    ''' <summary>
    ''' Class whose objects can execute commands on a determined database, specified by a Connection object
    ''' that we'll define. Nonetheless, those commands cannot expect any result in return; i.e. we can only execute the
    ''' ones that affect the DB elements (INSERT, UPDATE, DELETE...)
    ''' </summary>
    ''' <remarks>
    ''' <field name="ActiveConnection" type="ADODBNET.Connection"><see cref="Command.ActiveConnection"/></field>
    ''' <field name="CommandText" type="String"><see cref="Command.CommandText"/></field>
    ''' <field name="CommandTimeout" type="Integer"><see cref="Command.CommandTimeout"/></field>
    ''' </remarks>
    Public Class Command
        ''' <summary name="ActiveConnection" type="ADODBNET.Connection">Object that manages the connection to the DB</summary>
        ''' <remarks>The connection must be opened before executing the command.</remarks>
        Public ActiveConnection As ADODBNET.Connection
        ''' <summary name="CommandText" type="String">String containing the command we want to execute</summary>
        Public CommandText As String
        ''' <summary name="CommandTimeout" type="Integer"> Integer representing the amount of time (in seconds) 
        ''' that we'llbe waiting after executing a Command before we abort the connection</summary>
        Public CommandTimeout As Integer

        ''' <summary>
        ''' Method that begins the execution of the command on the DB. 
        ''' </summary>
        ''' <remarks>It doesn't take any parameter; instead, the values of the command and the connection are
        ''' given by the ones we've already specified.</remarks>
        Public Sub Execute()
            'En el nostre programa, anem configurant els paràmetros del Command sobre la marxa, així que no sabem que estan 
            'tots assignats fins que l'usuari li dona a Execute. Per això, serà ací quan configurem el objecte SQLCommand.
            Dim comm As New SqlCommand(CommandText)
            comm.Connection = ActiveConnection.conn
            comm.CommandTimeout = CommandTimeout

            'Obrim la connexió si no ho està ja.
            If Not Me.ActiveConnection.isOpen() Then
                Me.ActiveConnection.Open()
            End If

            'IMPORTANT: únicament podem executar sentències que no tornen ningun resultat (INSERT, DELETE, UPDATE…). Açó es 
            'deu a que en ADO.NET el comando “Execute” que tenien els objetos ADODB.Command ha quedat separat en “ExecuteReader”, que ens 
            'torna el resultat almacenat en un DataReader; “ExecuteScalar”, que únicament ens torna el primer resultat [primera fila, 
            'primera columna] de todos els que haguem obtés; y “ExecuteNonQuery”, que executa totes aquelles instruccions que realitzen
            'operacions sobre la DB sense esperar ningun resultat. En el nostre cas utilitzem ExecuteNonQuery(), perque totes les voltes
            'que cridàvem a Execute era per executar sentències de este últim tipo; no obstante, si volguéssim executar qualsevol altre 
            'tipus d'instrucció hauríem de canviar primer el código d'aquesta funció.
            comm.ExecuteNonQuery()
        End Sub

    End Class


    'Clase que servirá para representar las columnas que contenga nuestra tabla. 
    'En la versió original de ADO hi ha moltíssims més camps, però de moment nosaltres només necessitem eixos.
    'OJOALDATO: Açó soles ho gastem quan llegim o modifiquem els datos, és a dir, en un DataSet. Per tant, probablement la implementació siga més fàcil del que pareix.
    ''' <summary>
    ''' Class we'll use to represent the information we obtain from the cells of an ADODBNET.RecordSet. We can also update them through this objects. 
    ''' </summary>
    ''' <remarks>
    ''' <field name="Name" type="String"><see cref="Field.Name"/></field>
    ''' <field name="Type" type="Type"><see cref="Field.Type"/></field>
    ''' <field name="DefinedSize" type="Integer"><see cref="Field.DefinedSize"/></field>
    ''' <field name="rsDataAdapter" type="SqlDataAdapter"><see cref="Field.rsDataAdapter"/></field>
    ''' <field name="rsDataSet" type="DataSet"><see cref="Field.rsDataSet"/></field>
    ''' <field name="currentTable, currentRow, currentColumn" type="Integer"><see cref="Field.currentTable"/></field>
    ''' <field name="ADO_CON" type="ADODBNET.Connection"><see cref="Field.ADO_CON"/></field>
    ''' </remarks>
    Public Class Field

        ''' <summary> Name that the field has in the DB</summary>
        Public Name As String
        ''' <summary> Type of the data, so we can create variables or operate with it</summary>
        Public Type As Type
        ''' <summary> Size the object needs in memory</summary>
        ''' <remarks>IMPORTANT: When it comes to Strings/arrays/collections, the size isn't calculated properly. Will be update in future revisions</remarks>
        Public DefinedSize As Integer
        ''' <summary> DataAdapter that the RecordSet has used to connect to the DB. Needed for executing updates on it.</summary>
        Public rsDataAdapter As SqlDataAdapter
        ''' <summary> DataSet that represents the elements of the RecordSet. Needed for updating the DB.</summary>
        Public rsDataSet As DataSet
        ''' <summary> Integers that represent the position of the field </summary>
        Public currentTable, currentRow, currentColumn As Integer
        ''' <summary> Object that represents the connection to the DB. Is needed for executing updates on it.</summary>
        Public ADO_CON As ADODBNET.Connection

        ''' <summary>
        ''' Value of the field. If we change its value, it will authomatically apply the change to the Database. 
        ''' </summary>
        ''' <value>Value of the field</value>
        ''' <returns>Object of type Object with the value of the specified field in the DB.</returns>
        ''' <remarks>It returns the value as an Object type obj. We should cast it to the type we
        ''' need while using it. </remarks>
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

        ''' <summary>
        ''' Creates a new object of type Field, given its information on the DB table and
        ''' the objects that allow the Field object to directly connect to it
        ''' </summary>
        ''' <param name="fieldName">String with the field name</param>
        ''' <param name="fieldType">Type that the object should take</param>
        ''' <param name="value">Value of the field</param>
        ''' <param name="da">Data adapter. We use it to execute updates on the DB</param>
        ''' <param name="ds">Data Set. We use it to retrieve the data from the DB</param>
        ''' <param name="ADO_CON">Connection to the DB. We use it to execute updates on it.</param>
        ''' <param name="tablePosition">Current table</param>
        ''' <param name="rowPosition">Current row</param>
        ''' <param name="colPosition">Current column</param>
        ''' <remarks>It must set Value at the end, because it depends on the other fields of the class</remarks>
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

    ''' <summary>
    ''' Class whose objects store and represent the data we retrieve from the database in a
    ''' table-like format. Every cell in it can be treated as a Field object.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RecordSet
        ''' <summary> Data adapter that will be used to fill the RecordSet </summary>
        Public rsDataAdapter As SqlDataAdapter
        ''' <summary>Data set that will contain the information retrieved from the DB.</summary>
        Public rsDataSet As DataSet

        ''' <summary>Position in the DB </summary>
        Public currentTable As Integer = 0 'OJOALDATO: 0: Tabla real. 1: Tabla de pruebas (Costumers)
        ''' <summary>Position in the DB </summary>
        Public currentRow As Integer = 0
        ''' <summary>Position in the DB </summary>
        Public currentColumn As Integer = 0

        'Les asignamos los tipos por defecto de cada uno.
        ''' <summary>Location of the cursor in the connection </summary>
        Public CursorLocation As CursorLocationEnum = CursorLocationEnum.adUseServer
        ''' <summary> Type of the cursor</summary>
        Public CursorType As CursorTypeEnum = CursorTypeEnum.adOpenDynamic
        ''' <summary>Type of the Lock.</summary>
        Public LockType As LockTypeEnum = LockTypeEnum.adLockOptimistic

        'As it isn't part of the original ADODB Class, we'll define it as protected, so we can only use it internally
        ''' <summary>Connection that the RecordSet is using to access the DB.</summary>
        Protected ADO_CON As ADODBNET.Connection

        ''' <summary>
        ''' Read-Only property that counts the amount of fields in the rows of the RecordSet.
        ''' </summary>
        ''' <value>Amount of fields in the current row</value>
        ''' <returns>Integer representing the amount of fields in the current row</returns>
        Public ReadOnly Property FieldsCount As Integer 'Property que substituirà a Fields.Count
            Get
                'Returns the number of columns of the current row of the current table. 
                Return rsDataSet.Tables(currentTable).Rows(currentRow).ItemArray.Length
            End Get
        End Property

        'Variable that indicates if we've reached the end of the table that we're reading
        ''' <summary>
        ''' End Of File. Indicates if we've reached the end of the RecordSet.
        ''' </summary>
        ''' <value>Boolean indicating if we're positioned in the last row in the RecordSet</value>
        ''' <returns>True if there arent any rows left in the RecordSet, False otherwise</returns>
        Public ReadOnly Property EOF() As Boolean
            Get
                Dim numRows As Integer = rsDataSet.Tables(currentTable).Rows.Count
                Return (currentRow = numRows)
            End Get
        End Property

        ''' <summary>
        ''' Default constructor. Doesn't initialize any field.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.CursorLocation = CursorLocationEnum.adUseServer
            Me.CursorType = CursorTypeEnum.adOpenDynamic
            Me.LockType = LockTypeEnum.adLockOptimistic
        End Sub

        ''' <summary>
        ''' Connects to the DB and fills the RecordSet with the info that we've obtaint by executing
        ''' the SQL statement that we've passed as a parameter.
        ''' </summary>
        ''' <param name="consulta">SQL statement that will be executed on the DB</param>
        ''' <param name="ADO_CON">Connection to the DB</param>
        ''' <remarks></remarks>
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

        ''' <summary>
        ''' <see cref="RecordSet.Open"></see>
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Open(ByVal consulta As String, ByVal ADO_CON As ADODBNET.Connection)
            Open(consulta, ADO_CON, Me.CursorLocation, Me.CursorType, Me.LockType)
        End Sub

        'Function that creates a DataSet executing the SQL Query {consulta} within the DB specified
        'in the {conn} object.
        ''' <summary>
        ''' Internal-use function. It fills the DataSet that the RecordSet needs
        ''' to store the obtained information.
        ''' </summary>
        ''' <param name="consulta">SQL statement that will be executed on the DB</param>
        ''' <param name="ADO_CON">Connection to the DB</param>
        ''' <remarks></remarks>
        Public Sub createDataSet(ByVal consulta As String, ByVal ADO_CON As ADODBNET.Connection)
            Me.rsDataSet = New DataSet()
            Me.rsDataAdapter = New SqlDataAdapter(consulta, ADO_CON.conn)
            Me.rsDataAdapter.FillSchema(rsDataSet, SchemaType.Source, "IFIFPC")
            Me.rsDataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Me.rsDataAdapter.Fill(Me.rsDataSet, "IFIFPC")
        End Sub

        ''' <summary>
        ''' Retrieves a field from the current RecordSet.
        ''' </summary>
        ''' <param name="columnName">Column whose field we want to get</param>
        ''' <returns>Field object representing the value at the specified column in the current row</returns>
        ''' <remarks></remarks>
        Public Function Fields(ByVal columnName As String) As Field
            Dim item As Object = Me.rsDataSet.Tables(currentTable).Rows(currentRow).Item(columnName)
            Dim f As New Field(columnName, Me.rsDataSet.Tables(currentTable).Rows(currentRow).Item(columnName).GetType, item, rsDataAdapter, rsDataSet, Me.ADO_CON, currentTable, currentRow, currentColumn)
            Return f
        End Function

        ''' <summary>
        ''' Advances to the next row. 
        ''' </summary>
        ''' <remarks></remarks>
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
