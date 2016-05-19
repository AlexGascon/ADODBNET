# ADODBNET
Library that simplifies the VB6 to VB.NET migration process of the projects that use ADODB

# Clase ADODBNET: portando ADODB a VB.NET

# INTRODUCCIÓN

Con el paso de Visual Basic 6 a Visual Basic .NET, son muchas las librerías o los componentes que dejaron de funcionar o que tuvieron que cambiar radicalmente. Entre ellas, se encuentra ADO (ActiveX Data Objects), un conjunto de componentes utilizado para el acceso a bases de datos. Mientras que en VB6 podíamos realizar prácticamente todas las operaciones permitidas en la DB utilizando únicamente sus componentes _Connection, RecordSet_ y _Command_, en VB.NET el procedimiento ha cambiado completamente, hasta el punto de que estas tres clases ni tan siquiera existen.

Este gran cambio implicaría tener que reescribir desde cero todo el código de aquellos programas en los que usemos ADO, lo que sin duda puede implicar un trabajo demasiado costoso. No obstante, precisamente eso es lo que pretendemos evitar con la clase que hemos diseñado: ADODBNET

# ADODBNET

ADODBNET es una clase cuya única función es simplificar al máximo la migración entre ADO y ADO.NET. Para ello, dispone de clases y comandos cuyas llamadas son exactamente iguales a las que había disponibles en ADO (RecordSet, Field, Open, MoveNext…), pero que internamente funcionan utilizando ADO.NET. De esta forma, bastará con sustituir la declaración del objeto de tipo ADO que utilicemos en nuestro programa por su equivalente en ADODBNET, y nuestra clase se encargará de que todo el resto de nuestro código funcione perfectamente.

Lamentablemente la conversión no es perfecta, por lo que es posible que haya problemas de optimización temporal o de recursos, pero las funcionalidades básicas están cubiertas con exactitud permitiéndonos ejecutar nuestro programa sin mayores problemas.

# ESTRUCTURA DEL CÓDIGO

La finalidad principal de esta clase es, como ya hemos mencionado, que se asemeje lo máximo posible a los componentes originales de ADO. Por ello, también su estructura se ha realizado pensando de esta manera.

Tenemos una clase principal, ADODBNET, que es el equivalente a lo que en ADO era la clase ADODB. Por ello, no tiene ninguna funcionalidad directa, sino que sirve como contenedor de las enumeraciones y clases que utilizaremos para realizar acciones e interactuar con nuestra base de datos. Las enumeraciones incluidas son CursorType, CursorLocation y LockType, mientras que las clases disponibles son RecordSet, Connection, Command y Field; tanto las unas como las otras también para emular el comportamiento de sus homónimos de ADO. Si bien faltan algunas de las clases disponibles en el ADO original, sí están todas las que necesita nuestro proyecto. Además, el código está escrito y organizado de forma que se pueda extender sin necesidad de alterar ninguna de las clases ya diseñadas, por lo que en caso de que en un futuro necesitásemos otros componentes podríamos añadirlos sin arriesgarnos a inutilizar las mismas.

Vamos ahora a proceder a explicar en mayor detalle el funcionamiento de cada clase y sus campos y métodos.



## Clase Connection

Se trata del componente que se encargará de realizar la conexión con la base de datos. Dispone de todos los campos y métodos necesarios para ello, así como de otros opcionales que permiten definir con mayor nivel de precisión el comportamiento que tendrá nuestra conexión.

### Campos y propiedades:

- **Public**  **ConnectionString**  **As**  **String**: String en la que incluiremos la dirección de la base de datos a la que vayamos a conectarnos.
- **Public**  **CommandTimeOut**  **As**  **Integer**: Número entero que indica la cantidad de segundos que estaremos esperando a una respuesta por parte de la base de datos cuando ejecutemos un comando sobre la misma antes de abortar. El valor por defecto es 30.
- **Public**  **CursorLocation**  **As**  **CursorLocationEnum**: Sirve para seleccionar la ubicación (cliente o servidor) del cursor utilizado para recorrer la base de datos a la que nos conectemos.
- **Protected**  **sqlconn**  **As New**  **SqlConnection**(): Objeto que utilizaremos para la conexión a la base de datos. Está configurado como _Protected_ porque al no ser un componente incluido en la clase Connection de ADO nuestro programa no deberí­a poder acceder a él directamente. No obstante, es probable que si necesitemos acceso a él desde otras clases de ADODBNET.
- **Public ReadOnly Property**  **conn**() As **SqlConnection**: Propiedad que se utiliza para acceder al objeto sqlconn desde cualquier clase (por si en algún momento lo necesitásemos). No obstante, no permite cambiar su valor.

### Constructores:

- **Public Sub**  **New** (ByVal **connString** As String, ByVal **commTimeout** As Integer, ByVal **cursLoc** As CursorLocationEnum): Constructor que fija el valor de todos los campos del objeto _ADODBNET.Connection_ (excepto de sqlconn, pero porque lo definirá a partir de estos)
- **Public Sub**  **New** (ByVal **connString** As String): Constructor que nos permite especificar únicamente a qué base de datos queremos conectarnos. Los parámetros _CommandTimeout_ y _CursorLocation_ toman respectivamente los valores 15 y _adUseClient_, que son los que tomaban por defecto en ADO.

### Métodos:

- **Public Sub**  **Open**(): Método que abre la conexión con la base de datos. Utiliza para ello los parámetros del objeto _ADODBNET.Connection_ desde el que llamamos al método.
- **Public Sub**  **Open**(ByVal **connString** As String): Actualiza el campo _ConnectionString_ del objeto con la cadena especificada en el parámetro _connString_ y abre la conexión.



## Clase Command

La clase _ADODBNET.Command_ permite crear objetos que ejecuten por sí mismos comandos sobre una base de datos, a partir de una conexión a que nosotros mismos le especifiquemos. No obstante, únicamente permite ejecutar comandos que no esperen ningún resultado de vuelta (como INSERT, UPDATE, DELETE…)

### Campos y propiedades:

- **Public**  **ActiveConnection**  **As**  **ADODBNET.Connection**: Conexión a la base de datos en la que se ejecutará el comando especificado en el objeto. 
- **Public**  **CommandText**  **As**  **String**: String que contiene el comando a ejecutar en la base de datos.
- **Public**  **CommandTimeOut**  **As**  **Integer**: Número entero que indica la cantidad de segundos que estaremos esperando a una respuesta por parte de la base de datos cuando ejecutemos un comando sobre la misma antes de abortar. El valor por defecto es 30.

### Constructores:

 No se han diseñado constructores para la clase _ADODBNET.Command_. Dado que todos sus campos son _Public,_ si deseamos cambiar u obtener el valor de alguno de ellos, podemos acceder a ellos directamente desde el objeto que hayamos creado.

 Se ha optado por no diseñar ningún constructor simplemente por el hecho de que desde nuestro programa no se realizaba ninguna llamada que los utilizase. No obstante, en caso de que en un futuro se detectase que estos pueden ser de utilidad, no habría más que escribirlos y añadirlos al código ya existente.

 _(NOTA: Hay que tener en cuenta que, al no haber diseñado ningún constructor, para crear el objeto habrá que utilizar el constructor por defecto que VB.NET genera automáticamente con todas las clases desarrolladas, y que no recibe ningún parámetro y deja todos los campos internos del objeto a los valores por defecto de dichos tipos)_

### Métodos:

- **Public Sub**  **Execute**(): Método que realiza en la base de datos la conexión a la cual hemos definido mediante el campo ActiveConnection las operaciones dadas por la instrucción SQL que contenga el campo CommandText.

**IMPORTANTE**:** únicamente se pueden ejecutar sentencias que no devuelvan ningún resultado (INSERT, DELETE, UPDATE…). Esto se debe a que en ADO.NET el comando "Execute" que tenían los objetos ADODB.Command ha quedado separado en "ExecuteReader", que nos devuelve el resultado almacenado en un DataReader; "ExecuteScalar", que únicamente nos devuelve el primer resultado [primera fila, primera columna] de todos los obtenidos con la consulta; y "ExecuteNonQuery", que ejecuta todas aquellas instrucciones que realizan operaciones sobre la DB sin esperar ningún resultado de vuelta. En nuestro caso utilizamos ExecuteNonQuery(), dado que todas las veces que llamábamos a Execute era para ejecutar sentencias de este tipo; no obstante, si se desease ejecutar cualquier otro tipo de instrucción habría que cambiar antes el código de esta función.**

## Clase RecordSet

Probablemente, la clase más importante de todas las que hemos desarrollado. Los objetos de esta clase se utilizan para almacenar los resultados de una determinada consulta a la base de datos en una tabla para que podamos navegar por ellos, editarlos o introducir nuevos.

### Campos y propiedades:

- **Public**  **rsDataSet**  **As**  **DataSet**: Objeto en el que almacenaremos toda la información que obtengamos de la base de datos. Puede estar compuesto por varias tablas.
- **Public**  **rsDataAdapter**  **As**  **SqlDataAdapter**: Objeto que hará de puente entre el DataSet y la base de datos a la que nos queremos conectar. Más concretamente, se encarga de ejecutar la consulta en la DB para después volcar la información en el DataSet.
- **Public**  **currentTable**, **currentRow**,  **currentColumn**  **As**  **Integer**: Variables que utilizamos para saber en qué posición del DataSet nos encontramos, dado que es posible que queramos recorrerlo de forma secuencial o que en una determinada función necesitemos acceder a un valor concreto y no sepamos en qué posición estamos. Esta necesidad viene derivada de que el antiguo objeto ADODB.RecordSet estaba diseñado para recorrerse secuencialmente, a diferencia del DataSet que se recorre como si fuese una matriz. Por ello, él mismo se encargaba de recordar internamente en qué posición estaba nuestro cursor, lo que hací­a que el usuario no indicase explí­citamente la posición en la que realizar una operación y por tanto tengamos que ser nosotros los que controlemos esto.
- **Public**  **CursorLocation**  **As**  **CursorLocationEnum**: Variable que sirve para definir la localización en la que se encuentra el cursor en nuestro RecordSet. Sus posibles valores vienen dados por la enumeración _CursorLocationEnum_, definida en la clase ADODBNET y que veremos con detalle más adelante.
- **Public**  **CursorType**  **As**  **CursorTypeEnum**: Variable que sirve para definir el tipo de cursor que usamos en el objeto RecordSet en cuestión. Sus posibles valores vienen dados por la enumeración _CursorTypeEnum_, definida en la clase ADODBNET y que veremos con detalle más adelante.
- **Public**  **LockType**  **As**  **LockTypeEnum**: Variable que sirve para definir el tipo de cerrojo que utiliza nuestro RecordSet. Sus posibles valores vienen dados por la enumeración _LockTypeEnum_, definida en la clase ADODBNET y que veremos con detalle más adelante.
- **Public ReadOnly Property**  **FieldsCount**() As **Integer**: Propiedad que podemos utilizar para saber cuántas columnas tiene la fila en la que nos encontramos. Es únicamente de lectura.
- **Public ReadOnly Property**  **EOF**() As **Boolean**: Propiedad que nos indica si ya hemos llegado al final de la tabla o si por el contrario siguen quedando filas que examinar. En caso de que hayamos llegado al final, devuelve _True_; en caso contrario, devuelve _False_.

### Constructores:

- **Public Sub**  **New**(): Único constructor disponible, dado que habitualmente el lugar desde el que configuramos el RecordSet es desde el método Open. Crea un RecordSet con _CursorLocation = adUseServer_, _CursorType = adOpenForwardOnly_, _LockType = adLockReadOnly._

### Métodos:

- **Public Sub**  **Open**(ByVal **consulta** As String, ByVal **ADO\_CON** As ADODBNET.Connection, ByVal **cursLocation** As CursorLocationEnum, ByVal **cursType** As CursorTypeEnum, ByVal **lockType** As LockTypeEnum): Método que abre el RecordSet y carga en él la información que se obtiene al ejecutar la instrucción SQL _consulta_ en la base de datos con la que conectamos mediante _ADO\_CON_. El resto de parámetros definen los campos de las enumeraciones correspondientes, y sobreescriben los valores que puedan haber venido dados por otros objetos. _(NOTA: Se encarga de abrir la conexión de ADO\_CON)_
- **Public Sub**  **Open**(ByVal **consulta** As String, ByVal **ADO\_CON** As ADODBNET.Connection):** Sobrecarga del método anterior que únicamente varí­a en que recibe como parámetros la instrucción SQL a ejecutar y la conexión sobre la cuál hacerlo. Para el resto de parámetros, toma los valores que pueda haber fijados ya.
- **Public Sub**  **createDataSet**(ByVa  **consulta**  **As String, ByVal**  **ADO\_CON** As **ADODBNET.Connection**): Método de uso interno que usamos para crear y llenar el DataSet. Es llamado por el método _Open._
- **Public Sub**  **Fields**(ByVal **columnName** As String) As **Field**: Crea y devuelve un objeto de tipo _Field_ que representa el dato que hay en la columna _columnName_ de la fila _currentRow_ de la tabla _currentTable_. 
- **Public Sub**  **MoveNext**(): Método que aumenta en 1 el valor de _currentRow_, para representar que hemos avanzado una fila.

## Clase Field

Clase que usaremos para representar los objetos contenidos en las celdas de la base de datos. Nos permite tanto leer sus datos como leer y fijar su valor.

### Campos y propiedades:

- **Public**  **Name**  **As**  **String**: Nombre del campo al que pertenece este objeto
- **Public**  **Type**  **As**  **Type**: Tipo del campo en cuestión
- **Public**  **DefinedSize**  **As**  **Integer**: Tamaí±o que tienen por defecto los objetos de tipo _Type_. _OJO: A diferencia de lo que ocurrí­a en ADO, aquí DefinedSize no devuelve el tamaño del objeto en cuestión, sino el que tienen por defecto los objetos de dicho tipo. Esto no supone ningún problema en tipos básicos como Integer, Char, Boolean... , pero sí­ puede suponerlo en tipos de tamaño variable como String u Object o en Arrays._
- **Public**  **rsDataSet**  **As**  **DataSet**: Referencia al DataSet que estábamos recorriendo. De esta forma, podemos hacer consultas y modificar los valores correspondientes al objeto field en que nos encontramos.
- **Public**  **currentTable, currentRow, currentColumn**  **As**  **Integer**: Referencias a los valores equivalentes del DataSet que estábamos recorriendo. De esta forma, en caso de que tengamos que cambiar los valores del Field, podremos saber en qué posición nos encontrábamos y aplicar el cambio a la DB.
- **Public Property**  **Value**() As ** Integer **: Propiedad que nos permite leer y fijar directamente el valor del field. Actúa directamente sobre el DataSet, que se encarga cuando sobreescribimos su valor de guardar el cambio también en la base de datos.

### Constructores:

- **Public Sub**  **New**(ByVal **fieldName** As String, ByVal **fieldType** As Type, ByVal **value** As Object, ByRef **ds** As DataSet, ByRef **tablePosition** As Integer, ByRef **rowPosition** As Integer, ByRef **colPosition** As Integer):** Constructor mediante el que asignamos un valor o una referencia a todos los campos del objeto (excepto a _DefinedSize_ que obtiene su valor de _Value_).

### Métodos:

La clase Field no tiene ningún método.

### Actualizando nuestro código

En esta sección, detallaremos los pasos a seguir para que nuestro código quede correctamente actualizado, y no notemos si quiera la migración de ADO a ADODBNET. Bastará con realizar los reemplazos que aquí se detallan para que podamos olvidarnos completamente de todo lo relacionado con el cambio.

| CÓDIGO ORIGINAL | CÓDIGO DE ADODBNET | _Comentarios_ |
| --- | --- | --- |
| ADODB.{Whatever} | ADODBNET.{Whatever} | Cambiamos la clase principal |

