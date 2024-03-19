# intercam-batch

Programa para generar recibos de los socios de la Camara Comercial de Bolivar.

## Comenzando 🚀

_Esto proceso es unicamente necesario cuando clonamos por primera vez el repositorio._

### 1. Compilar Helix:
_Colocar el proyecto en modo "Release". Hacer click derecho sobre Helix en el explorador de soluciones. Esto desplegara un menu de opciones. Luego hacer click sobre "Build". Esperar a que el proceso se complete._

### 2. Compilar TextBoxX:
_Colocar el proyecto en modo "Release". Hacer click derecho sobre TextBoxX en el explorador de soluciones. Esto desplegara un menu de opciones. Luego hacer click sobre "Build". Esperar a que el proceso se complete._

### 3. Actualizar referencias a dependencias en Soccam:
_Hacer clikc derecho sobre SocCam_Mantenimiento. Esto desplegara un menu de opciones. Hacer click en la opcion "Add" o "Agregar". Esto desplegara un menu de opciones. Luego hacer click en "Reference" o "Referencia"._

_Esto abrira en un nuevo recuadro "Reference Manager". Hacer click en el boton "Browse" o "Examinar". Dirigirse a "helix\helix\helix\bin\Release" y seleccionar el archivo "helix.dll". Luego, volver a hacer clikc en el boton "Browse" o "Examinar" y dirigirse a "TextBoxX\TextBoxX\bin\Release" y seleccionar el archivo "TextBoxX.dll"._

_Por ultimo, hacer click en el boton "OK". Esto deberia agregar las referencias a cada una de estas librerias. El proyecto ya esta listo para ser compilado y ejecutado._

### 4. Configurar el sistema en modo dev:
_En el archivo "intercam/intercam/soccam/Helpers/ConfigDatabase.vb" setear la variable "Public Property Production As Boolean = True" en "False" para conectarse a la db local. Para esto es necesario levantar un dump de la base de datos de soccam en su sistema con el nombre "soccam_test". El sistema esta configurado para conectarse a SQLServer con las configuraciones por defecto._

### Ejecutar el proyecto.
_El codigo se ejecuta desde el script "crear_cuotas_sociales.vbs". Por ende, desde CMD dirigirse a la ubicaicon de este archivo en "SocCamMantenimiento\SocCam_Mantenimiento\SocCam_Mantenimiento\bin\Release" y escribir el nombre del script para ejecutarlo. Este cargar las cuotas con los movimientos y facturas para socio. Excluye a los que tienen caja de seguridad.

### Posibles problemas.
_Si al ejecutar el comando "git add ." en nuestro projecto Visual Basic en Visual Studio arroja el error:_
```
"fatal: LF would be replaced by CRLF in SocCam_Mantenimiento/SocCam_Mantenimiento/bin/Release/SocCam_Mantenimiento.exe.manifest"
```
1. Guardar los cambios en Visual Studio y cerrar el IDE.
2. Ejecutar el comando: "git config core.autocrlf false"
3. Ahora ya deberiamos poder ejecutar "git add ."

### Advertencia.
_La branch "generarCuotasPreviasFechaActual" no debe eliminarse ya que contiene el código para generar recibos previos en el sistema de soccam hasta la fecha actual. Originalmente, el sistema no tenía una función para generar recibos previos, ya que estaba diseñado para crearlos en la fecha de ejecución. Esta rama se creó para evitar la necesidad de modificar el código nuevamente en el futuro si se requieren recibos previos en algún momento._
_Para generar recibos previos desde la branch especificada se debe modificar el script "crear_cuotas_sociales.vbs", en "intercam-batch/SocCam_Mantenimiento/SocCam_Mantenimiento/bin/{debug or release}/", y donde se setean las variables de "mes" y "anio", colocar los valores deseados. Luego en el propio codigo, en el archivo "Executor.vb" en "intercam-batch/SocCam_Mantenimiento/SocCam_Mantenimiento/", modificar la linea 307 "c.monto = 1700.0" con el valor correspondiente a la cuota a generar. No olvidarse de establecer el sistema en modo produccion o test, como se especifico mas arriba, acorde a las nesecidades._

## Colaboradores 👥

* **Camila Levato** - [CamilaLev07](https://github.com/CamilaLev07)
* **Emanuel Acosta** - [EmanuelAcosta1695](https://github.com/EmanuelAcosta1695)
* **Alan Medina** - [alangabrielmedina](https://github.com/alangabrielmedina)
