# Visual Basic
Ejemplo de integración al webservice con Visual Basic

Se deberá hacer uso de las URL que hacen referencia al WSDL, en cada petición realizada:

- [Timbox Pruebas](https://staging.ws.timbox.com.mx/timbrado_cfdi33/wsdl)

- [Timbox Producción](https://sistema.timbox.com.mx/timbrado_cfdi33/wsdl)

Para hacer el POST con el envelope construido, se usa la URL:

- [POST](https://staging.ws.timbox.com.mx/timbrado_cfdi33/action)


En la clase cServicios.cls estan los ejemplos para construir la petión de timbrado asi como la de cancelación.

### Generacion de Sello
Para generar el sello hay dos opciones: 1) Utilizar el archivo pfx(*.pfx) con su password o 2) Utilizar el certificado(*.cer) y la llave privada (*.key) en formato PEM. 

Sin embargo termina haciendo los mismos procedimientos ya que un pfx es la combinacion del certificado y su llave privada. Entonces en este ejemplo se extrae la llave privada y el certificado para poder generar el sello. También es necesario incluir el XSLT del SAT para obtener transformar el XML a la cadena original.

De la cadena original se obtiene el digest y luego se utiliza el digest y la llave privada para obtener el sello. Finalmente el sello es actualizado en el archivo XML para que pueda ser timbrado.
Todo esto se realiza con librerias de encriptacion de .NET.
Para el funcionamiento del ejemplo se deben importar varias librerias:
```
Imports MSXML2
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
```
Se utiliza la libreria MSXML para poder formar las peticiones al servicio SOAP. Se [descarga]() y se hace la referencia en el proyecto.
Para 

