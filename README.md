# Visual Basic
Ejemplo de integración al webservice con Visual Basic

Se deberá hacer uso de las URL que hacen referencia al WSDL, en cada petición realizada:

- [Timbox Pruebas](https://staging.ws.timbox.com.mx/timbrado_cfdi33/wsdl)

- [Timbox Producción](https://sistema.timbox.com.mx/timbrado_cfdi33/wsdl)

Para hacer el POST con el envelope construido, se usa la URL:

- [POST](https://staging.ws.timbox.com.mx/timbrado_cfdi33/action)


En la clase cServicios.cls estan los ejemplos para construir la petión de timbrado asi como la de cancelación.

### Generacion de Sello
Para generar el sello hay dos opciones: 1) Utilizar el archivo pfx(*.pfx) con su password o 2) Utilizar el certificado(*.cer) y la llave privada (*.key) en formato PEM. Se obtiene el mismo resultado ya que un pfx es la combinación del certificado y su llave privada.  

En este ejemplo se utiliza el pfx para poder generar el sello. También es necesario incluir el XSLT del SAT, ya que se utiliza para poder transformar el XML y obtener la cadena original.

De la cadena original se obtiene el digest y luego se utiliza el digest y la llave privada para obtener el sello. Finalmente el sello es actualizado en el archivo XML para que pueda ser timbrado.

Todo esto se realiza con librerias de encriptacion de .NET.
Para el funcionamiento del ejemplo se deben importar varias librerias:
```
Imports MSXML2
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
```

Se debe tener la libreria MSXML instalada para poder hacer la peticion SOAP con este ejemplo:
- [MSXML](https://www.microsoft.com/en-US/download/details.aspx?id=3988)


Una vez instalada, se debe agregar a las referencias del proyecto
![referencias](https://imgur.com/TrSRCIK.png)
