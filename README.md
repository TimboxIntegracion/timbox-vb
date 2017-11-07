# Visual Basic
Ejemplo de integración al webservice con Visual Basic

Se deberá hacer uso de las URL que hacen referencia al WSDL, en cada petición realizada:

- [Timbox Pruebas](https://staging.ws.timbox.com.mx/timbrado_cfdi33/wsdl)

- [Timbox Producción](https://sistema.timbox.com.mx/timbrado_cfdi33/wsdl)

Para hacer el POST con el envelope construido, se usa la URL:

- [POST](https://staging.ws.timbox.com.mx/timbrado_cfdi33/action)


En la clase cServicios.cls estan los ejemplos para construir la petión de timbrado asi como la de cancelación.

Se debe tener la libreria MSXML instalada para poder hacer la peticion con este ejemplo:
- [MSXML](https://www.microsoft.com/en-US/download/details.aspx?id=3988)

Una vez instalada, se debe agregar a las referencias del proyecto
![referencias](https://imgur.com/TrSRCIK.png)
