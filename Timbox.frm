VERSION 5.00
Begin VB.Form Timbox 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtUUID 
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Text            =   "42C7DC8A-9955-451D-99D5-2966AD019985"
      Top             =   1440
      Width           =   7335
   End
   Begin VB.CommandButton BtnTimbrar 
      Caption         =   "Timbrar"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "UUID a Cancelar:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Timbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnCancelar_Click()
    Dim doc As New MSXML2.DOMDocument
    Dim strRFC As String
    Dim pfxBase64 As String
    Dim pfxPassword As String
    
    ' Parametros para cancelar
    strRFC = "AAA010101AAA"
    pfxBase64 = ""
    pfxBase64 = pfxBase64 & "MIIIWQIBAzCCCB8GCSqGSIb3DQEHAaCCCBAEgggMMIIICDCCBQcGCSqGSIb3DQEHBqCCBPgwggT0AgEAMIIE7QYJKoZIhvcNAQcBMBwGCiqGSIb3DQEMAQYw"
    pfxBase64 = pfxBase64 & "DgQIJJ+mrYnkX0UCAggAgIIEwIFwe2P1uJvnGnBZQ6aaNTCiuQK8/RF1EZOX5oicj6Sq2RdKkVEmiXKS/PhHuVpaqxJq3Mackatc1VjfwV63eenDRTYUc3Hz"
    pfxBase64 = pfxBase64 & "JWvNaB9ISDhpm66b+Y/KNQzSjO+giO59jfy8F9Ppks82V+SuLKV9pEWnb8bZGgjr+fiqO/bYRlxGU/P9Q3TTirlnol7RrtgcnEP6Jb06o6f8HmYPZuRuNqgO"
    pfxBase64 = pfxBase64 & "efEnAbH5K2n03DP2wx2PgoBBANzHe6o4mngtckdYVj1IkkgsQNta2lQCCXRa47nqKq+ex/cv2hNx79+mV7DZ+IKXXWNbGidXf2mZrPZOpro3VLE9+UP1WUgr"
    pfxBase64 = pfxBase64 & "nmkzdcq1kSr+wbR3LZW+zFOnjPGOEFKq7MMwdtsRoV++Uf8zKy1q5usiXAxuyhbeuYl9klhZrFJSP3U0uiO9oQaUrRBLAjYEvUtc2V4eHvXcucmQjtF/m9db"
    pfxBase64 = pfxBase64 & "8U6zOn7NlE6yO/ZwEOoqqnDMbPkEh8hpJBUo9Cc+FlhkkVhIOE5gIBktTdQ+vrH4bjwNGlKAdDKYJmShbsiEemL+Q4T0UX5zzbhjKu97cDLCENC++fYcOvay"
    pfxBase64 = pfxBase64 & "yv3PpJbIxgqjg0tniYYNV4JSNi9+HVY+mGGPhN0/qEj0V8d2Di60Q5KWP3uwVCNM/OZTTKTsbnbGDGAlH8qfOQ869+rRS1ZGqPxtKqBpZX2qgLOt7p0fSmyP"
    pfxBase64 = pfxBase64 & "YwYH/P5WGBj2iY2KUf0GatS0Vz1/w6ycVbSACYdvFUhOC0T5X1dGUJRNW6ml56PCgkWQ5b+IiuGBSUCDcRgKmXyM3FN4LOCF9+QFk88Iyt8DhZyFP2Jejize"
    pfxBase64 = pfxBase64 & "hNRq9LoJH8KmLA2YnhS9GBz7CcG2a2sH0L8ob4QQ266e6OIHkqTleTR48fuIqn80OlhRuNGHUtjVtICIHUd/d2ZuhkRiLVNPiUsKqctp/NbFjdBV76pTVb3n"
    pfxBase64 = pfxBase64 & "qgvnkUX20hEmOo0qDN2QRYq/wGMh+fsL4w0blJr/tHP20x94iAYlcNGDwRPFlYQMGeVWpCnaP2nSSwHPBk2h43Q8fx/Ai/LUNUhGSKTH6motZNPhIJgl/M6w"
    pfxBase64 = pfxBase64 & "aISyq+AIPFoOXQx9IVRcrZGWQPbCLihGIYGkKZiJ1wNXwOlLvZgrpaiaeMTgZ+SPeGYiPd95OIH2fPEEEclk7F3EnJceT23OVVLEvdyLSKja7nN63Sm0LFfw"
    pfxBase64 = pfxBase64 & "e6yDUuVXv8+GKeBMQ1N5y6TKq7bWEiWNUJbXWLkps2tqsgtSj3EB8iz3v4Lh8TkG8sA7mKUI5YFoWxz4ptIU+sbtBPwXHg8cN40rnk/EW2wBmrCuR1YU7H1T"
    pfxBase64 = pfxBase64 & "QE6w28XaEcLcylUJ++a4squOMBuiOaYTtWJwwHBdJaOs/GZHc5RAg7cya9YInHeZyTbOxzJmNJrwS80JUHUNG1GhCO120yT/G6tUAQjha5ER0CoFzIxtidQq"
    pfxBase64 = pfxBase64 & "S7/rLEbT+/E7jBB0ZntMVAHgWdXZtvdE/8hGB0DHszvQ7fom/n4epQxlsXV37E+6q3wXE2GvKkZB71iY12lhBZkuARpWzE2G1BwjgRB6OKUxAkze3ZpmURpp"
    pfxBase64 = pfxBase64 & "pJLGrJzaf+22vzIFMRxlRvdV4xFTu/3Oj8eKRnbhmiM/ebwr/h3fQinC+wDFn5JiLFKPg2JOuUcC3G5d7fhnZFowggL5BgkqhkiG9w0BBwGgggLqBIIC5jCC"
    pfxBase64 = pfxBase64 & "AuIwggLeBgsqhkiG9w0BDAoBAqCCAqYwggKiMBwGCiqGSIb3DQEMAQMwDgQIbApT/D/BVKMCAggABIICgEodMGPVDwDycAbmwahnRV2l7vHO69rETjufz0UK"
    pfxBase64 = pfxBase64 & "rYTqDTeZVx6Ur8J/lsWrMXMbXMVt1J06oAimzoBVWlhSQUkT+tYzOLG1aYhOySAjTrF/9mmZ82lEKbokEVBAS7CXBk1Mlzf8D8jEF+GZa0A34VbPeqr55hwJ"
    pfxBase64 = pfxBase64 & "HGRHzlVKcn8VgWdYnxCHIUgOeVoN2tSMGU2s/0l0FVQUpNdtkY+pVkXpXSBN73eKu3IC2Zo6N7TeGVOAasm4Lb5we/gqZxElRrNgO2FcR9r1sO1DTmxbtLgX"
    pfxBase64 = pfxBase64 & "VSqjCEH9aAq7ow5u61+e/1FYQ6nyAWWJ7C+JHFDnPw2VJ2KiPM8mc1TwwtrSIwofKPeV/nUC1kp6Zr6VD0Ju00H3TvvdE9OTA/8r1qIzE3KajFjeqANmiAgt"
    pfxBase64 = pfxBase64 & "ZGYzdlVJYLQKpEpGxgPL3chzwc9chhLCOQBUdP1yHPyNllOn52ogidh0qKDP0keFiowhTYucJ9usFuQLSe/NlIyUV1nk2JAKkMmA0bWgeJ+L96YL9InJySvX"
    pfxBase64 = pfxBase64 & "n8dO6wNYhBlJquV6FtnxCIou+yjCjA892DBItnmKM2xa+xQnI6roScLc9SUrJfx4EUJsu9IvSpX06g4cktc5qymF3BzwDSykLGQ365GEBUIK/fUrJNDHTyK7"
    pfxBase64 = pfxBase64 & "9lPA19MMWKI+sf45kCyAkV7Gvhi4EghnqdwpbqSgxQX/fhHmgsm1hUuo3fVIC01j9rX1gie0LewsjQkJcO7uIax2pScvLgz/5sEBhMGv6Jzr0/GZ+f8X1qJZ"
    pfxBase64 = pfxBase64 & "LJaLqx7rG0/m2bOOCC3fQqzFcxg7ZXA3UQ+Jt2eVHDz15oWoXR59Rr2Tn+if2Z6VjpYrjiK/HfrqcoINpMSe2SIjPFOJTpgxJTAjBgkqhkiG9w0BCRUxFgQU"
    pfxBase64 = pfxBase64 & "n0elLuqWflzq+6wFt5OhOMoDyKIwMTAhMAkGBSsOAwIaBQAEFGcEN6bOyqHAA92f6Ov6gu6ARzEABAjqgtqLbPJ4/QICCAA="
    
    pfxPassword = "12345678a"

    ' Llamar la funcion cancelar
    Dim cRequest   As cServicios
    Set cRequest = New cServicios
    doc.LoadXml cRequest.Cancelar(txtUUID.Text, strRFC, pfxBase64, pfxPassword)
    Debug.Print doc.Text

End Sub

Private Sub BtnTimbrar_Click()
    Dim doc As New MSXML2.DOMDocument
    Dim strXml As String
    Dim encodedXml As String
    strXml1 = FileToString("ejemplo.xml")
    encodedXml = Base64EncodeString(strXml1)
    ' Enviar el XML en formato base64
    

    ' Llamar la funcion timbrar
    Dim cRequest   As cServicios
    Set cRequest = New cServicios
    doc.LoadXml cRequest.Timbrar(encodedXml)
    Debug.Print doc.Text
End Sub

Function FileToString(strFilename As String) As String
  iFile = FreeFile
  Open strFilename For Input As #iFile
    FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
  Close #iFile
End Function

