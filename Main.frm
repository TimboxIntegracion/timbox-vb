VERSION 5.00
Begin VB.Form Main 
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
Attribute VB_Name = "Main"
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
    
    ' Enviar el XML en formato base64
    strXml = ""
    strXml = strXml & "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4KPGNmZGk6Q29tcHJvYmFudGUgeG1sbnM6Y2ZkaT0iaHR0cDovL3d3dy5zYXQuZ29iLm14"
    strXml = strXml & "L2NmZC8zIiB4bWxuczp4c2k9Imh0dHA6Ly93d3cudzMub3JnLzIwMDEvWE1MU2NoZW1hLWluc3RhbmNlIiBMdWdhckV4cGVkaWNpb249Ik3DqXhpY28iIGNl"
    strXml = strXml & "cnRpZmljYWRvPSJNSUlFWVRDQ0EwbWdBd0lCQWdJVU1qQXdNREV3TURBd01EQXlNREF3TURFME1qZ3dEUVlKS29aSWh2Y05BUUVGQlFBd2dnRmNNUm93R0FZ"
    strXml = strXml & "RFZRUUREQkZCTGtNdUlESWdaR1VnY0hKMVpXSmhjekV2TUMwR0ExVUVDZ3dtVTJWeWRtbGphVzhnWkdVZ1FXUnRhVzVwYzNSeVlXTnB3N051SUZSeWFXSjFk"
    strXml = strXml & "R0Z5YVdFeE9EQTJCZ05WQkFzTUwwRmtiV2x1YVhOMGNtRmphY096YmlCa1pTQlRaV2QxY21sa1lXUWdaR1VnYkdFZ1NXNW1iM0p0WVdOcHc3TnVNU2t3SndZ"
    strXml = strXml & "SktvWklodmNOQVFrQkZocGhjMmx6Ym1WMFFIQnlkV1ZpWVhNdWMyRjBMbWR2WWk1dGVERW1NQ1FHQTFVRUNRd2RRWFl1SUVocFpHRnNaMjhnTnpjc0lFTnZi"
    strXml = strXml & "QzRnUjNWbGNuSmxjbTh4RGpBTUJnTlZCQkVNQlRBMk16QXdNUXN3Q1FZRFZRUUdFd0pOV0RFWk1CY0dBMVVFQ0F3UVJHbHpkSEpwZEc4Z1JtVmtaWEpoYkRF"
    strXml = strXml & "U01CQUdBMVVFQnd3SlEyOTViMkZqdzZGdU1UUXdNZ1lKS29aSWh2Y05BUWtDRENWU1pYTndiMjV6WVdKc1pUb2dRWEpoWTJWc2FTQkhZVzVrWVhKaElFSmhk"
    strXml = strXml & "WFJwYzNSaE1CNFhEVEV6TURVd056RTJNREV5T1ZvWERURTNNRFV3TnpFMk1ERXlPVm93Z2RzeEtUQW5CZ05WQkFNVElFRkRRMFZOSUZORlVsWkpRMGxQVXlC"
    strXml = strXml & "RlRWQlNSVk5CVWtsQlRFVlRJRk5ETVNrd0p3WURWUVFwRXlCQlEwTkZUU0JUUlZKV1NVTkpUMU1nUlUxUVVrVlRRVkpKUVV4RlV5QlRRekVwTUNjR0ExVUVD"
    strXml = strXml & "aE1nUVVORFJVMGdVMFZTVmtsRFNVOVRJRVZOVUZKRlUwRlNTVUZNUlZNZ1UwTXhKVEFqQmdOVkJDMFRIRUZCUVRBeE1ERXdNVUZCUVNBdklFaEZSMVEzTmpF"
    strXml = strXml & "d01ETTBVekl4SGpBY0JnTlZCQVVURlNBdklFaEZSMVEzTmpFd01ETk5SRVpPVTFJd09ERVJNQThHQTFVRUN4TUljSEp2WkhWamRHOHdnWjh3RFFZSktvWklo"
    strXml = strXml & "dmNOQVFFQkJRQURnWTBBTUlHSkFvR0JBS1MvYmVVVnk2RTNhT0RhTnVMZDJTM1BYYVFyZTB0R3htWVRlVXhhNTV4MnQvNzkxOXR0Z09wS0Y2aFBGNUt2bFlo"
    strXml = strXml & "NHp0cVFxUDR5RVYrSGpIN3l5LzJkLytlN3QrSjYxalRyYmRMcVQzV0QwK3M1ZkNMNkpPckY0aHF5Ly9FR2RmdllmdGRHUk5yWkgrZEFqV1dtbDJTL2hyTjlh"
    strXml = strXml & "VXhyYVM1cXFPMWI3YnRsQWdNQkFBR2pIVEFiTUF3R0ExVWRFd0VCL3dRQ01BQXdDd1lEVlIwUEJBUURBZ2JBTUEwR0NTcUdTSWIzRFFFQkJRVUFBNElCQVFB"
    strXml = strXml & "Q1BYQVdaWDJEdUtpWlZ2MzVSUzFXRktnVDJ1YlVPOUMrYnlmWmFwVjZaellOT2lBNEttcGtxSFUvYmtaSHFLalIrUjU5aG9ZaFZkbitDbFVJbGlaZjJDaEho"
    strXml = strXml & "OHMwYTB2QlJOSjNJSGZBMWFrV2R6b2NZWkxYanozbTBFcjMxQlkrdVMzcVdVdFBzT05HVkR5Wkw2SVVCQlVsRm9lY1FoUDlBTzM5ZXI4ekliZVUyYjBNTUJK"
    strXml = strXml & "eEN0NHZiREtGdlQ5aTNWMFB1b28ra21ta2YxNUQyckJHUitkcmQ4SDhZZzhUREdGS2YyekttUnNnVDduSWVvdTZXcGZZcDU3MFdJdkxKUVkrZnNNcDMzNEQw"
    strXml = strXml & "NVVwNXlrWVNBeFVHYTMwUmRVekE0cnhONWhUK1c5d2hXVkdEODhURDMzTnc1NXVOUlVjUk8zWlVWSG1kV1JHK0dqaGxmc0QiIGZlY2hhPSIyMDE3LTAzLTIx"
    strXml = strXml & "VDE1OjUxOjMyIiBmb3JtYURlUGFnbz0iUGFnbyBlbiB1bmEgc29sYSBleGhpYmljaW9uIiBtZXRvZG9EZVBhZ289InRhcmpldGEgZGUgY3JlZGl0byBvIGRl"
    strXml = strXml & "Yml0byIgbm9DZXJ0aWZpY2Fkbz0iMjAwMDEwMDAwMDAyMDAwMDE0MjgiIHNlbGxvPSJrQUE5VlRRQUZvMTZIbks5d2l4djlZeE9pNngwSG1UUFlEMURMVkpa"
    strXml = strXml & "NE41QVFOUVRvNG0yL2VSVitOb25yN1pQdFF6VXRMZVcrWlJaSFRDRExqZ05OcTBYYjQwMVFUYWJ5ZDAvdDdGQ0IzTCtKRVprM0c0RjFIbE84endxK1F2Slhw"
    strXml = strXml & "UFYvNjhrQ3BjNGFQb2tHdnBNbElrcnVrcUpxendYQWRtY3lLdzlXMEE9IiBzdWJUb3RhbD0iMS4wMCIgdGlwb0RlQ29tcHJvYmFudGU9ImluZ3Jlc28iIHRv"
    strXml = strXml & "dGFsPSIxLjE2IiB2ZXJzaW9uPSIzLjIiIHhzaTpzY2hlbWFMb2NhdGlvbj0iaHR0cDovL3d3dy5zYXQuZ29iLm14L2NmZC8zIGh0dHA6Ly93d3cuc2F0Lmdv"
    strXml = strXml & "Yi5teC9zaXRpb19pbnRlcm5ldC9jZmQvMy9jZmR2MzIueHNkIj4KICA8Y2ZkaTpFbWlzb3IgcmZjPSJBQUEwMTAxMDFBQUEiPgogICAgPGNmZGk6UmVnaW1l"
    strXml = strXml & "bkZpc2NhbCBSZWdpbWVuPSJSZWdpbWVuIEdlbmVyYWwgZGUgTGV5IFBlcnNvbmFzIE1vcmFsZXMgZGUgUHJ1ZWJhIi8+CiAgPC9jZmRpOkVtaXNvcj4KICA8"
    strXml = strXml & "Y2ZkaTpSZWNlcHRvciByZmM9IkFBRDk5MDgxNEJQNyIvPgogIDxjZmRpOkNvbmNlcHRvcz4KICAgIDxjZmRpOkNvbmNlcHRvIGNhbnRpZGFkPSIxIiBkZXNj"
    strXml = strXml & "cmlwY2lvbj0iT3JvIMOhw6nDrcOzw7ogw7Egw4HDicONw5PDmiDDkSIgaW1wb3J0ZT0iMS4wMCIgdW5pZGFkPSJObyBBcGxpY2EiIHZhbG9yVW5pdGFyaW89"
    strXml = strXml & "IjEuMDAiLz4KICA8L2NmZGk6Q29uY2VwdG9zPgogIDxjZmRpOkltcHVlc3Rvcz4KICAgIDxjZmRpOlRyYXNsYWRvcz4KICAgICAgPGNmZGk6VHJhc2xhZG8g"
    strXml = strXml & "aW1wb3J0ZT0iMC4xNiIgaW1wdWVzdG89IklWQSIgdGFzYT0iMTYuMDAiLz4KICAgIDwvY2ZkaTpUcmFzbGFkb3M+CiAgPC9jZmRpOkltcHVlc3Rvcz4KPC9j"
    strXml = strXml & "ZmRpOkNvbXByb2JhbnRlPgo="

    ' Llamar la funcion timbrar
    Dim cRequest   As cServicios
    Set cRequest = New cServicios
    doc.LoadXml cRequest.Timbrar(strXml)
    Debug.Print doc.Text
End Sub

