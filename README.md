# Scrap-Contractorfind
Permite extraer datos de Contractorfind usando Python

## Información
Este script permite extraer datos de las urls de contractorfind y las exporta a un Excel.<br>
Los datos que extrae son:<br>
* URL
* Name
* Contact Person
* Tel
* Cel
* Physical Address
* Company Name
* More Additions & Alterations Contractors in West Rand

Y las URLs están almacenadas en un archivo txt de nombre<br>
<code>urls.txt</code><br><br>

Para automatizar los tiempos de extracción de los datos, se ha utilizado hilos (Threads) para hacer varias peticiones al mismo tiempo de distintas urls.<br>
Con esto logramos extraer más de 20000 datos de 864 ulrs en aproximadamente 3 minutos. Tiempo que no conseguiríamos si no usáramos hilos.

## Autor
Marco Weihmüller

## Licencia
Este proyecto está bajo la Licencia GNU General Public License v3.0
