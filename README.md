# EjercicioFinalEXCEL.
### Enunciado
## 
    1 Objetivo
  ### Evidenciar las competencias para usar las características básica del lenguaje VBA para automatizar aplicaciones.  Desarrollar buenas prácticas de programación.
   ## 
    1.1 ¿Qué se Entrega?:
### Se entrega el archivo que contiene las macros siguientes. Se debe mostrar una carpeta en GitHub con los códigos e imágenes de la aplicación, y el archivo xlm.
# 
    2. El juego del Baloto
#### En una lotería en la que el apostador, elije 6 números sin repetición, con valores entre 1 y 43 y una super balota que es un numero entre 1 y 16. La idea es acertar los seis números y la super balota. En el juego real solo se elijen cinco números, pero en esta práctica elegiremos seis.
# 
    3 La aplicación a desarrollar.
### 
    Cuando se abra el archivo de Excel de nombre miPracticaBaloto.xlsm debe aparecer el siguiente formulario:
![image](https://github.com/user-attachments/assets/b8cc0562-188d-4732-a2d6-509c4130e0ce)

####  En este formulario, cuando se carga los comboBox contienen los números del 1 al 43 y el combo de la balota del 1 al 16. Cuando la ventana aparece los combos muestran unos valores sugeridos en cada uno de ellos, que la aplicación define de manera aleatoria. En esta ventana el apostador elije sus los números a los que les va a apostar, eligiendo de cada comboBox, cuando se presiona el botón Jugar, la apuesta se registra en la hoja de Excel, como muestra la siguiente imagen, siempre y cuando no se hayan elegido números repetidos.
![image](https://github.com/user-attachments/assets/eb40f4bb-2779-4c44-9888-aa3ba75db11f)

####   Si en los números elegidos hay números repetidos, no se escriben en la hoja y se muestra un mensaje advirtiendo que se debe volver a jugar. El botón Salir debe cerrar Excel y guardar el archivo automáticamente, sin preguntar si se desea grabar. Cuando se presiona el botón Ingresar se muestra el siguiente formulario:
![image](https://github.com/user-attachments/assets/72590e32-0133-4a01-a571-49eeae984bb4)

#### Al presionar el botón: Obtener Ganador, la aplicación genera los seis numero ganadores sin repetición, y el numero de la Balota ganadora. Al presionar el botón: verificar ganadores, la aplicación busca en la hoja de Excel de apuestas, las apuestas con: 6 acierto y la balota acertada, con 6 aciertos sin balota y 5 aciertos y balota, muestra los ID de la apuesta y la fila en que se encuentran. Los resultados anteriores se deben mostrar en un MsgBox, asi:

![image](https://github.com/user-attachments/assets/1c72ee60-aac8-4d07-b501-af5061a78088)

#### Los datos que se muestran son meramente demostrativos. Si no hay resultados el MsgBox muestra el Mensaje no hubo ganadores correspondiente. El botón regresar hace invisible la ventana Jugar Baloto y vuelve a la ventana inicial.

## 
    4 Criterios de evaluación
#### • Funcionamiento de la aplicación.
#### • Se debe hacer uso de Procedimientos y/o funciones desarrolladas por los desarrolladores.
#### • En el día de la entrega se hará sustentación, el profesor podrá interrogar a los desarrolladores. La no debida sustentación puede afectar la nota, que corresponde al 20% del examen final.
#### • Se deberá enviar el link de GitHub, de donde se podrá descargar la aplicación.
