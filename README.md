# ComparadorLCM
Código de automatización y asistencia del proceso de comparación de bloques utilizando dispositivos Tesa, Fluke y Vaisala.



## Detalles y notas

Para corregir los permisos del gpiomem se deben de automatizar los comandos al inicio:
``` sudo chown root:gpio /dev/gpiomem ```
``` sudo chmod g+rw /dev/gpiomem ```

### Dispositivos Seriales utilizados
Definidos en el archivo `/etc/udev/rules.d/10-usb-serial.rules`

* TESA: Se utiliza la dirección personalizada "/dev/ttyUSBI" definida por los atributos 
```  SUBSYSTEM=="tty", ATTRS{idProduct}=="2303", ATTRS{idVendor}=="067b", ATTRS{version}==" 2.00", SYMLINK+="ttyUSBI"  ```


* Fluke: "/dev/ttyUSBK" definida por los atributos 
```  SUBSYSTEM=="tty", ATTRS{idProduct}=="7523", ATTRS{idVendor}=="1a86", ATTRS{version}==" 1.10", SYMLINK+="ttyUSBK"  ```

* Vaisala: "/dev/ttyUSBD" definida por los atributos 
```  SUBSYSTEM=="tty", ATTRS{idProduct}=="2303", ATTRS{idVendor}=="067b", ATTRS{version}==" 1.10", SYMLINK+="ttyUSBD"  ```
