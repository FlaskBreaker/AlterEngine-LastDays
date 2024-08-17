![Banner](banners/AELD0/AELD0-Firma.png)

# AlterEngine-LastDays

## Cliente

### Controles

* `Control / Ctrl`
* `Up / Arriba`
* `Down / Abajo`
* `Left / Izquierda`
* `Right / Derecha`
* `Shift` Correr
* `Return` Escribir en el chat
* **`F1`** Abrir el menú de administración en el cliente (solo usuarios con privilegios).

## Ejecutar

### Linux

Es posible ejecutar Alter Engine (cliente, servidor y editor) en GNU/Linux mediante el uso de [Wine](https://www.winehq.org/).

Requisitos previos:

* [`Wine`](https://www.winehq.org/)
* `make`
* `git`

Clonar el repositorio:

```shell
git clone https://github.com/FlaskBreaker/AlterEngine-LastDays.git
```

Crear nuevo entorno de Wine, instalar las libs y ejecutar el servidor.

```shell
cd  AlterEngine-LastDays
make -C linux
```

> [!IMPORTANT]
> A veces, cuando se inicia el servidor es posible que aparezca un error. Para resolverlo, vuelve a iniciarlo.

#### Otros comandos de `make`

* `make -C linux run`. Ejecuta el servidor.
* `make -C linux run-server`. Ejecuta el servidor.
* `make -C linux run-client`. Para ejecutar el cliente.
* `make -C linux run-editor`. Ejecuta el editor de scripts.
* `make -C linux clean`. Elimina el entorno de Wine.
* `make -C linux install-libs`. Instala las dependencias de Alter Engine.

## Colaboradores

* FlaskBreaker
* Mapacho
* Kyplim
