# Tarjeta de fidelización digital (Club DMD)

Servicio de Dany Marketing Digital: tarjeta de visitas virtual para los clientes de un local. **No usa QR**: en caja se busca al cliente por **teléfono** o por su **código de socio** (4 letras) y se toca "Sumar visita".

## Piezas

| Archivo | Qué es |
|---|---|
| `index.html` | Tarjeta del cliente: crear tarjeta (nombre, teléfono, correo), iniciar sesión, ver visitas, premios y novedades |
| `caja.html` | Panel del local, protegido con PIN: buscar cliente, sumar visita, canjear premios |
| `api.js` | Conexión con el backend. Si `API_URL` está vacío funciona en **modo demo** (PIN 1234) |
| `styles.css` | Estilos (los colores se toman del Sheet de cada local) |
| `apps-script/Code.gs` | Backend en Google Apps Script (uno solo para todos los locales) |

## Cómo funciona

1. El cliente entra a `club.danymarketing.cl/?l=<local>` y crea su tarjeta. Sus datos quedan en el Sheet del local (pestaña **Clientes**), que también sirve como base de datos de marketing.
2. En cada visita el cajero abre `caja.html?l=<local>` (queda con sesión iniciada), escribe el teléfono o código del cliente y toca **Sumar visita**. Solo se suma 1 visita por día, salvo que el cajero confirme "sumar igual".
3. Al llegar a una casilla con premio se crea un **cupón** con código (vigente 30 días). El cliente lo ve en su tarjeta y el cajero lo marca como canjeado.
4. Al completar la tarjeta, la próxima visita empieza una tarjeta nueva y queda registrado cuántas completó.

Premios de Mundo RDL: visita 3 🥤 jugo o café · visita 6 🎁 regalo sorpresa · visita 8 🛝 zona de juegos gratis · visita 10 ⭐ regalo gratis.

## Instalación (una sola vez)

1. En Drive (carpeta MENU DMD o una nueva "CLUB DMD") crea un Google Sheet llamado **CLUB DMD - Central**.
2. Menú **Extensiones → Apps Script**. Borra lo que trae, pega `apps-script/Code.gs` y guarda.
3. Vuelve al Sheet y recárgalo. Aparece el menú **Club DMD** → **Preparar Sheet central** (pedirá autorizar tu cuenta: aceptar).
4. En Apps Script: **Implementar → Nueva implementación → Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier persona**
   - Copia la URL que termina en `/exec`.
5. Pega esa URL en `api.js` (`const API_URL = '...'`) y sube los archivos a GitHub Pages (repo nuevo `club`, dominio `club.danymarketing.cl`, igual que la carta).

> Si cambias `Code.gs` después, usa **Implementar → Administrar implementaciones → editar → Nueva versión** para mantener la misma URL.

## Dar de alta un local

1. En el Sheet central: **Club DMD → Crear nuevo local…** (nombre y slug).
2. Se crea el Sheet "Club <Local>" con las pestañas listas y se muestra el **PIN de caja**.
3. Edita en ese Sheet: **Config** (nombre, logo, colores, descripción, Instagram), **Premios** y **Novedades**.
4. Entrega al local:
   - Link para sus clientes: `https://club.danymarketing.cl/?l=<slug>` (se puede crear la carpeta `<slug>/` con redirección, como en la carta).
   - Link de caja: `https://club.danymarketing.cl/caja.html?l=<slug>` + PIN.
   - Opcional: compartir su Sheet como lector o editor para que vea su base de clientes.

## Probar la demo

Abre `index.html?l=demo` y `caja.html?l=demo` en el mismo navegador (PIN `1234`). Los datos quedan solo en ese navegador.
