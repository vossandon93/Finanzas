# Instalación paso a paso

Son unos 20 minutos. Al final tendrás la app en tu iPhone, con Face ID, conectada a una planilla tuya en Google.

```
iPhone (app web)  ──►  Google Apps Script  ──►  Google Sheets "Finanzas"
GitHub Pages            (intermediario)          (Movimientos, Presupuesto…)
                              ▲
              Drive "Finanzas – importar"  ◄──  Claude (cartolas convertidas)
```

## 1. Crear la planilla y pegar el script

1. Entra a [sheets.new](https://sheets.new) con tu cuenta de Google y ponle de nombre **Finanzas**.
2. **Extensiones → Apps Script**.
3. Borra todo lo que aparece en `Código.gs` y pega el contenido completo de [`apps-script/Code.gs`](apps-script/Code.gs) (en GitHub: botón **Copiar**).
4. **💾 Guardar**.

## 2. Propiedades del script

En Apps Script: **⚙️ Configuración del proyecto → Propiedades del script → Agregar propiedad del script**.

| Propiedad | Valor | Para qué |
|---|---|---|
| `PIN` | un número de 4 dígitos | Para entrar mientras configuras Google (paso 7). Después deja de usarse. |
| `CASA_SHEET_ID` | el ID de la planilla de Casa App | Para usar los mismos centros de costo y cuentas. El ID es lo que va entre `/d/` y `/edit` en la dirección de esa planilla. |

**Guardar propiedades del script**.

## 3. Ejecutar setup

1. Vuelve al editor (`< >` en el menú izquierdo).
2. Arriba, en el selector de funciones, elige **setup** → **▶ Ejecutar**.
3. Google pedirá permisos: **Revisar permisos → tu cuenta → Configuración avanzada → Ir a … (no seguro) → Permitir**. Es normal: el script es tuyo. Pide permiso para la planilla, para Drive (la carpeta de importación) y para correr solo cada hora.
4. Abajo, en el **Registro de ejecución**, aparece el link a la carpeta **Finanzas – importar** de tu Drive. Ahí se dejan las cartolas convertidas.
5. Revisa la planilla: deben aparecer las hojas **Movimientos, Tarjetas, Presupuesto, Reglas, Centros, Cuentas e Importaciones**. En **Centros** deberían estar los de Casa App.

## 4. Publicar el script como aplicación web

1. **Implementar → Nueva implementación**.
2. En el engranaje de "Seleccionar tipo", elige **Aplicación web**.
3. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier usuario**.
4. **Implementar** y copia la **URL de la aplicación web** (termina en `/exec`).
5. **Pásale esa URL a Claude**: la pone en `config.js` y la deja en `main`.

"Cualquier usuario" es necesario para que la app se conecte sin pedir login de Google en cada llamada. Lo que protege tus datos es el PIN y, después, el ingreso con Google (paso 7).

## 5. Activar GitHub Pages

1. En el repo `Finanzas`: **Settings → General → Danger Zone → Change visibility → Public**. Solo queda público el código; tus movimientos están en tu Google Drive y las cartolas nunca se suben a GitHub.
2. **Settings → Pages → Build and deployment**: Source **Deploy from a branch**, Branch **main** y carpeta **/ (root)** → **Save**.
3. En 1–2 minutos la app queda en: **https://vossandon93.github.io/Finanzas/**

## 6. Instalar en el iPhone

1. Abre el link en Safari y entra con el PIN.
2. Botón **Compartir → Agregar a inicio**.

## 7. Entrar con tu cuenta de Google (y después Face ID)

Puedes usar el mismo cliente de Google que ya creaste para Casa App:

1. [console.cloud.google.com](https://console.cloud.google.com) → el proyecto de Casa App → **APIs y servicios → Credenciales** → tu cliente **Aplicación web**.
2. En **URI de redireccionamiento autorizados** agrega `https://vossandon93.github.io/Finanzas/` (con la barra final) → **Guardar**. El origen `https://vossandon93.github.io` ya está.
3. Copia el **ID de cliente**.
4. En Apps Script → **⚙️ Propiedades del script**, agrega:
   - `GOOGLE_CLIENT_ID` = el ID de cliente
   - `CORREOS_PERMITIDOS` = `vossandon93@gmail.com`
5. En el editor elige la función **activarAccesoGoogle** → **▶ Ejecutar** → acepta el permiso. El registro debe decir "✓" en las dos líneas.
6. **Implementar → Gestionar implementaciones → ✏️ → Versión: Nueva versión → Implementar.** (Nueva versión, **nunca** Nueva implementación: esa cambia la URL.)
7. En el iPhone cierra la app y ábrela: ahora pide **Entrar con Google**. Una vez dentro, toca **Activar Face ID**.

---

## Uso diario

- **Cargar cartolas:** pásaselas a Claude. Él las convierte y las deja en la carpeta de Drive; el script las carga solo cada hora, o al tiro con **Resumen → Cartolas → Importar ahora**.
- **Categorizar:** pestaña **🔎 Revisar** → toca un movimiento → elige tipo, centro y cuenta. Deja marcado **"Hacer lo mismo con todos los que digan…"** y la app crea una regla: los movimientos iguales (de antes y los que vengan) se categorizan solos.
- **Tipos:** solo **Gasto** e **Ingreso** cuentan en los resúmenes. **Pago tarjeta** (el pago de la tarjeta desde la cuenta corriente) y **Traspaso** (entre tus cuentas) no cuentan, para no sumar doble. **Ignorar** lo saca de todo.
- **Presupuesto:** pestaña **🎯 Presupuesto** → **+ Tope** en cada centro. Para una cuenta, escribe por ejemplo `cuenta Supermercado 300000`.
- **Centros y cuentas:** se cambian en **Casa App** (o su planilla) y aquí se actualizan solos en unos minutos. Para forzarlo, en el editor ejecuta **copiarCentrosDeCasaApp**.
- **Reglas:** están en la hoja **Reglas** de la planilla. Si agregas o corriges una a mano, ejecuta **aplicarReglas** en el editor.

## Si cambias el script más adelante

Pega el nuevo código y ve a **Implementar → Gestionar implementaciones → ✏️ editar → Versión: Nueva versión → Implementar**. Comprueba abriendo la URL `/exec`: debe mostrar la versión nueva.

## Problemas comunes

| Mensaje en la app | Solución |
|---|---|
| Falta configurar la URL de Apps Script en config.js | Paso 4.5: pásale la URL a Claude. |
| PIN incorrecto / El PIN no está configurado | Revisa la propiedad `PIN` (paso 2). |
| Falta ejecutar setup() | Paso 3. |
| Falta actualizar el script de Google | Pega el `Code.gs` nuevo y publica una **Nueva versión**. |
| No se pudieron leer los centros de Casa App | Revisa `CASA_SHEET_ID` (paso 2) y ejecuta **copiarCentrosDeCasaApp**. Mientras, se usa la última copia. |
| Sesión no válida | Vuelve a entrar con Google. |
| No se pudo conectar | Revisa que la implementación tenga acceso "Cualquier usuario" y que la URL termine en `/exec`. |
