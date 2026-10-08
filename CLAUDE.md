# Cuentas de Vale

Presupuesto personal en COP y USD, como app web instalable en el celular.

- **App:** `index.html` (todo el HTML, CSS y JS en un archivo, sin build). Se publica con GitHub Pages en https://cuentas.smarttics.org (`CNAME`).
- **Datos:** la hoja existente «Presupuesto Mensual Vale 2026» (Drive personal de Vale). No se le cambia la estructura: la app lee/escribe `Transactions Log` (títulos en fila 4, datos desde fila 5, columnas A–H con los mismos valores que usan sus fórmulas y desplegables) y lee las categorías de `Setup`. Solo se agregaron columnas I (ID, oculta), J (REEMBOLSO DE), K (EXTRAORDINARIO), L (PRÉSTAMO AHORROS: «Préstamo[: fondo]» en filas Ingreso/Gasto de Ahorros, «Devolución» en filas Ahorros). Le debes a tus ahorros = préstamos − devoluciones, por moneda. Ajustes y saldos reales van en Script Properties, no en la hoja. `apps-script/Codigo.gs` es la API (pegado a mano en Extensiones → Apps Script de esa hoja, publicado como aplicación web). Al editar sin cambiar el monto no se reescribe la columna D (hay montos con fórmula).
- Retiros de ahorro se escriben como en la hoja: Ingreso / «Gasto de Ahorros»; la app los cuenta como ahorro negativo. Reembolsos: Ingreso con «REEMBOLSO DE» lleno.
- **Conexión:** la URL `/exec` y la clave viven solo en `localStorage` de cada dispositivo (`cv.conexion`). Nunca deben ir al repositorio, que es público. Los respaldos `.json` con datos tampoco.
- **Lógica:** ver el bloque `/* ---------- lógica ---------- */` en `index.html`. Balance = ingreso − fijos − variables − ahorros − préstamos. «Pago a TC» no suma en fijos; la deuda de la TC se calcula (compras con TC − pagos), no se resta del balance. Reembolsos (`categoria: Reembolso`, `reembolsaA/reembolsaTipo`) restan del gasto indicado. Retiro de ahorros = ahorro negativo. «Gasto de Ahorros» como ingreso es un formato viejo que se lee como retiro.

## Probar y publicar

- Probar en local: `python3 -m http.server 8080` en esta carpeta y abrir http://localhost:8080 (se conecta a la misma hoja real).
- Si cambia `Codigo.gs`: pegarlo de nuevo en Apps Script y hacer Implementar → Gestionar implementaciones → editar → Nueva versión (la URL no cambia).
- Publicar la app: `git push` a `main`; GitHub Pages la actualiza en ~1 minuto.
