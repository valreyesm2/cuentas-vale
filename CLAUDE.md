# Cuentas de Vale

Presupuesto personal en COP y USD, como app web instalable en el celular.

- **App:** `index.html` (todo el HTML, CSS y JS en un archivo, sin build). Se publica con GitHub Pages en https://cuentas.smarttics.org (`CNAME`).
- **Datos:** una hoja de Google Sheets en el Drive personal de Vale. `apps-script/Codigo.gs` es la API (se pega a mano en Extensiones → Apps Script de esa hoja y se publica como aplicación web). Hojas: `Movimientos`, `Ajustes` (JSON en la fila `config`), `Saldos`.
- **Conexión:** la URL `/exec` y la clave viven solo en `localStorage` de cada dispositivo (`cv.conexion`). Nunca deben ir al repositorio, que es público. Los respaldos `.json` con datos tampoco.
- **Lógica:** ver el bloque `/* ---------- lógica ---------- */` en `index.html`. Balance = ingreso − fijos − variables − ahorros − préstamos. «Pago a TC» no suma en fijos; la deuda de la TC se calcula (compras con TC − pagos), no se resta del balance. Reembolsos (`categoria: Reembolso`, `reembolsaA/reembolsaTipo`) restan del gasto indicado. Retiro de ahorros = ahorro negativo. «Gasto de Ahorros» como ingreso es un formato viejo que se lee como retiro.

## Probar y publicar

- Probar en local: `python3 -m http.server 8080` en esta carpeta y abrir http://localhost:8080 (se conecta a la misma hoja real).
- Si cambia `Codigo.gs`: pegarlo de nuevo en Apps Script y hacer Implementar → Gestionar implementaciones → editar → Nueva versión (la URL no cambia).
- Publicar la app: `git push` a `main`; GitHub Pages la actualiza en ~1 minuto.
