# Roadmap `@interpolator/xlsx` v1.0.0

## 1. Corrección de errores de sintaxis y configuración

- [x] ✅ Corregir error de shorthand `{ name: 'Germán' }` → `data: { name: 'Germán' }` en tests.
- [ ] Asegurar que `tsconfig.json` y `vitest.config.ts` estén bien configurados para ESM.
- [x] Probar que `pnpm test` no falle por errores de sintaxis.

## 2. Implementación funcional de `interpolateXlsx`

- [x] ✅ Implementar lógica de lectura con `ExcelJS`.
- [x] Detectar filas con `[[]]` (y extraer nombre del array).
- [x] Verificar que el array exista y sea un array.
- [x] Eliminar la fila original si tiene `[[]]`.
- [x] Insertar N filas nuevas (una por ítem del array).
- [x] Interpolar `[[array.prop]]` en cada nueva fila con el ítem correspondiente.
- [x] Interpolar `{{key}}` con el objeto raíz (fuera del loop).
- [x] Dejar marcadores intactos si la clave no existe.
- [x] Convertir `null`/`undefined` a "".
- [x] Eliminar fila si el array es vacío.

## 3. Preservación de formato y fórmulas

- [x] Copiar estilos de celda original a las nuevas filas (color, borde, font, etc.).
- [ ] Preservar fórmulas y ajustar referencias relativas al clonar filas.
- [ ] Mantener merges de celdas si existen en la fila original.
  - Notas v1:
    - Al clonar filas, se replican manualmente los merges usando `mergeCells`.
    - Sin embargo, en la versión actual de `exceljs`, los merges añadidos dinámicamente no se reflejan de forma fiable tras `writeBuffer` + `load`.
    - Hay un test `should replicate merged cells for each expanded array row` en `functional.test.ts` marcado como `it.skip(...)` que describe el comportamiento deseado.
    - Cuando se retome este feature, revisar primero la API de merges de `exceljs` (y posibles upgrades de versión) antes de implementar una solución definitiva.

## 4. Testing funcional

- [x] Crear archivo XLSX de ejemplo (`template.xlsx`) con:
  - `{{user.name}}`
  - `[[items.id]]`, `[[items.qty]]`
  - Fórmula en otra columna (ej: `=B2*C2`)
- [x] Probar interpolación simple → valor correcto.
- [x] Probar expansión de array → N filas nuevas.
- [x] Probar array vacío → fila eliminada.
- [x] Probar clave faltante → marcador intacto.
- [x] Probar propiedad de ítem faltante → marcador intacto.
- [x] Probar propiedad de ítem `null` → celda vacía.
- [x] Probar fórmula → se ajusta en nuevas filas.

## 5. Manejo de errores

- [ ] Lanzar error si `[[]]` usa clave que no es array.
- [ ] Mensaje de error claro con contexto.

## 6. Empaquetado y distribución

- [ ] Asegurar que `tsup` genere correctamente `dist/` con `.js`, `.cjs`, `.d.ts`.
- [ ] Probar que `pnpm build` funcione sin errores.
- [ ] Verificar que `package.json` tenga `exports`, `files`, `main`, `module` bien definidos.

## 7. Documentación

- [ ] Crear `README.md` en `packages/xlsx` con ejemplo de uso.
- [ ] Incluir en el README los comportamientos clave:
  - `{{}}` vs `[[]]`
  - Claves faltantes → marcador intacto
  - Array vacío → fila eliminada
- [ ] Incluir criterios de aceptación resumidos en el README.

## 8. CI y publicación (opcional para v1)

- [ ] Configurar GitHub Actions (si se desea).
- [ ] Asegurar que `pnpm publish` funcione (cuando esté listo).
