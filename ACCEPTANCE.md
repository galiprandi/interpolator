# Criterios de Aceptación: `@interpolator/xlsx`

## 1. Sintaxis de marcadores

### 1.1. Interpolación simple con `{{}}`

- **Dado** una celda con `{{user.name}}`,
- **Y** `data` incluye `{ user: { name: "Germán" } }`,
- **Entonces** la celda debe contener `"Germán"`.

### 1.2. Interpolación con espacios

- **Dado** una celda con `{{ user.name }}`,
- **Y** `data` incluye `{ user: { name: "Germán" } }`,
- **Entonces** debe comportarse igual que `{{user.name}}`.

### 1.3. Interpolación anidada profunda

- **Dado** `{{profile.contact.email}}`,
- **Y** datos: `{ profile: { contact: { email: "g@a.com" } } }`,
- **Entonces** la celda debe contener `"g@a.com"`.

### 1.4. Clave no existe en raíz

- **Dado** `{{user.name}}`,
- **Y** `user` no existe en `data`,
- **Entonces** la celda debe contener literalmente `{{user.name}}`.

### 1.5. Propiedad intermedia no existe

- **Dado** `{{user.profile.email}}`,
- **Y** `user` existe pero no tiene `profile`,
- **Entonces** la celda debe contener `{{user.profile.email}}`.

### 1.6. Valor `null` o `undefined`

- **Dado** `{{user.name}}`,
- **Y** `user.name` es `null` o `undefined`,
- **Entonces** la celda debe quedar vacía (`""`).

---

## 2. Interpolación de arrays con `[[]]`

### 2.1. Array válido no vacío

- **Dado** una fila con `[[payments.id]]` y `[[payments.date]]`,
- **Y** `payments` es: `[{ "id": "P1", "date": "2025-01-01" }, { "id": "P2", "date": "2025-01-02" }]`,
- **Entonces** la fila original debe eliminarse,
- **Y** deben insertarse 2 filas nuevas con los valores correspondientes.

### 2.2. Array vacío

- **Dado** una fila con `[[payments.id]]`,
- **Y** `payments` es `[]`,
- **Entonces** la fila debe eliminarse del documento.

### 2.3. Clave de array no existe

- **Dado** `[[payments.id]]`,
- **Y** `payments` no existe en `data`,
- **Entonces** la celda debe contener `[[payments.id]]`,
- **Y** la fila **no debe eliminarse ni repetirse**.

### 2.4. Clave de array no es un array

- **Dado** `[[user.id]]`,
- **Y** `user` existe pero es un objeto (no array),
- **Entonces** debe lanzar un error con mensaje claro:  
  > “`[[user.id]]` requires 'user' to be an array. Received: [object Object]”.

### 2.5. Propiedad de ítem no existe

- **Dado** `[[payments.id]]`,
- **Y** un ítem en `payments` no tiene `id`,
- **Entonces** la celda debe contener `[[payments.id]]`.

### 2.6. Propiedad de ítem es `null`/`undefined`

- **Dado** `[[payments.amount]]`,
- **Y** un ítem tiene `amount: null`,
- **Entonces** la celda debe quedar vacía (`""`).

---

## 3. Comportamiento con fórmulas y estilos

### 3.1. Fórmulas se preservan y ajustan

- **Dado** una celda en fila repetible con fórmula `=B3*C3`,
- **Cuando** la fila se expande,
- **Entonces** cada nueva fila debe tener fórmula ajustada: `=B4*C4`, `=B5*C5`, etc.

### 3.2. Estilos se copian a nuevas filas

- **Dado** una fila con celdas con:
  - fondo azul,
  - borde grueso,
  - tipo de letra negrita,
- **Cuando** la fila se expande,
- **Entonces** todas esas propiedades deben copiarse fielmente a las nuevas filas.

### 3.3. Múltiples hojas

- **Dado** un archivo con 2 hojas,
- **Y** solo la primera contiene marcadores,
- **Entonces** la segunda hoja debe permanecer inalterada.

---

## 4. Coexistencia de `{{}}` y `[[]]`

### 4.1. Combinación válida

- **Dado** una fila con `{{user.name}}` y `[[payments.id]]`,
- **Y** `payments` es un array válido,
- **Entonces** la fila se repite N veces,
- **Y** `{{user.name}}` se resuelve contra el objeto raíz en cada repetición,
- **Y** `[[payments.id]]` se resuelve contra el ítem actual.

---

## 5. Entrada y salida

### 5.1. Entrada como Buffer

- **Dado** un `Buffer` de archivo `.xlsx` válido,
- **Y** un objeto `data` plano,
- **Entonces** debe devolver `Promise<Buffer>` del archivo resultante.

### 5.2. Salida válida

- **El Buffer resultante** debe abrirse sin errores en:
  - Microsoft Excel
  - Google Sheets
  - LibreOffice Calc

---

## 6. Arquitectura y stack técnico

### 6.1. Monorepo con pnpm

- **El paquete** debe estar en un workspace gestionado por `pnpm`.

### 6.2. ESM y CJS

- **Debe distribuirse** en ambos formatos: ESM y CJS.

### 6.3. Dependencia de ExcelJS

- **ExcelJS** debe ser una dependencia directa (no peer).
- **No debe traer polyfills** ni depender de navegador.

### 6.4. Testing con Vitest

- **Todos los tests** deben correr con Vitest.
- **La cobertura debe ser ≥90%**.

### 6.5. Tipos de TypeScript

- **La API debe estar completamente tipada** y generar `.d.ts`.

---

## 7. Comportamiento de errores

### 7.1. Error claro si clave no es array

- **Cuando** `[[]]` se usa con una clave que no es array,
- **Entonces** debe lanzar error con contexto: nombre del marcador, tipo recibido.

### 7.2. Error si hay mezcla de arrays en misma fila

- **Dado** una fila con `[[items.id]]` y `[[payments.id]]`,
- **Entonces** debe lanzar error: “Mixed array keys in row X: items vs payments”.

---

## 8. Preservación de contexto visual

### 8.1. Mantener fórmulas en celdas no interpoladas

- **Dado** una fila con `[[]]` y una celda con fórmula `=SUM(A:A)`,
- **Cuando** la fila se repite,
- **Entonces** la fórmula debe mantenerse en cada nueva fila.

### 8.2. Mantener merges

- **Dado** una fila con celdas mergiadas,
- **Cuando** se repite,
- **Entonces** las nuevas filas deben tener los mismos merges.

---

## 9. API pública

### 9.1. Nombre y firma

- **Debe exportar** `interpolateXlsx(options: { template: Buffer;  any })`.

### 9.2. Asincrónico

- **Debe retornar** `Promise<Buffer>`.

---

## 10. Documentación y uso

### 10.1. README debe incluir

- Ejemplo de uso básico.
- Explicación de `{{}}` vs `[[]]`.
- Comportamiento con arrays vacíos y claves faltantes.
