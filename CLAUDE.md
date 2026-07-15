# Claude Protocol

> Este archivo es leído automáticamente por Claude al abrir este proyecto.
> Define el protocolo de eficiencia de tokens para todo el trabajo.

---

## Mapa del sistema (LEER PRIMERO)

Antes de tocar cualquier cosa, leé **`docs/SISTEMA.md`**: explica qué es el sistema,
cómo funciona, dónde está cada cosa, el modelo de datos, los canales (stock/línea/comprado),
roles/permisos, los flujos y —crítico— los **puntos sensibles de seguridad** y las
**reglas de "no romper"**. No re-explores todo el repo: ese doc es la fuente de verdad.

---

## Principio base

Completar cada tarea con el mínimo de tokens posible sin sacrificar correctitud.
Cada decisión —modelo, formato, contexto, paralelismo— sirve a ese objetivo.

---

## Selección de modelo

| Nivel | Modelo | Cuándo |
|-------|--------|--------|
| Simple | `claude-haiku-4-5` | Lectura, búsqueda, edición 1-3 líneas, rename, formato, docstrings |
| Medio | `claude-sonnet-4-6` | Cambios multi-archivo, debug, features, refactors acotados |
| Complejo | `claude-opus-4-6` | Arquitectura, refactors grandes, seguridad, diseño de sistema |
| Muy Complejo | `claude-fable-5` | Cosas extremadamentes complejas o planes muy elaborados |

Al despachar subagentes (`Task` tool): usar **siempre el modelo más barato** que pueda
resolver esa subtarea.

---

## Formato de respuesta

**Nunca:**
- Preambles ni postambles ("Claro, voy a...", "¿Hay algo más?")
- Explicaciones no pedidas
- Archivos completos cuando solo cambió una parte

**Siempre:**
- Diffs o bloques mínimos para código. Usar `// ... código existente` para saltar secciones.
- Respuesta directa desde el primer token.

---

## Gestión de contexto

- Ejecutar `/compact` antes de cambiar de tarea principal.
- No releer archivos ya en contexto — referenciar por path.
- Leer solo la sección relevante de archivos grandes (grep/head/tail primero).

---

## Batching y subagentes

- Agrupar ediciones relacionadas en un solo prompt.
- Subtareas independientes → subagentes paralelos con `Task`.
- Subtareas read-only (investigación, análisis) → subagente Haiku separado.
- No crear subagente si la tarea tarda menos de 2 tool calls.

---

## Por tipo de tarea

| Tarea | Modelo base |
|-------|-------------|
| Bug con stack trace claro | Haiku → fix directo |
| Bug intermitente | Sonnet |
| Feature < 50 líneas | Sonnet completo |
| Feature grande | Opus (diseño) + Sonnet (impl.) |
| Refactor mecánico | Haiku |
| Refactor de patrón | Opus (decisión) + Sonnet (impl.) |
| Infra / config | Haiku |
| Docs / comentarios | Haiku |
| Tests unitarios | Haiku (paralelo por archivo) |
| Tests integración | Sonnet |

---

## Checklist antes de cada tarea

1. ¿Qué complejidad tiene? → Elegir modelo
2. ¿Hay subtareas independientes? → Subagentes en paralelo
3. ¿El contexto acumuló trabajo viejo? → `/compact`
4. ¿Necesito el archivo completo o solo una sección? → grep primero
5. ¿El prompt pide una sola cosa concreta? → Si no, dividirlo
