# Reglas del Proyecto Q-Centros

## Control de Versiones Automático (Obligatorio)

Existe una **única versión global para todo el programa**.

Cada vez que se realice una modificación, mejora o corrección en cualquier módulo o archivo del proyecto:
1. **Incrementar la versión global** siguiendo SemVer:
   - Incremento de parche (`v1.0.X`): para correcciones de errores, ajustes visuales o cambios menores.
   - Incremento menor (`v1.X.0`): para nuevas funcionalidades o mejoras en el sistema.
   - Incremento mayor (`vX.0.0`): para cambios arquitectónicos o rediseños globales.
2. **Actualizar la fecha** a la fecha actual (`DD/MM/AAAA`).
3. **Registrar la mejora** en el archivo `VERSIONES.md`, añadiendo una fila en la tabla con la nueva versión, fecha y resumen de cambios.
4. **Mantener la coherencia** si la versión se muestra en la interfaz web o archivos de configuración.
