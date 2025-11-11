# Data Merge

## Descripción del Proyecto
Esta aplicación web permite unificar contactos provenientes de distintas plataformas académicas (por ejemplo, Galileo y Aula Extendida) en un formato único compatible con Outlook.  
Su objetivo es facilitar la creación de categorías y la exportación de contactos de forma automática y sin procesos manuales.

---

## Funcionalidades Principales
- Autenticación con cuenta Microsoft.  
- Obtención automática de las categorías y contactos existentes en Outlook.  
- Visualización detallada de contactos por categoría.  
- Creación de nuevas categorías mediante la carga de dos archivos Excel (de diferentes fuentes universitarias).  
- Generación de un archivo .csv exportable a Outlook.  

---

## Flujo General de Uso
1. El usuario ingresa a la aplicación e inicia sesión con su cuenta Microsoft.  
2. La aplicación recupera las categorías y contactos asociados a Outlook.  
3. Desde el panel principal, el usuario puede:
   - Ver contactos de cada categoría.
   - Crear una nueva categoría con nombre personalizado.  
4. Al crearla, el sistema solicita cargar dos archivos Excel, los procesa y genera un archivo unificado (.csv) listo para importar a Outlook.  

---

## Requisitos Técnicos
**Frontend:** React
**Backend:** Node.js
**Base de datos:** Se utilizo PostgreSQL, tambien se sincroniza con outlook para la toma de categorias y contactos de categorias.  
**APIs:** Integración con la API Microsoft Graph para la gestión de contactos y categorías.  

**Requisitos del sistema:**
- Navegador: Google Chrome, Edge o Firefox (versión actualizada).  
- Conexión a Internet.  
- Cuenta Microsoft válida.  

---

## Instalación y Ejecución Local
1. Clonar el repositorio:  
   ```bash
   git clone https://github.com/josesandovall99/Outlook-B.git


URL de la aplicacion: https://outlook-f.onrender.com/
