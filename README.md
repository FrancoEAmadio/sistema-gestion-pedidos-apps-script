# Sistema de Gestión Logística RDP (Relación de Pedido)

## 📝 Descripción del Proyecto
Sistema integral de automatización y gestión documental desarrollado para la **División IV Material del Ejército Argentino**. 

El proyecto nació para dar solución a un cuello de botella operativo: la recepción manual y en papel de las solicitudes de suministros entre divisiones. Para resolverlo, diseñé una arquitectura centralizada utilizando el ecosistema de **Google Workspace** y un backend en **Google Apps Script (JavaScript)**.

## 🔄 ¿Cómo funciona el circuito?
1. **Ingreso de Solicitudes:** Las dependencias completan un **Google Form** estandarizado con los detalles de la Relación de Pedido (categoría, efectos solicitados, cantidades y firmas autorizadas).
2. **Centralización de Datos:** La información se almacena automáticamente en una base de datos centralizada en **Google Sheets**.
3. **Gestión Documental Automatizada:** A través del script, el sistema gestiona los archivos de respaldo en **Google Drive** de manera estructurada según la carpeta raíz (`ROOT_FOLDER_ID`).
4. **Seguimiento de Estados:** El flujo de trabajo permite administrar en tiempo real los diferentes estados logísticos del pedido:
   - `EN REVISIÓN`
   - `RECIBIDO` / `RECHAZADO`
   - `FIRMADO`
   - `EN DEPÓSITO DE SUMINISTROS`
   - `RETIRAR` / `ENTREGADO`

## 🛠️ Tecnologías y Herramientas Utilizadas
* **Lenguaje:** JavaScript (ES6)
* **Backend & Automatización:** Google Apps Script (GAS)
* **Base de Datos & Interfaz:** Google Sheets y Google Forms
* **Almacenamiento:** Google Drive API

## 💡 Impacto Operativo
Este desarrollo eliminó por completo el uso de papel y la dispersión de correos electrónicos informales, unificando la trazabilidad de los insumos y agilizando la toma de decisiones por parte de los mandos superiores. Demuestra la capacidad de aplicar programación y lógica de sistemas para optimizar la logística en entornos institucionales y corporativos exigentes.
