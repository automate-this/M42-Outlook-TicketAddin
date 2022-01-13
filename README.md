# M42 Outlook AddIn

# Adapt to your production environment
## src.taskpane.components.Constants
- DefaultCategoryGUID
- DefaultInitiatorGUID
## webpack.config.js
- urlProd

# Local Test
## Prerequisites
- cors-anywhere
    PS C:\VSCodeWorkspace\cors-anywhere> npm install cors-anywhere
    PS C:\VSCodeWorkspace\cors-anywhere\node_modules\cors-anywhere> npm start
- change M42 ServiceStore Service URL in plugin settings to:
    http://localhost:8080/https://YOUR.PRODUCTION.URL/M42Services
- Outlook -> Datei -> Add-Ins verwalten -> Meine Add-Ins -> Benutzerdefiniertes Add-In hinzufügen
    -> aus Datei hinzufügen -> manifest.xml aus m42-ticket-addin - Ordner auswählen (nicht vom dist-Ordner)
- npm run dev-server

# Build
- npm run build

## Set up your dev environment
- https://docs.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment

## Screenshots
![image](https://user-images.githubusercontent.com/81413189/149336748-707df9dc-9e9b-4454-b83c-72a10142917c.png)

