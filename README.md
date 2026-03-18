# trimble-extension-rech

Extension Trimble Connect pour rechercher les éléments d'une maquette via le PSET "PSET - Attributs Mensura".

## Usage
- Charger l'extension dans Trimble Connect (manifest.json).
- Renseigner un ou plusieurs critères (NOM, ZONE, SOURCE, ENTREPRISE D'EXECUTION, plage de DATE).
- Lancer la recherche : les objets trouvés sont sélectionnés, mis en surbrillance et le zoom est ajusté.
- Un récapitulatif indique combien d'éléments correspondent à chaque critère saisi.

L'API est utilisée via `TrimbleConnectWorkspace.connect(window.parent, ...)` telle que décrite dans la documentation Workspace API.
