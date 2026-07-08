# M365 / Windows Modern Sync Review

## Resultado ejecutivo

`OneDriveMapper.ps1` queda intacto. El script principal sigue siendo un mapper WebDAV:

- Autentica contra SharePoint con Edge/CDP.
- Extrae cookies `FedAuth` y `rtFa`.
- Inyecta cookies en WinINET.
- Mapea con `NET USE` y el servicio `WebClient`.

Ese enfoque no es el camino moderno recomendado para sincronizar archivos de SharePoint/OneDrive. Para una prueba actual con Windows 11 25H2 y Windows Server 2022, agregue `Start-OneDriveLibrarySyncGui.ps1`, que usa:

- Microsoft Graph PowerShell con `Connect-MgGraph` interactivo/MFA.
- Microsoft OneDrive sync app (`OneDrive.exe`) para la sincronizacion real.
- La clave documentada `HKCU\Software\Policies\Microsoft\OneDrive\TenantAutoMount`.
- Una GUI WinForms para pedir cuenta y URL de biblioteca/carpeta.

## Archivos nuevos

- `Start-OneDriveLibrarySyncGui.ps1`: flujo GUI moderno para resolver sitio/biblioteca y configurar sincronizacion.
- `M365-OneDriveSync-Review.md`: esta nota de revision.

## Uso piloto recomendado

Ejecutar con Windows PowerShell 5.1 en STA:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Start-OneDriveLibrarySyncGui.ps1
```

Valores por defecto incluidos para la prueba:

- Cuenta: `victor.gonzalez@geexsa.com`
- URL: `https://gesex.sharepoint.com/sites/Sharepoint_Test_Informatica/archivos/Forms/AllItems.aspx`

El boton `Resolver y sincronizar` hace lo siguiente:

1. Verifica Windows, PowerShell, OneDrive.exe y modulos Graph.
2. Abre login interactivo de Microsoft Graph. La cuenta y MFA se completan en navegador/GUI Microsoft.
3. Resuelve el sitio y lista/biblioteca con Graph.
4. Genera el Library ID usado por OneDrive.
5. Escribe `TenantAutoMount` en HKCU si esta marcada la opcion.
6. Inicia `OneDrive.exe /background`.
7. Intenta abrir el dialogo inmediato de OneDrive con `odopen://sync`.
8. Abre la biblioteca en el navegador como verificacion/manual fallback.

## Validaciones realizadas

- `OneDriveMapper.ps1` estaba sin cambios pendientes antes de empezar.
- `OneDriveMapper.ps1` parsea correctamente.
- `Start-OneDriveLibrarySyncGui.ps1` parsea correctamente en PowerShell 7.
- `Start-OneDriveLibrarySyncGui.ps1` parsea correctamente en Windows PowerShell 5.1.
- La URL de prueba se parseo como:
  - host: `gesex.sharepoint.com`
  - sitio: `/sites/Sharepoint_Test_Informatica`
  - biblioteca: `/sites/Sharepoint_Test_Informatica/archivos`
- Se probo generacion sintetica de Library ID y `odopen://sync`.
- Modulos presentes localmente:
  - `Microsoft.Graph.Authentication` 2.36.1
  - `Microsoft.Graph.Sites` 2.36.1
  - `Microsoft.Graph.Files` 2.36.1
  - `Microsoft.Online.SharePoint.PowerShell` 16.0.26712.12000
- OneDrive encontrado:
  - `C:\Program Files\Microsoft OneDrive\OneDrive.exe`
  - version `26.113.0614.0004`

## Hallazgos importantes

1. `OneDriveMapper.ps1` no usa cmdlets Microsoft 365 actuales para resolver contenido; es WebDAV + cookies + WinINET + `NET USE`.
2. WebDAV depende de `WebClient`, zonas IE/WinINET y cookies `FedAuth/rtFa`; son puntos fragiles ante cambios de autenticacion moderna.
3. Para MFA y login GUI, el flujo nuevo delega autenticacion a `Connect-MgGraph` y a la UI oficial de OneDrive/Microsoft.
4. La autosincronizacion documentada por Microsoft puede aplicar en la proxima sesion de OneDrive y dentro de una ventana de hasta 8 horas.
5. `odopen://sync` se usa como intento inmediato de abrir el dialogo de sincronizacion, pero la base soportada/documentada para administracion es `TenantAutoMount`.
6. En Windows Server 2022/RDS/VDI, Microsoft documenta OneDrive con instalacion per-machine y consideraciones FSLogix/perfil. El script detecta Server y lo advierte, pero no cambia configuraciones FSLogix.
7. El script original no esta firmado con Authenticode. En equipos con politica `AllSigned`, fallara aunque la sintaxis sea valida.

## Referencias Microsoft revisadas

- OneDrive policies / `TenantAutoMount`: https://learn.microsoft.com/en-us/sharepoint/use-group-policy
- OneDrive per-machine install: https://learn.microsoft.com/en-us/sharepoint/per-machine-installation
- OneDrive on VDI / Windows Server 2022: https://learn.microsoft.com/en-us/sharepoint/sync-vdi-support
- Microsoft Graph interactive auth alias: https://learn.microsoft.com/en-us/powershell/module/microsoft.entra.authentication/connect-entra
- Graph site API: https://learn.microsoft.com/en-us/graph/api/site-get
- Graph list API: https://learn.microsoft.com/en-us/graph/api/list-list
- SharePoint Online MFA login example: https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
