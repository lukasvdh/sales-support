# Verpa Support Desk

Interne ticketing-applicatie voor Verpa Benelux NV. Gebouwd als single-page app op basis van MSAL.js, Microsoft Graph en SharePoint Lists. Gehost op Cloudflare Pages.

---

## Bestanden

| Bestand | Beschrijving |
|---|---|
| `index.html` | Shell van de app: layout, modals, document viewer |
| `style.css` | Alle opmaak |
| `app.js` | Configuratie, logica, Microsoft Graph-integratie |
| `logo.jpg` | Verpa-logo (sidebar en loginscherm) |
| `favicon.png` | Browsertabblad-icoon |

---

## Architectuur

```
Browser (MSAL.js)
    │
    ├── Microsoft Login (Azure AD)
    │       └── Scopes: User.Read · Sites.ReadWrite.All · Files.ReadWrite.All
    │
    ├── Microsoft Graph API
    │       ├── SharePoint List "Tickets"         ← tickets opslaan / ophalen
    │       └── SharePoint Drive "Tickets/"       ← bijlagen uploaden
    │
    └── Cloudflare Pages (verpa-support.pages.dev)
            └── Statische bestanden (geen server-side logica)
```

---

## Configuratie (`app.js`)

Bovenaan `app.js` staat het `CONFIG`-object. Pas dit aan bij een nieuwe deployment:

```javascript
const CONFIG = {
  clientId:    "e82e1484-0864-44a8-a8fe-279915eec8bf",  // Azure AD App Registratie
  tenantId:    "e65dbe4b-d1e2-4283-b0f5-aa7717e81077",  // Verpa Benelux tenant
  redirectUri: "https://verpa-support.pages.dev",        // Cloudflare Pages URL
  siteHostname:"verpabenelux.sharepoint.com",
  sitePath:    "/sites/OfficeData",
  listName:    "Tickets",
  attachFolder:"Tickets",
  adminRole:   "Admin",
  adminEmails: ["lukas@verpa.be"]
};
```

---

## Behandelaars

Het team dat tickets kan ontvangen en afhandelen:

- Lukas Vanderheyden
- Aniel Haeyaert
- Sten Huygens
- Yana Verspreet

Aanpassen via de `TEAM`-constante bovenaan `app.js`.

---

## Ticket categorieën & formulieren

### Verkoop
| Onderdeel | Velden |
|---|---|
| Prijswijzigingen | Onderwerp, Omschrijving |
| Afnamerapporten | Klant, Klantnummer, Periode, Niveau, Omschrijving |
| Artikelen aanmaken | Artikelnummer, Artikelnaam, Leverancier, Inkoopprijs, Verkoopprijs, Op webshop, Migratieartikel |
| Klant assortiment | Klantnaam, Klantnummer, Klantgroep, Artikel, Vervanging |
| Andere vragen | Onderwerp, Omschrijving |

### Technisch
| Onderdeel | Velden |
|---|---|
| Webshop | Onderwerp, Klantnaam, Klantnummer, Email, Omschrijving |
| Webshop login | Gebruikersnaam, E-mailadres, Klantnummer, Klantnaam, Assortiment, Type account |
| IT-Probleem | Onderwerp, Omschrijving |
| Business Central | Onderwerp, Omschrijving |
| Andere vragen | Onderwerp, Omschrijving |

---

## Ticket statussen

| Status | Betekenis |
|---|---|
| Open | Nieuw, nog niet opgepikt |
| In Behandeling | Toegewezen en actief |
| Opgelost | Afgehandeld |
| Gesloten | Definitief gesloten |

Prioriteiten: **Laag · Gemiddeld · Hoog**

Referentienummers worden automatisch gegenereerd in het formaat `VRP-YYYY-XXXX`.

---

## Snelfilters (sidebar)

- Webshop logins
- Artikelen aanmaken
- Hoge prioriteit
- Niet toegewezen

---

## Rollen

- **Beheerder (Admin):** kan alle tickets zien, toewijzen, statuswijzigingen doorvoeren en interne notities plaatsen. Rol toegekend via Azure AD App Role `Admin`.
- **Standaard gebruiker:** kan eigen tickets aanmaken en het gesprek opvolgen.

---

## SharePoint vereisten

Vóór eerste gebruik moeten volgende zaken aangemaakt zijn in SharePoint:

1. **SharePoint List** `Tickets` op de site `/sites/OfficeData`
2. **Documentbibliotheek** met map `Tickets/` voor bijlagen

De app maakt automatisch de benodigde kolommen aan bij de eerste keer opstarten (als de ingelogde gebruiker admin-rechten heeft op de lijst).

---

## Deployment (Cloudflare Pages)

1. Zorg dat de Azure AD App Registratie de redirect URI `https://verpa-support.pages.dev` heeft.
2. Upload alle bestanden (`index.html`, `style.css`, `app.js`, `logo.jpg`, `favicon.png`) via Cloudflare Pages Direct Upload.
3. De app werkt **niet lokaal** — Microsoft-login vereist een geregistreerde redirect URI.

---

## Technische stack

| Onderdeel | Technologie |
|---|---|
| Authenticatie | MSAL.js 2.x (OAuth 2.0 / Azure AD) |
| Data opslag | SharePoint Lists via Microsoft Graph REST API |
| Bijlagen | SharePoint Drive via Microsoft Graph |
| Hosting | Cloudflare Pages (statisch) |
| UI | Vanilla HTML/CSS/JS, geen frameworks |
| Fonts | Inter (Google Fonts) |
