# Emme & Malene – Eventplanlægger til SharePoint
Dette projekt er et event-modul til SharePoint, udviklet i React og TypeScript.
Modulet gør det muligt at oprette, administrere og tilmelde sig interne events
via SharePoint-lister.


## Teknologier
SharePoint Framework (SPFx)
React
TypeScript
PnPjs
Fluent UI


## Forudsætninger
Node.js (v18 eller nyere)
npm
SharePoint Online tenant
SPFx development environment


## Funktionalitet
Oprettelse og redigering af events (admin)
Tilmelding og framelding (bruger)
Kapacitetsstyring og venteliste
Filtrering af events


## Projektstruktur
/src
  /components
  /hooks
  /services
  /types

  
## Forfattere
Emma Sku Jensen
Malene Koch Lau

Udviklet som en del af bachelorprojektet på Professionsbachelor i Webudvikling.



# For at kunne køre projektet lokalt kræver det, at dit udviklingsmiljø er korrekt sat op. Følg nedenstående trin i den angivne rækkefølge.

## 1. Kontroller Node.js og npm version

Projektet kræver en specifik version af Node.js for at fungere korrekt sammen med SharePoint Framework (SPFx).

Åbn en terminal/console og kør følgende kommandoer:

node -v
npm -v
nvm -v


Node.js skal være version 20.11.1

npm skal være installeret (versionsnummer er underordnet)

nvm anvendes til håndtering af Node-versioner

Hvis du ikke bruger Node.js version 20.11.1, kan der opstå kompatibilitetsproblemer under installation eller ved kørsel af projektet.


## 2. Installer Gulp og SPFx-generator (kun én gang)

Disse værktøjer er nødvendige for at arbejde med SharePoint Framework-projekter og skal kun installeres én gang pr. computer.

Kør følgende kommando:

npm install -g yo gulp-cli @microsoft/generator-sharepoint


Når installationen er færdig, kan du verificere at værktøjerne er korrekt installeret:

yo --version
gulp -v


## 3. Hent projektet fra GitHub og installer afhængigheder

Når projektet er klonet fra GitHub, skal alle projektets afhængigheder installeres.

Navigér til projektets rodmappe og kør:

npm install


Dette installerer alle nødvendige packages, som projektet er afhængigt af.


## 4. Godkend udviklingscertifikat (trust dev certificate)

Før projektet kan køres første gang, skal SharePoint Frameworks udviklingscertifikat godkendes.

Kør følgende kommando én gang, inden du starter projektet første gang:

gulp trust-dev-cert


Dette er nødvendigt for at kunne køre projektet lokalt over HTTPS.


## 5. Start udviklingsserveren

Når opsætningen er fuldført, kan udviklingsserveren startes:

gulp serve


Som standard åbnes en browser automatisk, hvor SharePoint Workbench vises.

Hvis du ikke ønsker automatisk browseråbning, kan du i stedet bruge:

gulp serve --nobrowser
