# Zielmodell IT-Ablauforganisation — „Homebase & Mannschaft"

> Betriebs- und Zusammenarbeitsmodell für die neue IT-/IPS-Organisation.
> Leitidee: **Die Org-Einheit ist die Homebase der Mitarbeitenden. Gespielt wird
> in Mannschaften über die Einheiten hinweg.**

Begleitende Deliverables:
- **Interaktive Fassung (Web):** siehe Artifact-Link in der Konversation
- **CIO-Kurzfassung (PPTX, 15 Slides):** `2026_07_08_IT_Ablauforganisation_Homebase_Mannschaft.pptx`
  (erzeugt via `build_org_operating_model.py`) — inkl. Ist/Soll der Operations-Koordination
  (Slides 10–11), SDM-Rollenprofil (12), Migrations-Roadmap (13), SDM-Pool-Modell (14)
  und Verortungsentscheidung zentral vs. verteilt (15)
- **Rollenbeschreibung SDM:** `Rollenbeschreibung_Service_Delivery_Manager.md`
  (inkl. Pool-Modell, Head of Service Delivery, Cluster-Zuordnung)
- **Kapazitäts-Allokationsmatrix (Excel-Template):**
  `2026_07_08_Kapazitaets_Allokationsmatrix.xlsx` (erzeugt via `build_capacity_matrix.py`)

---

## 1. Wie wird Operations geklammert?

Operations liegt heute an vier sinnvollen Stellen — bewusst **nah an der
Technologie**. Diese Nähe soll erhalten bleiben; Operations wird **nicht in eine
Säule zentralisiert**, sondern über eine **SIAM-Integrationsschicht** geklammert.

**Die vier Operations-Domänen (Homebase):**

| Domäne | Heimat im Org-Chart | Umfang |
|---|---|---|
| Platform & Infra Ops | Technical Solutions | Base ITSM Ops, Service Desk, Datacenter, Netzwerk, Compute |
| Application / Business Ops | Business Solutions | Solution Operations, SAP AMS, Release & Deployment |
| Local IT Operations | Technical Solutions (dezentral) | Local Service Desk, Device- & Server-Support vor Ort |
| Security Operations | Cyber Defense | SOC, Threat Detection, Incident Response |

**Die Klammer besteht aus vier Bausteinen:**

1. **Service Integrator (SIAM-Funktion)** — orchestriert interne Ops **und**
   Dienstleister als ein Liefernetzwerk. Heimat: *IT Service Management* im
   IPS Management Office (Governance).
2. **End-to-End Service Owner** je Service — über alle Einheiten hinweg
   verantwortlich für Ergebnis, Kosten, Qualität, SLA.
3. **Gemeinsame ITSM-Prozesse & Tool-Kette** — ein Incident-/Problem-/Change-/
   Request-Prozess (ITIL 4) über alle Einheiten und Provider, auf ServiceNow.
4. **Operations Command Layer** — Major-Incident-Management, Service Operations
   Center und ein Operations-/SIAM-Board als tägliche/wöchentliche Steuerung.

Dienstleister werden über **OLAs/UCs** an dieselbe Prozess- und SLA-Logik
gekoppelt wie interne Teams.

---

## 2. Zielbild Ablauforganisation — vier Steuerungsebenen

Die Aufbauorganisation (Chart) bleibt die Homebase. Die Ablauforganisation legt
sich als vier Ebenen darüber (Prinzip „Flight Levels").

| Ebene | Fokus | Frequenz | Gremien / Teams |
|---|---|---|---|
| **1 · Strategie** | Steuerung & Portfolio | quartalsweise | IPS Steering, Portfolio-/Investment-Board, Architecture Review Board, Capacity-Board |
| **2 · Taktik** | Demand, Programme, Services | monatlich / 2-wöchentlich | Value Streams, Demand & Sponsoring, CAB, Service Management Review |
| **3 · Operativ** | Delivery & Betrieb | wöchentlich / täglich | Squad/Team Daily, Ops Sync, Major Incident Mgmt, Provider Jour Fixe |
| **4 · Homebase** | Fachliche Heimat & Chapters | 2-wöchentlich | Linien-Jour-Fixe, Chapters/CoP, Skill- & Kapazitätsplanung |

Eine Mitarbeiterin hat **drei Zugehörigkeiten** mit klar getrennten Zwecken:
ihre **Homebase** (Ebene 4), ihr **Value-Stream-Team** (Ebene 2/3) und ihre
**Community of Practice**.

---

## 3. Zusammenarbeitsmodell — drei Typen virtueller Teams

| Typ | Zweck | Besetzung | Führung |
|---|---|---|---|
| **Value-Stream- / Service-Teams** (vertikal) | End-to-End-Ergebnis eines Services/Produkts | dauerhaft, cross-funktional aus mehreren Homebases | Service / Product Owner + Delivery Lead |
| **Chapters / Communities of Practice** (horizontal) | Standards, Skills, Wiederverwendung, Karriere | gleiche Disziplin über alle Teams | Chapter Lead / CoP Lead |
| **Boards / Gremien** (Governance) | übergreifende Entscheidungen & Priorisierung | Entscheider aus allen Bereichen | Board-Chair |

**Faustregel:** *Value Stream* liefert Wert · *Chapter* baut Können · *Board*
trifft Entscheidungen. Ein Mensch ist in **einem** Value Stream, **einer**
Chapter — und nur bei Bedarf in Boards.

### Konkrete Value-Stream-Schnitte (aus dem Org-Chart abgeleitet)

| Value Stream (Mannschaft) | Service-Owner-Heimat | Beitragende Homebases |
|---|---|---|
| Modern Workplace & Collaboration | Technical Sol. › Workplace Solutions | Local IT Ops · Cyber (Endpoint) · Business Sol. (Collab-Apps) |
| Network & Connectivity | Infra & Platforms › Network & Security Infra | Local Network Support · Cyber (Network Sec.) |
| Compute, Cloud & Datacenter | Infra & Platforms › Compute Platform Mgmt | Local Server Support · SAP Basis · Cyber |
| Identity & Access (IAM) | Infra & Platforms › Authentication / IDM | Cyber (Identity Sec.) · Business Solutions |
| SAP & Business Applications | Business Sol. › Solution Operations (AMS) | Technical Sol. (SAP Basis, Release & Deployment) |
| Data, Analytics & KI | Business Sol. › Reporting & Analytics & KI | EA & Innovation · Master Data Management |
| Digital Service Desk / ITSM | Technical Sol. › Base ITSM Ops / Service Desk | Local Service Desk · alle L2/L3 · SIAM |

*Quer dazu:* Cyber-Security-Services (SOC) werden von allen Streams konsumiert;
Owner bleibt Cyber Defense. Optional OT / Shopfloor-IT als eigener Stream.

**Empfohlene Chapters/CoPs:** Operations & Service Desk · Platform/Cloud
Engineering · Automation/DevOps · Architektur · Data/BI · *Cyber Security
Community* (bereits im Chart vorhanden).

---

## 4. Ritual-Kalender

| Ritual / Meeting | Ebene | Frequenz | Teilnehmer | Zweck & Output |
|---|---|---|---|---|
| IPS Steering / Portfolio-Board | 1 | Quartal | CIO, Bereichsleitungen, Portfolio-Mgmt | Priorisierung, Budget & Kapazität → **Roadmap** |
| Architecture Review Board | 1 | Quartal + on demand | EA, Lead-Engineers, Security-Architektur | Standards, Design-Reviews → **ADR** |
| Quarterly Business Review | 1 | Quartal | Service Owner, Business Relationship Mgr | Wertbeitrag & OKR je Wertstrom → **Ziele Q+1** |
| Demand & Sponsoring | 2 | Monat | Demand Mgmt, Owner, Sponsoren | Bewertung Anforderungen → **Backlog** |
| Service Management Review | 2 | Monat | Service Integrator, Owner, Provider-Mgr, Prozess-Owner | SLA/OLA, Provider-Steuerung → **Maßnahmen** |
| CAB — Change Advisory Board | 2 | Woche | Change Mgr, Ops-Teams, Security, Provider | Risikobewertung → **Change-Freigaben** |
| Value-Stream Planning | 2 | 2 Wochen | Owner, Squad, Architektur | Increment-Planung → **Team-Backlog** |
| Ops Sync / Service Desk Standup | 3 | täglich | Service Desk, L2/L3, Local IT, Provider | Lage, Incidents → **Tageslage & Zuweisungen** |
| Major Incident Management | 3 | bei Bedarf | MIM-Lead, Ops/Sec, Comms, Provider | P1/P2 → **Wiederherstellung + Post-Mortem** |
| Squad Daily | 3 | täglich | Value-Stream-Team | Fortschritt & Hindernisse → **Board** |
| Provider Ops Jour Fixe | 3 | Woche | Service Integrator, Provider-Lead, Owner | Provider-Steuerung → **Aktionsliste** |
| Chapter / CoP Meetup | 4 | 2 Wochen | Chapter-Mitglieder, Chapter Lead | Standards, Skill-Sharing → **Guidelines** |
| Linien-Jour-Fixe (Homebase) | 4 | 2 Wochen | Linienführung + Mitarbeitende | Personal, Skill, Kapazität → **Kapazitätsplan** |
| Retrospektive (Team & System) | 2–3 | 2–4 Wochen | Team; quartalsweise systemweit | Verbesserung → **Maßnahmen** |

> Prinzip: **Jedes Ritual hat genau einen Output.** Meetings ohne Output werden
> gestrichen.

---

## 5. Rollen für die virtuellen Teams

| Rolle | Zugehörigkeit | Verantwortung |
|---|---|---|
| **Service / End-to-End Service Owner** | Mannschaft | Ergebnis, Kosten, Qualität, SLA über alle Einheiten & Provider; priorisiert Service-Backlog. *Wichtigste neue Rolle.* |
| **Product Owner** | Mannschaft | Wert & Backlog produktorientierter Wertströme |
| **Delivery / Team Lead („Kapitän")** | Mannschaft | führt cross-funktionales Team durch die Lieferung, räumt Hindernisse |
| **Linienführungskraft (Homebase / People Lead)** | Homebase | Einstellung, Entwicklung, Kapazitätszusage, fachliche Exzellenz |
| **Chapter / CoP Lead** | übergreifend | Standards, Wiederverwendung, Skill-Level einer Disziplin |
| **Service Integration Manager (SIAM Lead)** | übergreifend | Ops + Provider als ein Netzwerk, End-to-End-Prozesse, OLAs/UCs |
| **Prozess-Owner** (Incident/Problem/Change/Request) | übergreifend | je ein ITSM-Prozess über alle Einheiten |
| **Capacity / Resource Manager** | Homebase-übergreifend | vermittelt Angebot & Nachfrage — „Schiedsrichter der Matrix" |

---

## 6. Homebase vs. Mannschaft — Spielregeln der Matrix

**Homebase entscheidet über:** Einstellung/Entwicklung/Beförderung · Skill-Profile
& Standards · verfügbare Kapazität · Personalthemen & Disziplinarik.

**Mannschaft entscheidet über:** Priorisierung & Lieferung · tägliche Arbeit im
Team · Service-/Produkt-Ergebnis, SLA, Roadmap · fachliche Zuweisung.

**Vier nicht verhandelbare Spielregeln:**

1. **Eine Priorität pro Person und Zeitraum.** Konflikte lösen Service Owner +
   Homebase-Lead, im Zweifel Capacity-Board — nie die Mitarbeiterin selbst.
2. **Kapazität wird zugesagt, nicht angenommen.** Homebase committet %-Anteile
   pro Mannschaft transparent (z. B. 70 % Stream A, 20 % Chapter, 10 % Linie).
3. **Dediziert vor geteilt.** Splitting über > 2 Teams wird vermieden.
4. **Ein Eskalationspfad.** fachlich → Service Owner; personell → Homebase-Lead;
   Ressourcenkonflikt → Capacity-Board.

### RACI

Legende: **R** Responsible · **A** Accountable · **C** Consulted · **I** Informed

| Entscheidung / Aktivität | Homebase Lead | Service Owner | Delivery Lead | Chapter Lead | SIAM / Capacity | CIO / Portfolio |
|---|:--:|:--:|:--:|:--:|:--:|:--:|
| Einstellung & Personalentwicklung | **A** | C | C | C | I | I |
| Skill-Standards & fachliche Qualität | C | I | I | **A** | I | I |
| Kapazitäts-Zusage an Mannschaften | R | C | I | I | **A** | I |
| Service-/Produkt-Priorisierung (Backlog) | I | **A** | R | C | I | C |
| Tägliche Delivery / Sprint-Ausführung | I | C | **A** | I | I | I |
| Betrieb & SLA-Einhaltung (End-to-End) | C | **A** | R | I | R | I |
| Provider-/Dienstleister-Steuerung | I | C | I | I | **A** | I |
| Major Incident (P1/P2) Steuerung | I | C | R | I | **A** | I |
| Ressourcen-Konflikt zwischen Teams | C | C | I | I | R | **A** |
| Portfolio- & Budget-Priorisierung | C | C | I | I | C | **A** |

> Der Kern: **Kapazität = Homebase accountable, Priorität = Service Owner
> accountable.** Genau ein **A** pro Zeile.

---

## 7. Einführung in vier Schritten

1. **Klammer setzen** — 5–8 Kern-Services definieren, je einen End-to-End
   Service Owner benennen, SIAM-Funktion im IPS Management Office verankern.
2. **Prozesse vereinheitlichen** — ein ITSM-Prozessmodell auf ServiceNow
   harmonisieren, OLAs/UCs mit Dienstleistern schließen.
3. **Rhythmus starten** — mit wenigen Ritualen beginnen (Ops Sync, CAB, Service
   Management Review, Linien-Jour-Fixe), jedes mit einem Output; nach 1–2
   Quartalen nachschärfen.
4. **Matrix leben** — mit 1–2 Pilot-Value-Streams starten (z. B. Workplace),
   %-Kapazitätszusagen transparent machen, Capacity-Board als Eskalationsinstanz
   etablieren.

---

## 8. SLA-/OLA-Skizze je Value Stream

Der Service Owner gibt dem Business eine **SLA**. Diese wird **back-to-back**
zerlegt in interne **OLAs** (mit den beitragenden Homebases) und externe
**Underpinning Contracts (UC)** mit den Dienstleistern.

**Goldene Regel:** `UC-Ziel ≤ OLA-Ziel ≤ SLA-Ziel` — jede Provider-Zusage ist
strenger als die interne, jede interne strenger als die Business-Zusage. So
entstehen keine Deckungslücken.

**Verfügbarkeitsklassen**

| Klasse | Verfügbarkeit | Kernzeit | Beispiele |
|---|---|---|---|
| Platinum | 99,9 % | 24×7 | IAM, Netzwerk, Compute, SAP-Prod |
| Gold | 99,5 % | 5×11 + Rufbereitschaft | Workplace, Data, Service Desk |
| Silver | 99,0 % | 5×11 | nachrangige Services |

**Prioritäten (Reaktion / Wiederherstellung)**

| Priorität | Reaktion | Wiederherstellung |
|---|---|---|
| P1 kritisch | 15 min | ≤ 4 h (bzw. Tier) |
| P2 hoch | 30 min | ≤ 8 h |
| P3 mittel | 4 h | 2 AT |
| P4 niedrig | 8 h | 5 AT |
| Service Request | – | 3 AT |

**SLA/OLA je Value Stream**

| Value Stream | Ziel-SLA | Kernzeit | P1 Restore | OLA-Bausteine (intern) | Underpinning Contracts (Provider) |
|---|---|---|---|---|---|
| Modern Workplace | Gold 99,5 % | 5×11 +Rufb. | 4 h | Client Platform Ops (WS) · Endpoint-Sec (Cyber) · Collab-Apps (BusSol) · Vor-Ort (Local) | Field-/Deskside-Services · MDM-Provider · Print-Services |
| Network & Connectivity | Platinum 99,9 % | 24×7 | 2 h | Network Ops (Infra) · Local Network (Local) · Network-Sec (Cyber) | WAN-Carrier · SD-WAN Managed Service |
| Compute, Cloud & DC | Platinum 99,9 % | 24×7 | 2 h | Platform Ops (Infra) · Server Vor-Ort (Local) · Cloud-Sec (Cyber) | Hyperscaler · Colocation-DC · Backup-Provider |
| Identity & Access | Platinum 99,9 % | 24×7 | 1 h | IDM Ops (Infra) · Identity-Security (Cyber) | IAM-SaaS (Entra/Okta) · PKI-Provider |
| SAP & Business Apps | Platinum 99,9 % | 6×14 | 2 h | AMS (BusSol) · SAP Basis (AppTech) · Release & Deploy | SAP-AMS-Dienstleister · Hosting (RISE/Hyperscaler) |
| Data, Analytics & KI | Gold 99,5 % | 5×11 | 8 h | BI Ops (BusSol) · Data Platform (EA/Infra) · MDM | Cloud-Analytics-Provider |
| Service Desk / ITSM | Gold 99,5 % | 24×7 (SD) | Dispatch < 15 min | Base ITSM Ops · Local Service Desk · ServiceNow-Plattform | Service-Desk-Dienstleister (1st Level) · ServiceNow (SaaS) |

*Quer:* SOC (Cyber Defense) — 24×7-Monitoring, SecInc-Reaktion 30 min, UC mit
MDR-Provider; als Sicherheits-OLA in jeden Stream eingebunden.

Governance: SLAs & OLAs werden **monatlich** im Service Management Review geprüft,
UCs vom SIAM/Provider-Management verantwortet.

---

## 9. Kapazitäts-Allokationsmatrix (Capacity-Board)

Das Excel-Template `2026_07_08_Kapazitaets_Allokationsmatrix.xlsx` macht sichtbar,
wie viel Kapazität jede **Homebase** (Linie) an welche **Mannschaft** (Value
Stream) zusagt — und ob die Summe die geplante Nachfrage deckt.

**Beispiel-Allokation (%, je Zeile = 100 %)**

| Homebase | FTE | M.Work | Netz | Compute | IAM | SAP | Data | Svc Desk | Chapter | Run&Line |
|---|--:|--:|--:|--:|--:|--:|--:|--:|--:|--:|
| Workplace Solutions | 18 | 70 | – | – | – | – | – | 5 | 15 | 10 |
| Infrastructure & Platforms | 30 | – | 25 | 30 | 20 | – | – | – | 10 | 15 |
| Local IT Operations | 22 | 30 | 10 | – | – | – | – | 40 | 5 | 15 |
| Business Solutions (Design/AMS) | 26 | – | – | – | 5 | 40 | 25 | – | 15 | 15 |
| Application Technology (SAP Basis/Dev) | 14 | – | – | – | – | 60 | 10 | – | 15 | 15 |
| Cyber Defense / SecOps | 12 | 5 | 5 | 5 | 10 | – | – | 5 | 20 | 50 |
| Cyber GRC | 8 | – | – | – | 10 | – | – | – | 30 | 60 |
| EA & Innovation | 6 | – | – | – | – | – | 20 | – | 40 | 40 |

**Angebot vs. Nachfrage (FTE, Beispiel)**

| Value Stream | Angebot | Nachfrage | Deckung (Gap) |
|---|--:|--:|--:|
| Modern Workplace | 19,8 | 22 | **−2,2** |
| Network & Connectivity | 10,3 | 10 | +0,3 |
| Compute, Cloud & DC | 9,6 | 10 | **−0,4** |
| Identity & Access | 9,3 | 9 | +0,3 |
| SAP & Business Apps | 18,8 | 18 | +0,8 |
| Data, Analytics & KI | 9,1 | 10 | **−0,9** |
| Service Desk / ITSM | 10,3 | 10 | +0,3 |

Bei Unterdeckung (rot) entscheidet das Capacity-Board: Nachfrage priorisieren,
Kapazität umschichten, Dienstleister zuschalten oder Skill über Chapter aufbauen.
Angebot rechnet sich automatisch aus `%-Allokation × FTE`; die Werte oben sind
Beispielwerte.

**Rhythmus:** quartalsweise Grundallokation (Portfolio-Board), monatliche
Nachjustierung (Service Management Review).

---

## 10. Operations-Koordination in der IPS — Ist & Soll

### Step 1 · Status quo (Ist)

Heute koordiniert **jeder Service Owner seinen Support selbst**. Beispiel BSS:
ein Service = Business-Applikation + Applikationsbetreuung durch **DXC (AMS)** +
zugelieferte Services aus **TSS** (z. B. Windows Server) + **zentraler Service
Desk** + ggf. **Regional Support Center** und weitere AMS-Dienstleister. Alle
Koordinationslinien laufen manuell beim Service Owner zusammen.

**Probleme, die daraus entstehen:**

| # | Problem | Wirkung |
|---|---|---|
| P1 | Rollenüberladung Service Owner | Service fachlich verantworten UND alle Service Levels koordinieren = zwei Jobs |
| P2 | Keine Klammerfunktion | Niemand ist übergeordnet für alle Supportaktivitäten zuständig |
| P3 | Inkonsistente Support-Modelle | Jeder Service wird anders betrieben, kein Standard, schwer auditierbar |
| P4 | Fragmentierte Provider-Steuerung | Jeder SO steuert DXC & Co. einzeln — keine Skaleneffekte, uneinheitliche SLAs |
| P5 | Kein End-to-End-SLA über die Kette | App (DXC) + Infra (TSS) + Desk je eigene Zusage, niemand verantwortet das Ganze |
| P6 | Ungeregelte Schnittstellen BSS↔TSS↔AMS↔SD | An Übergaben versanden Tickets & Eskalationen |
| P7 | Governance läuft ins Leere | Practice Supplier Management setzt Standards, niemand operationalisiert sie je Service |
| P8 | Personen- statt Prozessabhängigkeit | „SO kennt Service am besten" → Key-Person-Risiko, keine Skalierung, keine Transparenz |

**Kernbefund:** Die IPS hat die richtigen Bausteine (Service Owner, Practice
Supplier Management, Service Desk, AMS), aber keinen Dirigenten. Koordination ist
Privatsache des Service Owners — genau die Lücke, die BSS mit dem Ruf nach einer
„dedizierten Ressource" beschreibt.

### Step 2 · Zielmodell (Soll) — die Operations-Klammer

Die Klammer trennt **WAS** (Ergebnis) von **WIE** (Koordination):

- **Service Owner** bleibt End-to-End-**accountable** für das Service-**Ergebnis**
  (WAS) — gibt Ziele, SLA und Priorität vor, koordiniert aber nicht mehr selbst.
- **Service Delivery Manager (SDM, neu)** übernimmt die operative **Koordination**
  (WIE): steuert alle AMS-Dienstleister *und* internen Support-Einheiten (TSS,
  Service Desk, RSC) je Service oder Service-Cluster. Das ist genau die
  „dedizierte Ressource", die BSS fordert — als **Rolle einer Klammerfunktion**,
  nicht als Improvisation je SO.
- **Service Integration & Operations (SIAM, neu)** ist die Klammer selbst:
  End-to-End-Prozesse, Provider-Integration, konsolidierte Service-/Kostensicht.
  Enthält Service Level Management, Major Incident Management und die zentralen
  Prozess-Owner (Incident/Problem/Change/Request).
- **Practice Supplier Management** bleibt die Governance-Disziplin (Verträge,
  SLA/KPI-Rahmen, ServiceNow-Stammdaten, Meeting-Standards) — jetzt mit einem
  operativen Gegenüber, das die Standards je Service operationalisiert.

**Fluss:** Service Owner (Ziel & SLA) → SDM (plant & steuert die Kette) →
Provider + interne Einheiten liefern (OLA/UC) → SIAM misst End-to-End-SLA & eskaliert.

**SDM-Pool statt Einzelstelle:** Es gibt *mehrere* SDMs — je Service-Cluster einen,
geführt von einem **Head of Service Delivery** im IPS MO. Alle SDMs teilen dieselbe
Heimat (Solid-Line, zentral) und werden fachlich an die Cluster ausgeliehen
(Dotted-Line, dezentral). **Verteilter Sitz ist gut, verteilte Berichtslinie in die
Liefereinheiten ist riskant** (Peer-koordiniert-Peer, Standard-Drift, Rückfall in
Silos). Ein Head of SDM zentral mit Solid-Line zu co-located SDMs ist tragfähig —
faktisch der zentrale Pool mit dezentralem Einsatz. Details siehe
`Rollenbeschreibung_Service_Delivery_Manager.md`.

**Zusätzlich benötigte Rollen/Einheiten**

| Rolle | Neu? | Verantwortet | Heimat |
|---|---|---|---|
| Service Integration & Operations (SIAM) | **neu** | Klammer: E2E-Prozesse, Provider-Integration, konsolidierte Sicht | IPS Management Office |
| Service Delivery Manager (SDM) | **neu** | Operative Koordination aller Provider + interner Support je Cluster | SIAM-Klammer (je Cluster) |
| Service Level / Major Incident Mgr, Prozess-Owner | **neu** | SLA-Reporting, P1/P2-Steuerung, durchgängige ITSM-Prozesse | SIAM-Klammer |
| Service Owner | Bestand | Ergebnis, Business-Wert, SLA-Zusage, Priorität (WAS) | BSS / Fach-Einheit |
| Practice Supplier Management | Bestand | Provider-Governance (Verträge, SLA/KPI, ServiceNow-Stammdaten) | IPS Management Office |
| Service Desk / RSC / AMS | Bestand | Liefererbringung auf ihrem Level, gekoppelt via OLA/UC | TSS / Region / Provider |

### Wo wird geklammert — und warum dort?

**Empfehlung: die SIAM-Klammer ins IPS Management Office** — nicht in eine der
vier Liefer-Einheiten.

- **Neutralität:** Das IPS MO steht *über* BSS, TSS und Cyber. Läge die Klammer in
  TSS, würde eine Liefer-Einheit die Provider und den Support der anderen steuern
  — Interessenkonflikt und fehlendes Mandat.
- **Andockpunkte sind da:** Governance, IT Service Management, Practice Supplier
  Management, Kapazitäts-/Ressourcenplanung, Lizenz- & Financial Management, PMO —
  alle bereits im IPS MO.
- **ServiceNow-Nähe:** Stammdaten, SLA/KPI und Dashboards werden dort schon
  governt; die operative Koordination setzt darauf auf.

**Abgrenzung zur Alternative „Klammer in TSS":** TSS bleibt der Ort der
*technischen* Betriebsausführung (Infra-Ops, Service-Desk-Betrieb). Die
*übergreifende* Koordination braucht ein neutrales Mandat oberhalb der
Liefer-Einheiten. Praktischer Schnitt: **Betriebsausführung in TSS,
Betriebs-Integration/-Koordination (SIAM) im IPS MO** — mit fachlicher
Dotted-Line zwischen SDM und den TSS-Ops-Teams.

**Koordinations-Rituale:** Service Operations Review je Cluster (wöchentlich,
SDM-geführt) · Supplier & Operations Board (monatlich, SIAM + Practice Supplier
Management) · zentraler Major-Incident-Prozess (bei Bedarf) · ServiceNow als
führendes Tool.

**Fazit:** Keine neue Säule — eine Klammer. `Service Integration & Operations
(SIAM)` im IPS Management Office, mit dem `Service Delivery Manager` als operativem
Koordinator je Service-Cluster. Das entlastet die Service Owner, bündelt die
Provider-Steuerung und schließt die Schnittstellen zwischen BSS/TSS/AMS/Service Desk.

---

## 11. Migrationsvorgehen — vom „pro Service Owner" zur SIAM-Klammer

Wellenweise über ~12 Monate, **ohne Big Bang**. Der Service Owner behält seine
Accountability; die Koordination wandert schrittweise in die Klammer.

| Phase | Dauer | Fokus | Kernaktivitäten |
|---|---|---|---|
| **0 · Setup & Mandat** | Wo. 1–4 | Entscheidung & Auftrag | CIO-Entscheid (Klammer im IPS MO) · SIAM-Lead benennen · Zielbild & Guardrails |
| **1 · Foundation** | M 1–3 | Fundament legen | ITSM-Prozesse auf ServiceNow harmonisieren · Practice Supplier Mgmt stärken (SLA/KPI/Stammdaten) · Service-Kataster: Einheiten & Provider je Service mappen |
| **2 · Pilot** | M 3–6 | Beweis am Einzelfall | 1 Pilot-Cluster (z. B. SAP: DXC + TSS Windows + Service Desk) · SDM besetzen · OLAs/UCs schließen · Service Operations Review starten · E2E-SLA messen · Lessons Learned |
| **3 · Rollout** | M 6–12 | Skalieren | Cluster wellenweise anbinden, SDM-Team aufbauen · Supplier & Operations Board etablieren · Major-Incident-Prozess zentralisieren · Dashboards produktiv |
| **4 · Stabilisierung & CI** | ab M 12 | Verstetigen | Betriebsmodell als Standard · KPI-Steuerung · kontinuierliche Verbesserung |

**Zentrale Risiken & Gegenmaßnahmen**

| Risiko | Gegenmaßnahme |
|---|---|
| Widerstand der Service Owner (Machtverlust) | Klare WAS/WIE-Trennung — SO behält Accountability und wird entlastet |
| „SO kennt Service am besten" | SDM arbeitet MIT dem SO, strukturierter Wissenstransfer, SO bleibt fachlicher Sponsor |
| Verträge nicht back-to-back (UC ≤ OLA ≤ SLA) | Practice Supplier Management priorisiert UC-Anpassung |
| Tool-Fragmentierung | ServiceNow als führendes System (Single Source of Truth) |
| Überlastung des SDM | Klare, nicht zu breite Cluster-Schnitte; realistische Service-Anzahl je SDM |

**Quick Wins** (früh sichtbar): zentraler Major-Incident-Prozess und ein
gemeinsames End-to-End-SLA-Dashboard — beide entfalten sofort Wirkung, auch
bevor alle Cluster migriert sind.

---

*Referenzrahmen: SIAM (Service Integration & Management) · ITIL 4 Service Value
System · Flight Levels · Team-of-Teams / Matrix-Operating-Model.*
