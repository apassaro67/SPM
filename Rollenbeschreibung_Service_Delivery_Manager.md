# Rollenbeschreibung — Service Delivery Manager (SDM)

**Organisation:** IPS (Integrated Processes & Solutions)
**Einheit / Heimat:** Service Integration & Operations (SIAM) — IPS Management Office
**Berichtslinie:** Head of Service Delivery → SIAM / Service Integration Lead
**Spannweite:** 1 SDM je Service-Cluster (Bündel fachlich verwandter Services);
mehrere SDMs bilden gemeinsam einen **Pool**

---

## Mission

Der Service Delivery Manager sichert den **durchgängigen, wirtschaftlichen und
qualitätsgerechten Betrieb** eines Service-Clusters. Er/sie **koordiniert alle
beteiligten internen Support-Einheiten und externen Dienstleister end-to-end**
und **entlastet den Service Owner vom operativen „Wie"**. Der SDM ist der
**einzige Koordinationspunkt** für den Betrieb der zugeordneten Services.

> Kurzformel: Der **Service Owner** verantwortet das **WAS** (Ergebnis), der
> **Service Delivery Manager** das **WIE** (Koordination).

---

## Organisatorische Einordnung

| Aspekt | Beschreibung |
|---|---|
| Heimat | SIAM / Service Integration & Operations (IPS Management Office) |
| Berichtet an | SIAM / Service Integration Lead |
| Enge Zusammenarbeit | Service Owner (BSS / Fach-Einheit), Practice Supplier Management, TSS-Operations, zentraler Service Desk, Regional Support Center (RSC), AMS-/Ops-Dienstleister (z. B. DXC) |
| Führungscharakter | **Laterale Führung** — fachliche Koordination ohne disziplinarische Weisungsbefugnis |

---

## Pool-Modell & Verortung — einer oder mehrere SDMs?

Es gibt **nicht einen SDM, sondern einen Pool**. Ein einzelner SDM kann nicht
alle Services koordinieren; skaliert wird über **Service-Cluster** (je ein SDM,
große Cluster ggf. ein kleines Team). Ein **Head of Service Delivery** führt den
Pool und berichtet an den SIAM Lead.

**Prinzip: zentral verankert, dezentral eingesetzt** (das Homebase-/Mannschaft-
Prinzip, angewandt auf die Rolle selbst):

- **Heimat / Führung (Solid-Line):** alle SDMs gehören zu SIAM im IPS Management
  Office und werden vom Head of Service Delivery geführt.
- **Fachlicher Einsatz (Dotted-Line):** jeder SDM ist einem Service-Cluster
  zugeordnet und arbeitet dort eng mit Service Owner(n) und Liefereinheiten.

**Illustrative Cluster-Zuordnung** (Schnitt entlang der Value Streams):

| Service-Cluster | SDM-Bedarf (Richtwert) |
|---|---|
| SAP & Business Applications (DXC-lastig) | 1 SDM, ggf. 2 (großes AMS-Volumen) |
| Modern Workplace & Collaboration | 1 SDM |
| Network & Connectivity | 1 SDM |
| Compute · Cloud · Identity & Access | 1 SDM (ggf. IAM geteilt) |
| Service Desk / ITSM (querschnittlich) | 1 SDM |
| Data, Analytics & KI | mit SAP/BSS-Cluster geteilt |

**Sizing-Regel:** nicht nach Anzahl Services, sondern nach **Koordinationslast**
(Zahl beteiligter Provider + interner Einheiten, Kritikalität, Ticketvolumen).
Faustwert: ein SDM steuert eine große oder eine Handvoll mittlerer Services.

### Warum die SDMs NICHT in die Liefereinheiten verteilt werden sollten

Entscheidend ist, was „verteilt" meint: verteilter **Sitz/Einsatz** ist gut —
verteilte **Berichtslinie** in BSS/TSS/Cyber ist riskant.

| Zentraler Pool (IPS MO) — empfohlen | SDMs berichten IN die Einheiten — Risiko |
|---|---|
| Neutrales Mandat über BSS / TSS / Cyber | Peer koordiniert Peer → schwaches Mandat |
| Ein Standard: Methode, KPIs, ServiceNow | „Acht Dialekte": Methode/KPI/Tool driften |
| Gegenseitige Vertretung & Karrierepfad | Loyalitätskonflikt zur eigenen Einheit |
| Head of SDM balanciert Kapazität | Verantwortung ohne Autorität beim Head of SDM |
| Kein Rückfall in „pro-Service-Owner"-Silos | Faktischer Rückfall zum heutigen Status quo |

### Kann der Head of SDM zentral sitzen und die SDMs verteilt?

**Ja — unter Bedingungen.** Tragfähig, solange der Head of Service Delivery eine
fachliche **Solid-Line** zu allen SDMs hat und Methode, KPIs, Tooling sowie das
einheitenübergreifende Mandat zentral bleiben. Dann ist es faktisch der zentrale
Pool mit dezentralem Einsatz — nicht „verteilt" im problematischen Sinn.

- **Zielbild:** zentraler SDM-Pool im IPS MO.
- **Übergang / politischer Kompromiss:** Head of SDM zentral, SDMs co-located in
  den Clustern, aber mit **Solid-Line zum Head of SDM** — als Phase-3-Etappe der
  Migration. Entscheidend ist die Berichtslinie, nicht der Schreibtisch.

---

## Kernaufgaben (Responsibilities)

1. **End-to-End-Koordination** aller Support-Ebenen des Service-Clusters —
   L1 (Service Desk), L2/L3 (intern & Provider), Onsite/RSC.
2. **Steuerung der AMS-/Ops-Dienstleister** im Tagesgeschäft: Ticketfluss,
   Einhaltung von UC/SLA, Eskalationsmanagement — auf Basis der von Practice
   Supplier Management gesetzten Verträge und KPIs.
3. **Sicherstellung des End-to-End-SLA** über die gesamte Service-Kette; Konsistenz
   von SLA ↔ OLA ↔ UC (back-to-back).
4. **Betrieb der operativen Rituale:** Service Operations Review (wöchentlich,
   SDM-geführt), Provider Jour Fixe, Beitrag zum Supplier & Operations Board.
5. **Incident-/Problem-/Change-Koordination** über Einheiten und Provider hinweg;
   aktive Rolle im Major-Incident-Prozess.
6. **Schnittstellenmanagement** BSS ↔ TSS ↔ AMS ↔ Service Desk — Übergabepunkte
   (OLA) klären, „Zwischen-den-Stühlen"-Tickets vermeiden.
7. **Transparenz herstellen:** konsolidiertes Reporting zu Service-Health, SLA,
   Provider-Performance und Kosten/Verbrauch in ServiceNow.
8. **Continuous Improvement** des Betriebsmodells je Cluster.

---

## Abgrenzung — was der SDM *nicht* verantwortet

| Nicht Aufgabe des SDM | Verantwortlich ist |
|---|---|
| Fachliche Service-Definition, Priorisierung, Business-Wert | Service Owner |
| Vertragshoheit, kommerzielle Verhandlung, Lieferantenauswahl | Practice Supplier Management / Einkauf |
| Technische Betriebsausführung | TSS-Operations-Teams / Provider |
| Disziplinarische Führung der Support-Mitarbeitenden | Linienorganisation (Homebase) |

---

## Mandat / Befugnisse

- **Fachliches Weisungsrecht** gegenüber den beteiligten Support-Einheiten und
  Providern im Kontext der zugeordneten Services.
- **Eskalationsrecht** bei Gefährdung des End-to-End-SLA.
- **Auslöserecht** für den Major-Incident-Prozess.
- **Priorisierung** operativer Tickets im Rahmen der Vorgaben des Service Owners.

---

## Erfolgsmessung (KPIs)

- End-to-End-SLA-Erfüllung je Service
- MTTR (Mean Time to Restore) · Eskalationsquote
- Provider-SLA-/UC-Einhaltung
- Erstlösungsquote (First Time Fix)
- Anwenderzufriedenheit (CSAT)
- Kosten/Service bzw. Betriebskosten-Transparenz

---

## Anforderungsprofil

**Fachlich**
- ITIL 4 / ITSM, insb. Service Integration, Incident-, Problem- und
  Change-Management; SIAM-Verständnis
- Provider-/Vendor-Management-Erfahrung; Vertrags- und SLA-Kompetenz
- Erfahrung mit Application Management Services (z. B. SAP / DXC) und
  Infrastruktur-Betrieb
- Sicherer Umgang mit ServiceNow (ITSM, Reporting/Dashboards)

**Persönlich**
- Ausgeprägte Koordinations-, Kommunikations- und Eskalationsfähigkeit
- Durchsetzungsstärke in **lateraler Führung** (ohne disziplinarische Macht)
- Analytisch, datengetrieben, service- und kundenorientiert
- Sprachen: Deutsch und Englisch (verhandlungssicher)

---

## Schnittstellen — RACI (Kurzform)

Legende: **R** Responsible · **A** Accountable · **C** Consulted · **I** Informed

| Aktivität | Service Owner | **SDM** | Practice Supplier Mgmt | SIAM Lead | Provider / Einheit |
|---|:--:|:--:|:--:|:--:|:--:|
| Service-Priorität & Ziele setzen | **A** | C | I | I | I |
| Operative Koordination der Kette | I | **A/R** | I | C | R |
| Provider-Vertrag & KPI-Rahmen | I | C | **A** | C | I |
| End-to-End-SLA sicherstellen | C | **A** | C | R | R |
| Major Incident (P1/P2) | I | **R** | I | A | R |
| Betriebsmodell verbessern | C | **R** | C | A | C |
