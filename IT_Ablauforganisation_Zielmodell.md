# Zielmodell IT-Ablauforganisation — „Homebase & Mannschaft"

> Betriebs- und Zusammenarbeitsmodell für die neue IT-/IPS-Organisation.
> Leitidee: **Die Org-Einheit ist die Homebase der Mitarbeitenden. Gespielt wird
> in Mannschaften über die Einheiten hinweg.**

Begleitende Deliverables:
- **Interaktive Fassung (Web):** siehe Artifact-Link in der Konversation
- **CIO-Kurzfassung (PPTX):** `2026_07_08_IT_Ablauforganisation_Homebase_Mannschaft.pptx`
  (erzeugt via `build_org_operating_model.py`)

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

*Referenzrahmen: SIAM (Service Integration & Management) · ITIL 4 Service Value
System · Flight Levels · Team-of-Teams / Matrix-Operating-Model.*
