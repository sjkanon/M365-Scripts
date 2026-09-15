# Werken met bestanden in Teams en SharePoint

**Handleiding voor medewerkers van Petsolutions NV.** Je hoeft niets te installeren en
niets te onthouden — deze uitleg duurt vijf minuten en daarna weet je genoeg.

---

## In het kort

| Vraag | Antwoord |
|---|---|
| Wat verandert er? | Geen mappenbomen meer. Je zet een bestand in het juiste kanaal en beantwoordt twee tot vier vragen |
| Waarom? | Zodat "alle prijslijsten van Butterstone in het Frans" één klik is in plaats van tien mappen doorzoeken |
| Moet ik bestanden verplaatsen? | Nee. Een label geven is genoeg — het bestand blijft staan waar het staat |
| Wat als een bestand bij twee merken hoort? | Kies **Beide**. Eén bestand, en het verschijnt in allebei de merkoverzichten. Geen kopieën meer |
| Kan ik iets stukmaken? | Nee. Een label veranderen raakt de inhoud van het bestand niet aan |

---

## 1. De kern: labels in plaats van mappen

Vroeger bepaalde de map waar iets stond wat het was. Nu bepalen **labels** dat.

Een label (in SharePoint heet dat een kolom) is gewoon een extra kenmerk naast de
bestandsnaam — zoals Merk, Taal of Contenttype. Je ziet ze als kolommen naast elkaar:

```
 Naam                      Merk         Leverancier  Contenttype  Taal     Vertrouwelijkheid
 ────────────────────────────────────────────────────────────────────────────────────────────
 📄 Catalogus 2026.pdf      Butterstone  Lev. 1       Catalogus    NL, FR   Intern
 📄 Prijslijst Q1.xlsx      Beide        Lev. 1       Prijslijst   NL       Deelbaar met klant
 📄 Certificaat CE.pdf      Laseto       Lev. 2       Afb.+cert.   EN       Intern
```

Het grote voordeel: **één bestand kan in meerdere overzichten opduiken.** Een prijslijst
voor beide merken staat één keer opgeslagen en verschijnt zowel bij Butterstone als bij
Laseto. Geen twee kopieën die uit elkaar gaan lopen.

---

## 2. Waar zet ik iets neer?

Elk kanaal in Teams is één pijler. Je kiest het kanaal, verder niets:

| Kanaal | Waarvoor |
|---|---|
| **MGMT** | Directie en beheer (besloten kanaal — alleen leden zien het) |
| **Leveranciers** | Catalogi, prijslijsten en certificaten van leveranciers |
| **Verkopers** | Verkoopmateriaal per regio |
| **Klanten** | Klantdossiers |
| **Marketing** | Marketingmateriaal |
| **TD** | Technische dienst: handleidingen, schema's, fiches |
| **FUTECH Images and videos** | Beeld en video die met klanten gedeeld wordt |

> **Maak geen submappen aan.** Dat hoeft niet meer, en het maakt terugvinden juist
> moeilijker. Gebruik de labels — daar zijn ze voor.

---

## 3. Een bestand toevoegen

Er zijn drie manieren, en ze gedragen zich **niet** hetzelfde. Dit is het belangrijkste
stuk van deze handleiding.

### A. Nieuw aanmaken — de beste manier

In het kanaal → tabblad **Bestanden** → **+ Nieuw**. Je ziet daar alleen het documenttype
van dat kanaal staan, bijvoorbeeld *Leveranciersdocument*. Je krijgt meteen de vragen te
zien en je bent klaar.

### B. Uploaden via de knop — ook goed

**↑ Uploaden** → kies je bestand. SharePoint vraagt daarna om de ontbrekende labels:

```
 Leveranciersdocument
 ─────────────────────────────────────────────
  Naam                Catalogus 2026.pdf
  Merk             *  [ Butterstone        ▾ ]
  Leverancier      *  [ Lev. 1             ⌕ ]   ← typ de eerste letters
  Contenttype      *  [ Catalogus          ▾ ]
  Taal             *  [☑ NL] [☑ FR] [☐ DE] [☐ EN]
  Vertrouwelijkheid*  [ Intern             ▾ ]   ← staat al goed
 ─────────────────────────────────────────────
  Pijler              Leveranciers               ← vult zichzelf in
  Status              Actief                     ← vult zichzelf in
  Deelstatus          Niet gedeeld               ← wordt automatisch bijgehouden
```

De velden met een `*` zijn verplicht. De drie onderste vullen zichzelf in — daar hoef je
niets mee te doen.

### C. Slepen of via de gesynchroniseerde map — let op

Sleep je een hele map bestanden in het venster, of kopieer je ze in Verkenner naar de
gesynchroniseerde Teams-map, dan **vraagt SharePoint niets**. De bestanden komen gewoon
binnen, met lege labels.

Ze zijn niet weg en niemand raakt iets kwijt, maar ze staan er ongelabeld bij en zijn
dus slecht terug te vinden. Je herkent ze aan de melding **"Vereiste info"** in de
bestandslijst.

**Zo ruim je dat op:** ga naar de weergave **Nog te taggen** (bovenaan in het
weergavemenu). Daar staat alles zonder merk bij elkaar. Selecteer meerdere bestanden
tegelijk, klik op het informatiepaneel (ⓘ rechtsboven) en vul de labels in één keer voor
de hele selectie in.

---

## 4. Wat gebeurt er als je een bestand een label geeft?

Dit is de vraag die het vaakst gesteld wordt. Het antwoord is geruststellend:

| | |
|---|---|
| **Het bestand verhuist niet** | Het blijft staan waar het staat. Links die je gedeeld hebt, blijven werken |
| **De inhoud verandert niet** | Een label staat *naast* het bestand, niet erin |
| **Overzichten passen zich meteen aan** | Zet je Merk op Butterstone, dan staat het bestand direct in het Butterstone-overzicht |
| **Kies je Beide** | Dan verschijnt het in het Butterstone- én het Laseto-overzicht. Eén bestand, twee plekken waar je het ziet |
| **Zoeken loopt iets achter** | De overzichten en filters zijn direct bij. De zoekbalk kan een paar minuten tot een uur nodig hebben voor hij het nieuwe label kent |
| **Je kunt het altijd aanpassen** | Verkeerd gelabeld? Verander het gewoon. Er gaat niets verloren |
| **Rechten veranderen niet** | ⚠ Zie hieronder — dit is belangrijk |

> ### ⚠ Een label is géén slot
>
> **Vertrouwelijkheid op "Vertrouwelijk" zetten sluit niemand buiten.** Het is een
> afspraak, geen beveiliging. Wie toegang heeft tot het kanaal, ziet het bestand nog
> steeds.
>
> Wat het label wél doet: er draait elke nacht een controle die kijkt of er bestanden
> met "Intern" of "Vertrouwelijk" extern gedeeld staan. Die komen op een lijstje bij IT.
> Zo merkt iemand het als er iets naar buiten staat dat dat niet hoort.
>
> Wie bij welk kanaal mag, wordt geregeld via groepen. Dat vraag je aan bij IT.

---

## 5. De labels, één voor één

| Label | Wat je invult | Tip |
|---|---|---|
| **Merk** | Butterstone, Laseto of **Beide** | Twijfel je? Beide. Beter te ruim dan een kopie maken |
| **Leverancier** | De leverancier, uit de lijst | Typ de eerste letters, hij vult aan. Staat hij er niet bij? Vraag IT om hem toe te voegen |
| **Contenttype** | Catalogus / Prijslijst / Schrijfrichtlijn / Afbeelding+certificaat / Marketingslag | Dit is *wat voor soort document* het is |
| **Taal** | NL / FR / DE / EN / Geen taal | **Meerdere mag.** Een tweetalige folder krijgt NL én FR. Een foto zonder tekst: Geen taal |
| **Regio** | Benelux / Duitsland / Frankrijk / Export | Alleen in het kanaal Verkopers |
| **Vertrouwelijkheid** | Intern / Deelbaar met klant / Vertrouwelijk | Staat standaard op Intern. Zet hem op "Deelbaar met klant" als een klant het mag krijgen |
| **Status** | Actief / Te archiveren / Verouderd | Staat op Actief. Zet op "Te archiveren" wat weg mag maar nog niet verwijderd |
| **Pijler** | — | Vult zichzelf in op basis van het kanaal |
| **Deelstatus** | — | **Niet invullen.** Dit wordt automatisch bijgehouden, zie punt 7 |

---

## 6. Iets terugvinden

### Weergaven — het snelst

Bovenaan de bestandenlijst staat een menu met weergaven. Per kanaal is er een handige
groepering klaargezet:

| Kanaal | Weergave | Groepeert op |
|---|---|---|
| Leveranciers | Op leverancier | Leverancier |
| Verkopers | Op regio | Regio |
| Klanten | Op vertrouwelijkheid | Vertrouwelijkheid |
| Marketing | Op merk en taal | Merk |
| TD | Op contenttype | Contenttype |

En over alle kanalen heen, in de hoofdbibliotheek:

| Weergave | Wat je ziet |
|---|---|
| **Alles - Butterstone** | Elk bestand van Butterstone (en Beide), uit alle pijlers, als één lijst |
| **Alles - Laseto** | Idem voor Laseto |
| **Nog te taggen** | Alles zonder merk — het opruimlijstje |
| **Extern gedeeld** | Alles wat buiten de organisatie open staat |
| **Te archiveren** | Alles met Status "Te archiveren" of "Verouderd" |

### Filteren — voor een specifieke vraag

Klik rechtsboven op het filterpictogram. Je kunt meerdere labels tegelijk aanvinken:
Merk = Butterstone **én** Contenttype = Prijslijst **én** Taal = FR. Het aantal
resultaten telt live mee.

### Zoeken — als je niet weet waar het staat

De zoekbalk zoekt ook **in** documenten, niet alleen in bestandsnamen. Combineer gerust
met een filter achteraf.

---

## 7. Delen met een klant

Delen doe je zoals altijd: rechtsklik op het bestand → **Delen** of **Koppeling kopiëren**.

Wat er daarna gebeurt: elke nacht kijkt een controle na hoe elk bestand werkelijk gedeeld
staat, en zet dat in de kolom **Deelstatus**:

| Deelstatus | Betekenis |
|---|---|
| Niet gedeeld | Alleen wie bij het kanaal mag, kan erbij |
| Intern gedeeld | Er is een link of een persoon binnen Petsolutions toegevoegd |
| Extern - alleen bekijken | Iemand van buiten kan het bekijken |
| Extern - bewerken | Iemand van buiten kan het wijzigen |

Je vult die kolom dus nooit zelf in — hij vertelt je wat er ís, niet wat jij wil.

**Waar dit op let:** staat er een bestand met Vertrouwelijkheid "Intern" of
"Vertrouwelijk" op "Extern", dan komt dat op een lijstje bij IT. Niemand trekt automatisch
je link in — er wordt naar gekeken en desnoods gebeld.

---

## 8. Veelgestelde vragen

**Ik heb per ongeluk het verkeerde merk gekozen.**
Verander het gewoon. Klik het bestand aan, open het informatiepaneel (ⓘ) en pas het label
aan. De overzichten volgen meteen.

**Moet ik voor elk merk een kopie maken?**
Nee — dat is precies wat we niet meer doen. Kies **Beide**.

**Mag ik nog mappen maken?**
Liever niet. Labels doen hetzelfde werk en werken over kanalen heen; een map niet.

**Een leverancier staat niet in de lijst.**
Vraag IT om hem toe te voegen aan de lijst. Daarna staat hij er voor iedereen bij, en
bestaande bestanden hoeven niet opnieuw gelabeld te worden.

**Ik zie een kanaal niet in Teams.**
Dan zit je niet in de groep die er toegang toe heeft. Vraag dat aan bij IT.

**Ik zie het kanaal wél, maar krijg een foutmelding bij Bestanden.**
Dat kan kloppen: het kanaal is zichtbaar voor het hele team, maar de bestanden zijn
afgeschermd per pijler. Vraag IT om je in de juiste groep te zetten.

**Er staat "Vereiste info" bij mijn bestanden.**
Die zijn binnengesleept of via de synchronisatiemap gekopieerd, en missen labels. Ga naar
de weergave **Nog te taggen** en vul ze bij — dat kan voor meerdere tegelijk.

**Werkt dit ook in de Teams-app op mijn telefoon?**
Bestanden bekijken en delen: ja. Labels invullen gaat het makkelijkst via de browser of
de Teams-desktopapp.

---

## Hulp nodig?

Neem contact op met de IT-servicedesk. Vermeld het kanaal en de bestandsnaam — dan is het
meteen duidelijk waar het over gaat.
