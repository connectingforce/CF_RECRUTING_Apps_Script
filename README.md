# CF_RECRUTING — Google Apps Script

Synchronizacja zgłoszeń do istniejącego arkusza CF_RECRUTING (zakładka `KANDYDACI`).

## Bezpieczeństwo
- Kod nie zawiera identyfikatorów prywatnych arkuszy, kandydatów ani danych osobowych.
- Repozytorium powinno być **Private**.
- Najpierw uruchom `previewRecruitingSync()`: nie zapisuje niczego.
- Nie uruchamiaj automatycznego importu przed sprawdzeniem wyniku podglądu i podstawy prawnej wykorzystania zgłoszeń pochodzących od odrębnego administratora danych.
- Import dopisuje tylko nowe wiersze. Istniejących statusów, notatek i terminów kontaktu nie modyfikuje.

## Instalacja
1. Otwórz docelowy arkusz `CF_RECRUTING` → Rozszerzenia → Apps Script.
2. Utwórz plik `RecruitingSync.gs` i wklej **całą** zawartość `src/RecruitingSync.gs`.
3. W Apps Script → Ustawienia projektu → Właściwości skryptu dodaj:
   - `CF_RECRUTING_TARGET_ID` — identyfikator arkusza docelowego,
   - `CF_RECRUTING_SOURCE_NEODAK_ID` — identyfikator arkusza źródłowego.
4. Wykonaj `previewRecruitingSync()` i sprawdź Dziennik wykonywania.
5. Po weryfikacji zgodności zgód, liczby rekordów, mapowania i braku duplikatów uruchom `importRecruitingBatch()`.
6. Po testach produkcyjnych można włączyć `installRecruitingTrigger()` (co 15 minut).

## Zakres v0.1
- Źródło: zakładka `Roofers` w arkuszu NEODAK, typ ROOFER.
- Dedup: `Submission ID` plus e-mail oraz nazwisko+telefon dla rekordów historycznych.
- Kolumna AS (45) w docelowym arkuszu: techniczny `SOURCE_KEY` dla kolejnych synchronizacji.
- Statusy rekruterskie w już istniejących rekordach pozostają nietknięte.
- Nowy rekord: status NEW i źródło NEODAK_RecrutingSources.
- Import porcjami do 50 nowych zgłoszeń na wykonanie, od najnowszych do najstarszych.
- Nowe źródła można dopisywać w `RECRUITING_SOURCES` po dokładnym rozpoznaniu ich nagłówków.

**Ważne:** Podgląd nie jest testem z prawdziwym zapisem. Automatyczny trigger wymaga wdrożenia i autoryzacji w Apps Script na koncie mającym dostęp do obu arkuszy.
