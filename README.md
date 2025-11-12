# Analiza terminowości dostaw w Excel VBA

## 1. Cel projektu
Celem projektu było opracowanie zautomatyzowanego narzędzia analitycznego do oceny terminowości dostaw i czasu trwania tras logistycznych.  
Projekt powstał w środowisku **Excel VBA** i umożliwia pełną automatyzację procesu raportowania danych operacyjnych.  
Dzięki wdrożeniu rozwiązania czas przygotowania raportu został skrócony z kilku/kilkunastu godzin do kilku sekund, przy jednoczesnym wyeliminowaniu błędów manualnych.

---

## 2. Zakres danych
Dane źródłowe pochodzą z systemu klasy TMS (Transport Management System) i obejmują informacje dotyczące realizacji transportów.  
Poniżej przedstawiono przykładową strukturę danych wykorzystywaną w analizie:

| Pole | Opis |
|------|------|
| KOD_TRASY | Unikalny identyfikator transportu (np. 734550910684251) |
| NUMER_ZAMOWIENIA | Numer referencyjny zlecenia transportowego |
| DATA_DOSTAWY | Data planowana lub faktyczna realizacji dostawy |
| GODZINA_OD / GODZINA_DO | Zakres czasowy okna dostawy |
| PUNKT_ROZLADUNKU | Nazwa lub kod lokalizacji docelowej |
| STATUS_TRANSPORTU | Status realizacji (np. zakończony, w toku, anulowany) |

Analiza obejmuje wyznaczenie różnic między zaplanowanymi a faktycznymi godzinami dostaw oraz identyfikację przekroczeń progów SLA.

---

## 3. Architektura rozwiązania

Projekt składa się z kilku modułów VBA, zorganizowanych w logiczną strukturę przetwarzania danych:

### 3.1. Moduł przygotowania danych
- Walidacja i czyszczenie danych wejściowych.  
- Identyfikacja unikalnych tras na podstawie pola `KOD_TRASY`.  
- Przypisanie numerów oddziałów na podstawie fragmentu kodu trasy.

### 3.2. Moduł analityczny
- Obliczanie różnicy czasowej pomiędzy pierwszym i ostatnim rozładunkiem w trasie.  
- Sprawdzenie, czy różnica przekracza ustalony próg (np. 3 godziny).  
- Klasyfikacja trasy jako „TAK” (przekroczenie) lub „NIE” (w normie).

### 3.3. Moduł raportowy
- Automatyczne generowanie arkuszy szczegółowych dla tras z przekroczeniami.  
- Tworzenie zestawienia zbiorczego z wykorzystaniem tabeli przestawnej.  
- Dodanie hiperłączy prowadzących do arkuszy szczegółowych.  
- Formatowanie raportu zgodnie z zasadami czytelności i przejrzystości raportów operacyjnych.

### 3.4. Moduł czyszczenia
- Usuwanie arkuszy tymczasowych.  
- Reset wartości w panelu sterowania.  
- Przygotowanie skoroszytu do kolejnej sesji analitycznej.

---

## 4. Panel użytkownika
Projekt zawiera arkusz „Panel Sterowania”, który pełni funkcję prostego interfejsu użytkownika.  
Użytkownik może:
- określić próg dopuszczalnego czasu realizacji (w godzinach),
- wybrać zakres analizowanych danych,
- uruchomić pełną analizę jednym przyciskiem „START”.

Cały proces nie wymaga znajomości VBA i może być wykonywany przez osoby nietechniczne.

<img width="2553" height="1388" alt="image" src="https://github.com/user-attachments/assets/4b497b42-8e5d-440a-a14e-39798d30e8fd" />

---

## 5. Przykładowe zastosowania

1. Analiza operacyjna tras w celu wykrycia opóźnień i wąskich gardeł procesowych.  
2. Raportowanie SLA i wskaźników terminowości w ujęciu tygodniowym lub miesięcznym.  
3. Analiza trendów i porównanie efektywności operacyjnej między oddziałami.  
4. Weryfikacja kompletności i jakości danych z systemów TMS/ERP.  

---

## 6. Wyniki i efekty

| Wskaźnik | Przed wdrożeniem | Po wdrożeniu |
|-----------|------------------|--------------|
| Czas przygotowania raportu | 3–4 godziny dziennie | poniżej 10 sekund |
| Błędy manualne | częste | wyeliminowane |
| Powtarzalność wyników | zależna od użytkownika | w pełni standaryzowana |
| Dostępność danych szczegółowych | ograniczona | pełna, z hiperłączami |

---

## 7. Technologie i narzędzia

| Obszar | Technologia |
|--------|--------------|
| Automatyzacja | Excel VBA (Visual Basic for Applications) |
| Analiza danych | Tabele przestawne, dynamiczne zakresy, formatowanie warunkowe |
| Źródło danych | Eksport SQL z systemu klasy TMS |
| Raportowanie | Automatyczne arkusze i dashboardy w Excelu |
| Interfejs użytkownika | Panel sterowania z przyciskiem „START” |

---

## 8. Wnioski

Projekt stanowi przykład skutecznej automatyzacji analiz operacyjnych w środowisku Excel.  
Zastosowanie VBA pozwoliło całkowicie wyeliminować manualne działania analityków i uzyskać powtarzalne wyniki w ustandaryzowanym formacie raportowym.  
Rozwiązanie może być wykorzystane w analizie terminowości dostaw, raportowaniu SLA, ocenie jakości procesów logistycznych oraz monitorowaniu wskaźników KPI.

<img width="2552" height="1388" alt="image" src="https://github.com/user-attachments/assets/ea52ff4b-51b2-4559-8ae5-a8ba5b0b8af2" />

<img width="2549" height="1388" alt="image" src="https://github.com/user-attachments/assets/6a2a4f21-4a48-42ab-b0ce-fbbb0423eaef" />




