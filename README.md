# Bot do sprawdzania stanów magazynowych CedarDrop.com

Sprawdza dostępność **wszystkich** produktów w sklepie CedarDrop.com dwa razy
dziennie i wysyła e-mail z **pełną listą produktów w załączniku CSV**,
podzieloną na kategorie (portfele, torebki, plecaki...), plus krótkie
podsumowanie zmian w treści wiadomości.

## Zanim zaczniesz - rozważ lepszą opcję

CedarDrop.com to platforma dropshippingowa (IdoSell) skierowana do firm.
Prawie każda polska hurtownia dropshippingowa udostępnia partnerom
**oficjalny plik/feed XML lub CSV ze stanami magazynowymi** (czasem też API) -
to szybsze, dokładniejsze i nie obciąża ich serwera tysiącami zapytań co 12h.

Warto zapytać obsługę CedarDrop (**drop@cedardrop.com**, +48 607 900 998) albo
swojego opiekuna handlowego, czy taki plik jest dostępny - to prawdopodobnie
lepsze rozwiązanie niż ten bot.

Ten skrypt to "plan B": odwiedza stronę każdego produktu, bo strony kategorii
same w sobie nie pokazują dostępności - tylko status "cena po zalogowaniu".

## Jak to działa

1. Wchodzi na strony wszystkich głównych kategorii i zbiera linki do produktów,
   zapamiętując, z jakiej kategorii pochodzi każdy z nich.
2. Odwiedza stronę każdego produktu i sprawdza, czy jest dostępny (status
   "dostępny / wyprzedany" widać bez logowania, w przeciwieństwie do cen),
   oraz próbuje odczytać SKU.
3. Porównuje wynik z poprzednim sprawdzeniem (zapisanym w `state.json`), żeby
   oznaczyć, co się zmieniło.
4. Wysyła e-mail: w treści krótkie podsumowanie (ile dostępnych/niedostępnych
   w każdej kategorii + co się zmieniło), a w załączniku CSV **pełną listę
   wszystkich sprawdzonych produktów**.
5. Zapisuje nowy stan z powrotem do repozytorium na GitHub.

Sklep ma tysiące produktów (sama kategoria portfeli męskich to 712 sztuk),
więc **jedno sprawdzenie może trwać nawet około godziny** - to normalne.

### Kolumny w załączniku CSV

| Kolumna | Opis |
|---|---|
| Kategoria | np. "Portfele męskie", "Torebki damskie" - plik jest posortowany wg tej kolumny, więc produkty z tej samej kategorii są razem |
| Nazwa produktu | pełna nazwa ze strony |
| Status | Dostępny / Wyprzedany |
| Zmiana od ostatniego raportu | puste = bez zmian, albo "wypadł z magazynu" / "wrócił do magazynu" / "nowy w ofercie" |
| SKU | najlepsza możliwa próba odczytu (może być puste - patrz uwaga niżej) |
| Link do produktu | bezpośredni link, żeby łatwo znaleźć produkt na stronie |

Plik CSV używa średnika (`;`) jako separatora i kodowania UTF-8 z BOM, żeby
Excel poprawnie pokazywał polskie znaki od razu po otwarciu.

## Konfiguracja (ok. 15 minut)

### 1. Załóż repozytorium na GitHub

Jeśli nie masz konta - załóż darmowe na [github.com](https://github.com).
Stwórz nowe **publiczne** repozytorium (np. `cedardrop-stock-bot`) i wgraj
do niego wszystkie pliki z tego folderu, **zachowując strukturę katalogów**
(łącznie z `.github/workflows/stock_check.yml`).

Najprościej: przeciągnij cały folder na stronę
`github.com/nowe-repo/upload` albo użyj `git`:
```
cd cedardrop-stock-bot
git init
git add .
git commit -m "Pierwsza wersja bota"
git branch -M main
git remote add origin https://github.com/TWOJA-NAZWA/cedardrop-stock-bot.git
git push -u origin main
```

**Dlaczego publiczne repozytorium?** GitHub Actions dla prywatnych repo ma
darmowy limit 2000 minut miesięcznie. Sprawdzanie całego sklepu 2x dziennie
może go przekroczyć. Publiczne repozytoria mają nielimitowane darmowe minuty.
`state.json` nie zawiera żadnych wrażliwych danych - tylko nazwy produktów
i status dostępności. Hasło do e-maila jest bezpieczne osobno (patrz krok 3).

### 2. Przygotuj konto e-mail do wysyłki (Gmail)

1. Włącz weryfikację dwuetapową na koncie Gmail, z którego bot ma wysyłać:
   https://myaccount.google.com/security
2. Wygeneruj "hasło aplikacji": https://myaccount.google.com/apppasswords
   (wybierz "Inna (niestandardowa nazwa)", wpisz np. "CedarDrop bot")
3. Zapisz to 16-znakowe hasło - będzie potrzebne w kroku 3.

(Możesz też użyć innej skrzynki niż Gmail - zobacz sekcję niżej.)

### 3. Dodaj sekrety w repozytorium

W repo na GitHub: **Settings → Secrets and variables → Actions →
New repository secret**. Dodaj trzy sekrety:

| Nazwa sekretu    | Wartość                                          |
|------------------|---------------------------------------------------|
| `SMTP_USER`      | Twój adres Gmail (np. `jan@gmail.com`)             |
| `SMTP_PASSWORD`  | Hasło aplikacji wygenerowane w kroku 2 (16 znaków) |
| `EMAIL_TO`       | Adres, na który mają przychodzić raporty           |

### 4. Przetestuj na małej próbce

Dodaj czwarty, tymczasowy sekret: `MAX_PRODUCTS` o wartości np. `20` -
dzięki temu pierwszy test sprawdzi tylko 20 produktów zamiast całego sklepu
i szybko dowiesz się, czy wszystko działa.

Zakładka **Actions** w repo → workflow "Sprawdzanie stanow magazynowych
CedarDrop" → przycisk **Run workflow** (uruchamia od razu, nie trzeba czekać
na harmonogram). Sprawdź log i skrzynkę e-mail.

Gdy test się powiedzie - **usuń sekret `MAX_PRODUCTS`**, żeby kolejne
uruchomienia obejmowały cały sklep.

### 5. Gotowe

Bot będzie teraz uruchamiał się automatycznie o **6:00 i 18:00 UTC**
(ok. 8:00 i 20:00 czasu polskiego latem, 7:00 i 19:00 zimą - Polska zmienia
czas, a harmonogram GitHub Actions działa w UTC).

Żeby zmienić godziny, edytuj linie `cron` w pliku
`.github/workflows/stock_check.yml` (format: minuta godzina dzień miesiąc
dzień-tygodnia, zawsze w UTC).

## Pierwszy e-mail będzie inny

Przy pierwszym uruchomieniu bot nie ma z czym porównać wyników, więc dostaniesz
pełną listę CSV, ale kolumna "Zmiana" będzie pusta dla wszystkich produktów
(brak punktu odniesienia). Dopiero **kolejne** raporty pokażą w tej kolumnie
realne zmiany (co zniknęło / wróciło / jest nowe).

## Inny dostawca poczty (nie Gmail)

Dodaj dodatkowo sekrety `SMTP_HOST` i `SMTP_PORT` z danymi swojego dostawcy
(np. skrzynki firmowej). Zwykle port to `587` (STARTTLS) - dokładne dane
znajdziesz w ustawieniach swojej poczty lub u dostawcy.

## Testowanie logiki bez zużywania minut Actions

Plik `test_logic.py` sprawdza wykrywanie linków i statusu "dostępny/wyprzedany"
na przykładowym, spreparowanym HTML - możesz go uruchomić lokalnie
(`pip install -r requirements.txt && python3 test_logic.py`), żeby upewnić
się, że logika działa, zanim zlecisz botowi pełny przebieg.

## Uwaga o dokładności wykrywania

Status "dostępny/wyprzedany" jest wykrywany na podstawie tekstu strony
produktu (stała `OUT_OF_STOCK_MARKERS` w `stock_checker.py`). Jeśli CedarDrop
zmieni szatę graficzną strony albo użyje innych sformułowań dla części
produktów, wykrywanie może wymagać drobnej poprawki w tym miejscu.

Kolumna SKU to również "najlepsza możliwa próba" (regex szukający etykiety
"SKU" w tekście strony) - dla części produktów może zostać pusta, jeśli
CedarDrop użyje innego układu strony niż w produkcie, na którym to testowałem.
Kolumna Link zawsze jest wiarygodna i wystarczy do znalezienia produktu.
