"""Testy logiki (bez sięgania po prawdziwą stronę - sandbox nie ma do niej dostępu sieciowego).
Sprawdzają regex do wykrywania linków produktów, wykrywanie dostępności/SKU
oraz budowanie CSV i treści e-maila."""

import csv
import io

import stock_checker as sc

CATEGORY_HTML = """
<div class="product-list">
  <a href="/product-pol-31436-Duzy-skorzany-portfel-meski-w-czarnym-kolorze-model-w-orientacji-poziomej.html">Produkt 1</a>
  <a href="/product-pol-31418-Jasnobrazowy-portfel-meski-ze-skory-naturalnej.html">Produkt 2</a>
  <a href="/product-pol-31436-Duzy-skorzany-portfel-meski-w-czarnym-kolorze-model-w-orientacji-poziomej.html">Duplikat (ta sama karta produktu w dwóch miejscach)</a>
  <a href="/pol_m_Portfele-meskie-1915.html?counter=1">Następna strona</a>
</div>
"""

OUT_OF_STOCK_HTML = """
<html><head><title>Duży, skórzany portfel męski w czarnym kolorze | CedarDrop.com</title>
<script>var dataLayer = [{"sku": "FAKE-FROM-JS-SHOULD-BE-IGNORED"}];</script></head>
<body>
<h1>Duży, skórzany portfel męski w czarnym kolorze, model w orientacji poziomej</h1>
<p>Produkt wyprzedany</p>
<p>Otrzymasz od nas powiadomienie e-mail o ponownej dostępności produktu.</p>
<button>Powiadom mnie o dostępności</button>
<p><strong>SKU</strong>N0035-P-CHM-NL-9495</p>
</body></html>
"""

IN_STOCK_HTML = """
<html><head><title>Klasyczny portfel damski | CedarDrop.com</title></head>
<body>
<h1>Klasyczny, skórzany portfel damski w czerwonym kolorze</h1>
<p>Produkt dostępny w bardzo dużej ilości</p>
<button>Dodaj do koszyka</button>
<p><strong>SKU</strong>ABC-123-XYZ</p>
</body></html>
"""


def test_link_extraction():
    matches = sc.LINK_RE.findall(CATEGORY_HTML)
    ids = {pid for _, pid in matches}
    assert ids == {"31436", "31418"}, f"Nieoczekiwane ID: {ids}"
    print("OK: wykrywanie linków do produktów (z deduplikacją ID) działa")


def test_out_of_stock_and_sku():
    result = sc.parse_product(OUT_OF_STOCK_HTML, "https://cedardrop.com/product-pol-31436-x.html", "31436")
    assert result["in_stock"] is False, "Produkt wyprzedany powinien być wykryty jako NIEDOSTĘPNY"
    assert "portfel" in result["name"].lower()
    assert result["sku"] == "N0035-P-CHM-NL-9495", f"SKU niepoprawnie wykryte: {result['sku']!r}"
    assert "FAKE-FROM-JS" not in result["sku"], "SKU nie powinno pochodzić ze <script> - skrypty muszą być pomijane"
    print(f"OK: 'wyprzedany' -> in_stock={result['in_stock']}, SKU='{result['sku']}' (JS poprawnie zignorowany)")


def test_in_stock_and_sku():
    result = sc.parse_product(IN_STOCK_HTML, "https://cedardrop.com/product-pol-99999-y.html", "99999")
    assert result["in_stock"] is True
    assert result["sku"] == "ABC-123-XYZ", f"SKU niepoprawnie wykryte: {result['sku']!r}"
    print(f"OK: dostępność -> in_stock={result['in_stock']}, SKU='{result['sku']}'")


def test_changes_and_csv():
    old_products = {
        "111": {"name": "Portfel A", "url": "https://cedardrop.com/product-pol-111-a.html", "category": "Portfele męskie", "in_stock": True},
        "222": {"name": "Portfel B", "url": "https://cedardrop.com/product-pol-222-b.html", "category": "Portfele damskie", "in_stock": False},
        "333": {"name": "Portfel C", "url": "https://cedardrop.com/product-pol-333-c.html", "category": "Portfele męskie", "in_stock": True},
    }
    new_results = {
        "111": {"id": "111", "name": "Portfel A", "url": "https://cedardrop.com/product-pol-111-a.html", "category": "Portfele męskie", "in_stock": False, "sku": "A1"},
        "222": {"id": "222", "name": "Portfel B", "url": "https://cedardrop.com/product-pol-222-b.html", "category": "Portfele damskie", "in_stock": True, "sku": "B1"},
        "333": {"id": "333", "name": "Portfel C", "url": "https://cedardrop.com/product-pol-333-c.html", "category": "Portfele męskie", "in_stock": True, "sku": "C1"},
        "444": {"id": "444", "name": "Torebka D", "url": "https://cedardrop.com/product-pol-444-d.html", "category": "Torebki damskie", "in_stock": True, "sku": "D1"},
    }
    changes = sc.compute_changes(old_products, new_results)
    assert changes["111"] == "wypadł z magazynu"
    assert changes["222"] == "wrócił do magazynu"
    assert changes["333"] == ""
    assert changes["444"] == "nowy w ofercie"
    print("OK: compute_changes poprawnie klasyfikuje zmiany (w tym nowy produkt)")

    csv_text = sc.build_csv(new_results, changes)
    reader = list(csv.reader(io.StringIO(csv_text), delimiter=";"))
    header, rows = reader[0], reader[1:]
    assert header == ["Kategoria", "Nazwa produktu", "Status", "Zmiana od ostatniego raportu", "SKU", "Link do produktu"]
    assert len(rows) == 4, f"Oczekiwano 4 wierszy (pełna lista!), jest {len(rows)}"
    # posortowane wg kategorii, potem nazwy -> Portfele damskie przed Portfele męskie przed Torebki damskie
    categories_in_order = [r[0] for r in rows]
    assert categories_in_order == sorted(categories_in_order), "CSV powinien być posortowany wg kategorii"
    names = {r[1] for r in rows}
    assert names == {"Portfel A", "Portfel B", "Portfel C", "Torebka D"}, "CSV musi zawierać WSZYSTKIE produkty, nie tylko zmiany"
    print(f"OK: CSV zawiera pełną listę {len(rows)} produktów, posortowaną wg kategorii, z kolumną Zmiana")
    print("\n--- podgląd CSV ---")
    print(csv_text)

    body = sc.build_summary_body(new_results, changes, errors_count=0, is_first_run=False, total_tracked=4, csv_filename="test.csv")
    assert "Portfele męskie" in body and "Portfele damskie" in body and "Torebki damskie" in body
    assert "test.csv" in body
    print("--- podgląd treści e-maila (podsumowanie) ---")
    print(body)


def test_first_run_body():
    new_results = {"1": {"id": "1", "name": "P1", "url": "u1", "category": "Paski", "in_stock": True, "sku": ""}}
    body = sc.build_summary_body(new_results, {}, errors_count=0, is_first_run=True, total_tracked=1, csv_filename="x.csv")
    assert "pierwsze sprawdzenie" in body.lower()
    print("OK: treść dla pierwszego uruchomienia ma poprawny format")


if __name__ == "__main__":
    test_link_extraction()
    test_out_of_stock_and_sku()
    test_in_stock_and_sku()
    test_changes_and_csv()
    test_first_run_body()
    print("\nWSZYSTKIE TESTY PRZESZŁY POMYŚLNIE")
