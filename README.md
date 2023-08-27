# Przeszukiwarka PDF i Outlook

## Opis
Program umożliwia przeszukiwanie plików PDF oraz wiadomości e-mail w Outlooku w celu znalezienia określonych fraz. Użytkownik ma możliwość określenia kilku parametrów wyszukiwania, takich jak: ścieżka do folderu źródłowego, ścieżka do folderu docelowego, szukane frazy, operator (AND, OR, NOT), czy wyszukiwanie ma być rekurencyjne oraz czy ma uwzględniać wielkość liter. Użytkownik może również określić, czy chce przeszukiwać wiadomości e-mail w określonym folderze w Outlooku.

## Instalacja
1. Upewnij się, że masz zainstalowaną odpowiednią wersję Pythona (zalecana wersja: 3.8+).
2. Sklonuj repozytorium na swój lokalny komputer.
3. Przejdź do folderu z projektem.
4. Zainstaluj wymagane biblioteki używając poniższej komendy:

pip install -r requirements.txt


## Użycie
1. Uruchom program poprzez komendę:
python nazwa_pliku.py

2. Wypełnij wszystkie wymagane pola w interfejsie użytkownika.
3. Kliknij przycisk "Szukaj".

## Potencjalne zastosowania
- Automatyczne sortowanie dokumentów PDF na podstawie zawartości.
- Szybkie wyszukiwanie ważnych wiadomości e-mail w Outlooku.
- Archiwizacja dokumentów na podstawie określonych kryteriów.

## Kierunki rozwoju
- Dodanie wsparcia dla innych formatów plików (np. DOCX, XLSX).
- Możliwość wyszukiwania w innych klientach pocztowych.
- Rozszerzenie funkcjonalności interfejsu użytkownika, np. dodanie funkcji podglądu znalezionych plików.
- Optymalizacja i poprawa wydajności wyszukiwania, zwłaszcza dla dużych zbiorów danych.

