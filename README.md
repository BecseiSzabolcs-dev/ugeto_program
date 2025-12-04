#hello world

Ez a Python alkalmazás egy grafikus felhasználói felülettel (GUI) rendelkező eszköz, amelyet a ló- vagy agárversenyek (a kódban szereplő logikák alapján valószínűleg ügetőversenyek) hivatalos PDF-programjainak feldolgozására, a kivonatolt adatok szerkesztésére, valamint CSV fájlok és PowerPoint prezentációk generálására terveztek.

⚙️ Funkciók és Képességek
PDF Betöltése: Képes betölteni egy PDF fájlt, és automatikusan kivonatolni belőle a versennyel kapcsolatos adatokat (pl. futamok címei, lovasok/hajtók, lovak, időpontok, vélemények).

Adatszerkesztő Felület (GUI): A kivonatolt adatok két fülön szerkeszthetők:

Titles (Címek): Itt találhatók a futamokra vonatkozó fő adatok (pl. Azonosító, Cím, Táv, Időpont, Start típusa, Vélemény).

Drivers (Hajtók): Itt találhatók a futam résztvevőinek adatai (pl. Lószám, Lónév, Lótáv, Hajtó neve, Futam száma, Futott-e státusz).

Adatszerkesztés: Lehetőség van a cellák tartalmának közvetlen módosítására, sorok hozzáadására és törlésére.

Keresés: Mindkét táblában van keresési funkció az adatok gyors szűrésére.

Adatok Mentése CSV-be: A szerkesztett adatok exportálhatók két külön CSV fájlba (titles_data.csv és drivers_data.csv), amelyek a csv/ mappában jönnek létre.

PPT Készítése: Egyéni formázású PowerPoint prezentációkat generál a szerkesztett adatokból (futamonként egy .pptx fájlt) a ppt/ mappában. A prezentációk tartalmazzák a futam adatait, a résztvevőket, és diát az eredményeknek/osztalékoknak.
