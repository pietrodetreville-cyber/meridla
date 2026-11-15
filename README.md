# Fotoapp SVJ Od Vysoké

Appka pro odečty měřidel (telefon) + skript v Google Apps Scriptu (příjem odečtů, tabulka, e-maily).

## Dvě prostředí

| | Ostrá verze | Testovací verze |
|---|---|---|
| Větev v Gitu | `main` | `test` |
| Adresa appky | `https://pietrodetreville-cyber.github.io/meridla/app/` | `https://pietrodetreville-cyber.github.io/meridla/test/app/` |
| Tabulka a skript | ostrá tabulka `mereni_SVJ` | **kopie** tabulky se **vlastním** skriptem |
| Označení | – | oranžový pruh „TESTOVACÍ VERZE", číslo verze s „· TEST", e-maily s předponou `[TEST]` |

Testovací appka nikdy nesahá do ostré tabulky: bez adresy testovacího skriptu nic nevolá, a má oddělené úložiště v telefonu.

## Jak se pracuje

1. Nové věci se vyvíjejí a zkouší na větvi **`test`** (testovací adresa, testovací tabulka).
2. Až je vše ověřeno, změny se **sloučí do `main`** (vydání). Tím se ostrá appka aktualizuje a uživatelům se ukáže hláška „Je k dispozici nová verze X".
3. Skript v Googlu se při vydání **ručně** aktualizuje v ostré tabulce (vložit kód → `setup` → nová verze nasazení). Stejně se postupuje v testovací kopii.

## Zřízení testovací tabulky (jednorázově)

1. V Google Tabulkách: Soubor → Vytvořit kopii (název např. `mereni_SVJ_TEST`). Zkopíruje se i vázaný skript.
2. Rozšíření → Apps Script: v nastavení projektu → Vlastnosti skriptu přidat `ENV` = `test` (e-maily budou mít předponu `[TEST]`), `GEMINI_KEY` (klíč pro čtení z fotek) a případně změnit `NOTIFY` na testovací adresu.
3. Spustit `setup`, potom `resetToken` (zapsat si nový tajný kód) a Implementovat → Nové nasazení → Webová aplikace (Spustit jako: já, Přístup: kdokoli).
4. Adresu `…/exec` vložit do `app/index.html` do `ENDPOINTS.test` (větev `test`).
5. Odkaz pro testování: `https://pietrodetreville-cyber.github.io/meridla/test/app/#k=TESTOVACÍ_KÓD`.
