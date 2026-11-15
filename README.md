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

1. V **ostré** tabulce: menu **Měřidla → Vytvořit testovací sadu** (funkce `createTestSet`). V `SVJ_app/TESTOVACI` vznikne kopie tabulky `mereni_SVJ_TEST` a obě složky pro fotky.
2. V kopii: Rozšíření → Apps Script (pokud se skript nezkopíroval, vložit kód z `apps-script/Code.gs`). V nastavení projektu → Vlastnosti skriptu přidat `ENV` = `test`, `GEMINI_KEY` a případně změnit `NOTIFY` na testovací adresu.
3. Spustit `setup` (použije složky v TESTOVACI), potom `clearTestData` (smaže zkopírované odečty a tajné kódy, běží jen s `ENV = test`) a `resetToken` (zapsat si nový tajný kód).
4. Implementovat → Nové nasazení → Webová aplikace (Spustit jako: já, Přístup: kdokoli).
5. Adresu `…/exec` vložit do `app/index.html` do `ENDPOINTS.test` (větev `test`).
6. Odkaz pro testování: `https://pietrodetreville-cyber.github.io/meridla/test/app/#k=TESTOVACÍ_KÓD`.

## Uspořádání na Google Disku

```
SVJ_app/
  OSTRE/        tabulka mereni_SVJ, Fotodokumentace odečtů, Místa měřidel – plánky a fotky
  TESTOVACI/    tabulka mereni_SVJ_TEST, Fotodokumentace odečtů, Místa měřidel – plánky a fotky
```

Soubory a složky se uspořádají jedním spuštěním (menu tabulky **Měřidla → Uspořádat Disk (SVJ_app)**, nebo funkce `organizeDrive`).
Testovací kopie skriptu musí mít předtím ve vlastnostech skriptu `ENV` = `test`, aby se její složky založily v podsložce `TESTOVACI`.
Přesun nic nerozbije: odkazy a identifikátory souborů zůstávají stejné.

## Nasazení skriptu do Googlu (clasp)

Kód skriptu se do Google Apps Scriptu nahrává příkazem, ne ručním vkládáním:

```
scripts/deploy-script.sh test     # testovací tabulka
scripts/deploy-script.sh prod     # ostrá tabulka (až bude nastaven .clasp.prod.json)
```

Příkaz nahraje `apps-script/Code.gs` a `appsscript.json`, vytvoří novou verzi a přenasadí **stejnou** webovou aplikaci (adresa `…/exec` se nemění).
Jaká verze skriptu v Googlu běží, ukáže `ADRESA_NASAZENÍ/exec?action=version` (např. `{"version":"1.18.1","env":"test"}`).

Jednorázově: nainstalovat clasp (`npm i -g @google/clasp`), `clasp login` (účet vlastníka skriptů) a zapnout Apps Script API na https://script.google.com/home/usersettings.
Přihlašovací údaje clasp jsou jen v počítači (`~/.clasprc.json`) a do repozitáře nepatří.
Vlastnosti skriptu (`ADMIN_PIN`, `GEMINI_KEY`, `ENV`, …) se nenahrávají, zůstávají v Googlu.
První povolení nových oprávnění (nová služba Google) musí člověk potvrdit v editoru skriptu.
Funkce typu `setup` se spouští ručně v editoru.
