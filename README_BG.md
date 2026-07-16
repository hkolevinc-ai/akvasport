# AkvaSport → Temu scraper

Скрейпърът обхожда категорията **Примамки** в AkvaSport, създава отделен ред за всяка вариация и попълва лист **Template** в оригиналния Temu файл.

## Настройки, заложени в пакета

- Категория: `https://akvasport.com/category/291/primamki.html`
- Тестов лимит: **10 основни продукта**
- Всяка вариация е на отделен ред
- Цена: **EUR**
- Неналична вариация: `Quantity = 0`
- Manufacturer: `Tianyun Fishing Tackle Co.,Ltd`
- EU Responsible person: `AKVASPORT EOOD`
- Shipping Template: `Магазин`
- Handling Time: `1 Day`
- Fulfillment Channel: `I will ship this item myself`
- Не се използват изображения от описанието, `.thumb.webp` или `.box.webp`
- При липса на официална list price се попълва `Not available for List price = N/A`; не се измисля фалшива препоръчителна цена

## Качване в GitHub

1. Създай ново repository.
2. Качи **цялото съдържание** на тази папка, включително скритата папка `.github`.
3. Отвори `Actions` → `AkvaSport scraper` → `Run workflow`.
4. За теста остави `product_limit = 10`.
5. След приключване отвори изпълнението и изтегли артефакта `akvasport-results-...`.

В артефакта ще има:

- `TEMU_AKVASPORT_UPLOAD.xlsx` — готовият Temu файл
- `akvasport_raw_export.csv` — контролен суров експорт
- `scraper.log` — подробен лог

## Пускане за цялата категория

След като тестовият файл бъде проверен, пусни workflow отново с:

```text
product_limit = 0
```

`0` означава всички продукти и всички страници на категорията.

## Важна бележка за наличността

При продукти с вариации сайтът публикува точно количество за всяка вариация и скрейпърът го използва. Ако продуктът няма публикувано числово количество, но е маркиран като наличен, се записва `1`; ако е неналичен — `0`.
