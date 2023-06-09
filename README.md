# Percent_discount
Выстраивает ценообразование для клиента

Предпосылки: 
--------
В компании широкая линейка SKU различных категорий, наценка на которые составляет от 1% до трехзначных значений в зависимости от категории. Инструмент общей скидки от базовой цены всегда казался нецелесообразным, пока мы не заключили контракт, где в спецификации прописаны конкрентные позиции, категории, группы товаров с установленным процентом скидки.

Описание: 
--------
Программа берет список артикулов из файла "Заказ от партнера.csv" и проставляет цены с учетом процента скидки от базовой цены.
Проставьте скидки в файле "Спецификация.csv" до запуска программы.
Можно проставить скидки на Актикул, Ассортиментную группу (АГ), Товарную категорию (ТК), Группу (ТГ), Весь рынок (ТР) или общую скидку на все.
Сценарии скидок можно комбинировать. Например, установить скидки только на список Артикулов и Товарный рынок.
Приоритет цены выстраивается от меньшей категории товара к большей.
Если на весь товарный рынок Мебели применена скидка 10%, а на отдельный артикул кресла 15%, применится скидка 15%.
Цены со скидокой вы увидите в файле "Результат.csv".

Функции:
--------
- обновление ценообразования
- автоматическое формирование шаблонов для работы при их отсутствии
- настраиваемая приоритезация, какая скидка должна быть важнее
- маркеры угрозы при снижении цены ниже порогового значения
