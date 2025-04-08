# Color Patches

Инструменты для подготовки цветопробы. В разработке.

- Имя файла: `elvin_ColorPatches.gms`.
- Автор: **elvin-nsk**.
- Проверенно работает в версии **16**.
- Языки: **без интерфейса**.
- Распространяется **бесплатно**, код **открытый**.
- **Поддерживается автором**.

## Установка

[Стандартная](https://github.com/elvin-nsk/cdr-vba/blob/master/articles/installation.md).

## Использование

#### Функция ProcessPatches

Подписывает цветовые патчи. Для работы необходимы объекты (патчи), представляющие из себя группу из двух объектов: один из них текст, второй - объект, который может иметь заливку (прямоугольник, кривая...). Делаем нужное количество патчей разных цветов. Выделяем их, запускаем функцию. Текст в каждом патче заменяется на рецепт/название цвета заливки второго объекта. Если текст пересекается с объектом, то он перекрашивается в чёрный или белый в зависимости от яркости цвета (чтобы было видно на фоне).

Пример патчей в папке `test`.