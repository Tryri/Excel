import csv
import openpyxl
import re

filename = 'Без пустых долей и помещений.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
# row = sheet['B2'].value
# a = re.split(r' ул', r' пер',  row)
# print(a)
splits = (' ул', ' пер')
for row in sheet['A']:
    a = row.value
    # a = re.sub(' ул', '', a)
    a = re.sub(r'\bул', 'ул,', a)
    # a = ' '.join(a)
    a = re.sub(r'пер ', 'пер, ', a)
    # a = ' '.join(a)
    a = re.sub(r'шоссе', 'ш,', a)
    # a = ' '.join(a)
    a = re.sub(r'бульвар', 'б-р,', a)
    # a = ' '.join(a)
    a = re.sub(r'площадь', 'пл,', a)
    # a = ' '.join(a)
    a = re.sub(r'проезд', 'проезд,', a)
    # a = re.sub(r'набережная', 'набережная,', a)
    a = re.sub(r'поселок', 'п,', a)
    a = re.sub(r'городок', 'городок,', a)
    # a = ' '.join(a)
    # a = re.sub(r'\d корп.\d', '', a)
    a = re.sub(r'1-я Курская', 'Курская 1-я,', a)
    a = re.sub(r'1-я Посадская', 'Посадская 1-я,', a)
    a = re.sub(r'1-я Пушкарная', 'Пушкарная 1-я,', a)
    a = re.sub(r'2-я  Курская', 'Курская 2-я, ', a)
    a = re.sub(r'2-я Курская', 'Курская 2-я,', a)
    a = re.sub(r'2-я Посадская', 'Посадская 2-я,', a)
    a = re.sub(r'2-я Пушкарная', 'Пушкарная 2-я,', a)
    a = re.sub(r'3-я Курская', 'Курская 3-я,', a)
    a = re.sub(r'4-я Курская', 'Курская 4-я,', a)
    a = re.sub(r'5 Августа', '5 Августа,', a)
    a = re.sub(r'5-й Орловской стрелковой дивиз', '5-й Орловской стрелковой дивизии, ул, ', a)
    a = re.sub(r'6-й Орловской дивизии', '6-й Орловской дивизии,', a)
    a = re.sub(r'60-летия Октября', '60-летия Октября,', a)
    a = re.sub(r'7 Ноября', '7 Ноября,', a)
    a = re.sub(r'8 Марта', '8 Марта,', a)
    a = re.sub(r'Абрамова и Соколова', 'Абрамова и Соколова,', a)
    a = re.sub(r'Авиационная', 'Авиационная,', a)
    a = re.sub(r'Автовокзальная', 'Автовокзальная,', a)
    a = re.sub(r'Автогрейдерная', 'Автогрейдерная,', a)
    a = re.sub(r'Алроса', 'Алроса,', a)
    a = re.sub(r'Андреева', 'Андреева,', a)
    a = re.sub(r'Андриабужная', 'Андриабужная,', a)
    a = re.sub(r'Андрианова', 'Андрианова,', a)
    a = re.sub(r'Антонова', 'Антонова,', a)
    a = re.sub(r'Аптечный', 'Аптечный,', a)
    a = re.sub(r'Артельный', 'Артельный,', a)
    a = re.sub(r'Багажный', 'Багажный,', a)
    a = re.sub(r'Балтийский', 'Балтийский,', a)
    a = re.sub(r'Березовая', 'Березовая,', a)
    a = re.sub(r'Бетонный', 'Бетонный,', a)
    a = re.sub(r'Ботанический', 'Ботанический,', a)
    a = re.sub(r'Брянская', 'Брянская,', a)
    a = re.sub(r'Бунина', 'Бунина,', a)
    a = re.sub(r'Бурова', 'Бурова,', a)
    a = re.sub(r'Васильевская', 'Васильевская,', a)
    a = re.sub(r'Ватная', 'Ватная,', a)
    a = re.sub(r'Веселая', 'Веселая,', a)
    a = re.sub(r'Воскресенский', 'Воскресенский,', a)
    a = re.sub(r'Высоковольтная', 'Высоковольтная,', a)
    a = re.sub(r'Выставочная', 'Выставочная,', a)
    a = re.sub(r'Гагарина', 'Гагарина,', a)
    a = re.sub(r'Гайдара', 'Гайдара,', a)
    a = re.sub(r'Гвардейская', 'Гвардейская,', a)
    a = re.sub(r'Генерала Жадова', 'Генерала Жадова,', a)
    a = re.sub(r'Генерала Родина', 'Генерала Родина,', a)
    a = re.sub(r'Герцена', 'Герцена,', a)
    a = re.sub(r'Гористый', 'Гористый,', a)
    a = re.sub(r'Городская', 'Городская,', a)
    a = re.sub(r'Гостиная', 'Гостиная,', a)
    a = re.sub(r'Грановского', 'Грановского,', a)
    a = re.sub(r'Грузовая', 'Грузовая,', a)
    a = re.sub(r'Гуртьева', 'Гуртьева,', a)
    a = re.sub(r'Дарвина', 'Дарвина,', a)
    a = re.sub(r'Деповская', 'Деповская,', a)
    a = re.sub(r'Детский', 'Детский,', a)
    a = re.sub(r'Дмитрия Блынского', 'Дмитрия Блынского,', a)
    a = re.sub(r'Достоевского', 'Достоевского,', a)
    a = re.sub(r'Дубровинского набережная ул', 'Дубровинского, наб', a)
    a = re.sub(r'Елецкая', 'Елецкая,', a)
    a = re.sub(r'Емельяна Пугачева', 'Емельяна Пугачева,', a)
    a = re.sub(r'Емлютина', 'Емлютина,', a)
    a = re.sub(r'Жилинская', 'Жилинская,', a)
    a = re.sub(r'Житный', 'Житный,', a)
    a = re.sub(r'Загородный', 'Загородный,', a)
    a = re.sub(r'Запрудная', 'Запрудная,', a)
    a = re.sub(r'Зареченская', 'Зареченская,', a)
    a = re.sub(r'Зеленина', 'Зеленина,', a)
    a = re.sub(r'Игнатова', 'Игнатова,', a)
    a = re.sub(r'Калинина', 'Калинина,', a)
    a = re.sub(r'Каменского', 'Каменского,', a)
    a = re.sub(r'Карачевская', 'Карачевская,', a)
    a = re.sub(r'Карачевский', 'Карачевский,', a)
    a = re.sub(r'Карачевское', 'Карачевское,', a)
    a = re.sub(r'Карла Маркса', 'Карла Маркса,', a)
    a = re.sub(r'Картукова', 'Картукова,', a)
    a = re.sub(r'Каштановый', 'Каштановый,', a)
    a = re.sub(r'Кировский', 'Кировский,', a)
    a = re.sub(r'Кирпичная', 'Кирпичная,', a)
    a = re.sub(r'Кирпичного завода', 'Кирпичного завода,', a)
    a = re.sub(r'Кирпичный', 'Кирпичный,', a)
    a = re.sub(r'Коллективная', 'Коллективная,', a)
    a = re.sub(r'Колпакчи', 'Колпакчи,', a)
    a = re.sub(r'Коммуны', 'Коммуны,', a)
    a = re.sub(r'Комсомольская', 'Комсомольская,', a)
    a = re.sub(r'Комсомольский', 'Комсомольский,', a)
    a = re.sub(r'Контактная', 'Контактная,', a)
    a = re.sub(r'Корчагина', 'Корчагина,', a)
    a = re.sub(r'Космонавтов', 'Космонавтов,', a)
    a = re.sub(r'Костомаровская', 'Костомаровская,', a)
    a = re.sub(r'Красина', 'Красина,', a)
    a = re.sub(r'Красноармейская', 'Красноармейская,', a)
    a = re.sub(r'Кромская', 'Кромская,', a)
    a = re.sub(r'Кромское', 'Кромское,', a)
    a = re.sub(r'Кузнецова', 'Кузнецова,', a)
    a = re.sub(r'Куйбышева', 'Куйбышева,', a)
    a = re.sub(r'Кукушкина', 'Кукушкина,', a)
    a = re.sub(r'Лазо', 'Лазо,', a)
    a = re.sub(r'Латышских Стрелков', 'Латышских Стрелков,', a)
    a = re.sub(r'Левоовражный', 'Левоовражный,', a)
    a = re.sub(r'Левый Берег реки Оки', 'Левый Берег реки Оки,', a)
    a = re.sub(r'Левый берег реки Оки', 'Левый Берег реки Оки,', a)
    a = re.sub(r'Ленина', 'Ленина,', a)
    a = re.sub(r'Лермонтова', 'Лермонтова,', a)
    a = re.sub(r'Лескова', 'Лескова,', a)
    a = re.sub(r'Лесная', 'Лесная,', a)
    a = re.sub(r'Летний', 'Летний,', a)
    a = re.sub(r'Ливенская', 'Ливенская,', a)
    a = re.sub(r'Линейная', 'Линейная,', a)
    a = re.sub(r'Ломоносова', 'Ломоносова,', a)
    a = re.sub(r'Льва Толстого', 'Льва Толстого,', a)
    a = re.sub(r'Ляшко', 'Ляшко,', a)
    a = re.sub(r'Магазинная', 'Магазинная,', a)
    a = re.sub(r'Максима Горького', 'Максима Горького,', a)
    a = re.sub(r'Маринченко', 'Маринченко,', a)
    a = re.sub(r'Маслозаводской', 'Маслозаводской,', a)
    a = re.sub(r'Матвеева', 'Матвеева,', a)
    a = re.sub(r'Матроса Силякова', 'Матроса Силякова,', a)
    a = re.sub(r'Матросова', 'Матросова,', a)
    a = re.sub(r'Машиностроительная', 'Машиностроительная,', a)
    a = re.sub(r'Машкарина', 'Машкарина,', a)
    a = re.sub(r'Маяковского', 'Маяковского,', a)
    a = re.sub(r'Мебельная', 'Мебельная,', a)
    a = re.sub(r'Медведева', 'Медведева,', a)
    a = re.sub(r'Межевой', 'Межевой,', a)
    a = re.sub(r'Металлургов', 'Металлургов,', a)
    a = re.sub(r'Мира', 'Мира,', a)
    a = re.sub(r'Михалицына', 'Михалицына,', a)
    a = re.sub(r'Мичурина', 'Мичурина,', a)
    a = re.sub(r'Молодежи', 'Молодежи,', a)
    a = re.sub(r'Молодогвардейский', 'Молодогвардейский,', a)
    a = re.sub(r'Мопра', 'МОПРа', a)
    a = re.sub(r'МОПРа', 'МОПРа,', a)
    a = re.sub(r'Московская', 'Московская,', a)
    a = re.sub(r'Московское', 'Московское,', a)
    a = re.sub(r'Моховская', 'Моховская,', a)
    a = re.sub(r'Наугорское', 'Наугорское,', a)
    a = re.sub(r'Новая', 'Новая,', a)
    a = re.sub(r'Новикова', 'Новикова,', a)
    a = re.sub(r'Новосильская', 'Новосильская,', a)
    a = re.sub(r'Новосильский', 'Новосильский,', a)
    a = re.sub(r'Новосильское', 'Новосильское,', a)
    a = re.sub(r'Нормандия-Неман', 'Нормандия-Неман,', a)
    a = re.sub(r'Огородный', 'Огородный,', a)
    a = re.sub(r'Октябрьская', 'Октябрьская,', a)
    a = re.sub(r'Орловская область, д.Жилина, Генерала Лаврова', 'Генерала Лаврова,', a)
    a = re.sub(r'Орелстроевская', 'Орелстроевская,', a)
    a = re.sub(r'Орлицкий', 'Орлицкий,', a)
    a = re.sub(r'Орловских Партизан', 'Орловских Партизан,', a)
    a = re.sub(r'Осипенко', 'Осипенко,', a)
    a = re.sub(r'Панчука', 'Панчука,', a)
    a = re.sub(r'Парижской Коммуны', 'Парижской Коммуны', a)
    a = re.sub(r'Паровозная', 'Паровозная,', a)
    a = re.sub(r'Песковская', 'Песковская,', a)
    a = re.sub(r'Пионерская', 'Пионерская,', a)
    a = re.sub(r'Пищевой', 'Пищевой,', a)
    a = re.sub(r'Планерная', 'Планерная,', a)
    a = re.sub(r'Плещеевская', 'Плещеевская,', a)
    a = re.sub(r'Победы', 'Победы,', a)
    a = re.sub(r'Пожарная', 'Пожарная,', a)
    a = re.sub(r'Покровская', 'Покровская,', a)
    a = re.sub(r'Полевая', 'Полевая,', a)
    a = re.sub(r'Полесская', 'Полесская,', a)
    a = re.sub(r'Поселковая', 'Поселковая,', a)
    a = re.sub(r'Поликарпова', 'Поликарпова,', a)
    a = re.sub(r'Полковника Старинова', 'Полковника Старинова,', a)
    a = re.sub(r'Полярный', 'Полярный,', a)
    a = re.sub(r'Почтовый', 'Почтовый,', a)
    a = re.sub(r'Приборостроительная', 'Приборостроительная,', a)
    a = re.sub(r'Привокзальная', 'Привокзальная,', a)
    a = re.sub(r'Привокзальный', 'Привокзальный,', a)
    a = re.sub(r'Придорожная', 'Придорожная,', a)
    a = re.sub(r'Пролетарская Гора', 'Пролетарская Гора, ул,', a)
    a = re.sub(r'Прядильная', 'Прядильная,', a)
    a = re.sub(r'Пушкина', 'Пушкина,', a)
    a = re.sub(r'Пятницкий', 'Пятницкий,', a)
    a = re.sub(r'Рабочий', 'Рабочий,', a)
    a = re.sub(r'Раздольная', 'Раздольная,', a)
    a = re.sub(r'Революции', 'Революции,', a)
    a = re.sub(r'Рельсовая', 'Рельсовая,', a)
    a = re.sub(r'Речной', 'Речной,', a)
    a = re.sub(r'Рижский', 'Рижский,', a)
    a = re.sub(r'Родзевича-Белевича', 'Родзевича-Белевича,', a)
    a = re.sub(r'Розы Люксембург', 'Розы Люксембург,', a)
    a = re.sub(r'Рощинская', 'Рощинская,', a)
    a = re.sub(r'Русанова', 'Русанова,', a)
    a = re.sub(r'Ручейный', 'Ручейный,', a)
    a = re.sub(r'Рябиновая', 'Рябиновая,', a)
    a = re.sub(r'Садовского', 'Садовского,', a)
    a = re.sub(r'Садовый', 'Садовый,', a)
    a = re.sub(r'Салтыкова-Щедрина', 'Салтыкова-Щедрина,', a)
    a = re.sub(r'Саханская', 'Саханская,', a)
    a = re.sub(r'Светофорный', 'Светофорный,', a)
    a = re.sub(r'Связистов', 'Связистов,', a)
    a = re.sub(r'Северная', 'Северная,', a)
    a = re.sub(r'Семинарская', 'Семинарская,', a)
    a = re.sub(r'Серпуховская', 'Серпуховская,', a)
    a = re.sub(r'Силикатная', 'Силикатная,', a)
    a = re.sub(r'Скульптурная', 'Скульптурная,', a)
    a = re.sub(r'Смоленская', 'Смоленская,', a)
    a = re.sub(r'Советская', 'Советская,', a)
    a = re.sub(r'Соляной', 'Соляной,', a)
    a = re.sub(r'Спивака', 'Спивака,', a)
    a = re.sub(r'Степана Разина', 'Степана Разина,', a)
    a = re.sub(r'Студенческая', 'Студенческая,', a)
    a = re.sub(r'Сурена Шаумяна', 'Сурена Шаумяна,', a)
    a = re.sub(r'Тамбовская', 'Тамбовская,', a)
    a = re.sub(r'Товарный', 'Товарный,', a)
    a = re.sub(r'Трамвайный', 'Трамвайный,', a)
    a = re.sub(r'Транспортный', 'Транспортный,', a)
    a = re.sub(r'Тульская', 'Тульская,', a)
    a = re.sub(r'Тургенева', 'Тургенева,', a)
    a = re.sub(r'Узловая', 'Узловая,', a)
    a = re.sub(r'Федотовой', 'Федотовой,', a)
    a = re.sub(r'Фомина', 'Фомина,', a)
    a = re.sub(r'Фурманова', 'Фурманова,', a)
    a = re.sub(r'Хвойный', 'Хвойный,', a)
    a = re.sub(r'Хлебозаводской', 'Хлебозаводской,', a)
    a = re.sub(r'Хлебный', 'Хлебный,', a)
    a = re.sub(r'Холодная', 'Холодная,', a)
    a = re.sub(r'Цветаева', 'Цветаева,', a)
    a = re.sub(r'Циолковского', 'Циолковского,', a)
    a = re.sub(r'Чапаева', 'Чапаева,', a)
    a = re.sub(r'Часовая', 'Часовая,', a)
    a = re.sub(r'Черепичная', 'Черепичная,', a)
    a = re.sub(r'Черкасская', 'Черкасская,', a)
    a = re.sub(r'Чечневой', 'Чечневой,', a)
    a = re.sub(r'Чкалова', 'Чкалова,', a)
    a = re.sub(r'Шахматный', 'Шахматный,', a)
    a = re.sub(r'Шпагатный', 'Шпагатный,', a)
    a = re.sub(r'Шульгина', 'Шульгина,', a)
    a = re.sub(r'Щепная', 'Щепная,', a)
    a = re.sub(r'Щорса', 'Щорса,', a)
    a = re.sub(r'Элеваторный', 'Элеваторный,', a)
    a = re.sub(r'Энгельса', 'Энгельса,', a)
    a = re.sub(r'Южный', 'Южный,', a)
    a = re.sub(r'Яблочная', 'Яблочная,', a)
    a = re.sub(r'Ягодный', 'Ягодный,', a)
    a = re.sub(r'Ермолова', 'Ермолова,', a)
    a = re.sub(r'Смоленский', 'Смоленский,', a)
    print(a)
    row.value = a
    # print(f'в строке {n} значение {a}')
# for row in sheet['B']:
#     a = re.split(' пер', row.value)
#     print(a)
# print(ulica, dom)
wb.save('1.xlsx')

