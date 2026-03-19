import pandas as pd
import warnings
warnings.filterwarnings('ignore')


def analyze(csv_path: str = "renteasy_data/renteasy_dataset.csv"):
    df = pd.read_csv(csv_path, encoding='utf-8-sig')

    print("=== АНАЛИЗ ДАТАСЕТА RentEasy KZ ===\n")
    print(f"Всего записей: {len(df)}")
    print(f"Источники: {df['source'].value_counts().to_dict()}")
    print(f"Города: {df['city'].value_counts().to_dict()}\n")

    print("--- Средняя цена аренды по городам (тенге/мес) ---")
    city_prices = df.groupby('city')['price_tenge'].mean()
    for city, price in city_prices.items():
        print(f"  {city}: {price:,.0f} тг")

    print("\n--- Средняя цена за кв.м по городам (тенге/кв.м) ---")
    city_pm2 = df.groupby('city')['price_per_m2'].mean()
    for city, price in city_pm2.items():
        print(f"  {city}: {price:,.0f} тг/м²")

    print("\n--- Распределение по количеству комнат ---")
    rooms_map = {0: 'Студия', 1: '1-комн', 2: '2-комн', 3: '3-комн', 4: '4-комн'}
    rooms_dist = df['rooms'].value_counts().sort_index()
    for rooms, count in rooms_dist.items():
        label = rooms_map.get(rooms, f'{rooms}-комн')
        pct = count / len(df) * 100
        print(f"  {label}: {count} ({pct:.1f}%)")

    print("\n--- Топ-5 самых дорогих районов Астаны ---")
    astana = df[df['city'] == 'Астана']
    if not astana.empty:
        top_districts = astana.groupby('district')['price_tenge'].mean().nlargest(5)
        for district, price in top_districts.items():
            print(f"  {district}: {price:,.0f} тг/мес")

    print("\n--- Топ-5 самых дорогих районов Алматы ---")
    almaty = df[df['city'] == 'Алматы']
    if not almaty.empty:
        top_districts = almaty.groupby('district')['price_tenge'].mean().nlargest(5)
        for district, price in top_districts.items():
            print(f"  {district}: {price:,.0f} тг/мес")

    print("\n--- Распределение по типу продавца ---")
    seller_dist = df['seller_type'].value_counts()
    for stype, count in seller_dist.items():
        pct = count / len(df) * 100
        print(f"  {stype}: {count} ({pct:.1f}%)")

    print("\n--- Средняя площадь по типу жилья (м²) ---")
    area_by_rooms = df.groupby('rooms')['area_m2'].mean()
    for rooms, area in area_by_rooms.items():
        label = rooms_map.get(rooms, f'{rooms}-комн')
        print(f"  {label}: {area:.1f} м²")

    print("\n--- Как данные помогают принимать решения по RentEasy KZ ---")
    print("1. Разрыв цен по районам → показывает где спрос выше предложения → приоритет для аналитики")
    print("2. Доля агентств vs владельцев → потенциал вытеснения посредников")
    print("3. Средняя площадь vs цена → какой тип жилья в дефиците → фокус платформы")
    print("4. Сезонность дат объявлений → когда запускать маркетинговые кампании")
    print("5. Распределение по городам → где запускать платформу в первую очередь")

    return df


if __name__ == "__main__":
    df = analyze()
