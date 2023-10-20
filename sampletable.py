import pandas as pd
import random
import datetime
from faker import Faker


class Utils:
    def form_id(product_list):
        id_base = "ID"
        id_count = {}
        pairs = []

        for product in product_list:
            if product in id_count:
                id = id_count[product]
            else:
                id_numerical = random.randint(1000, 9999)
                id = id_base + str(id_numerical)
                id_count[product] = id

            pairs.append([id, product])

        return pairs

    def form_date(product_list):
        fake = Faker()
        date_pairs = []
        for product in product_list:
            new_date = fake.date_between(start_date="-1y", end_date="today")
            date_pairs.append([new_date, product])

        return date_pairs

    def get_product_data(product_list):
        product_data = []
        unit_price_product = {}

        for product in product_list:
            if product in unit_price_product:
                unit_price = unit_price_product[product]
            else:
                unit_price = round(random.uniform(0.01, 99.99), 2)
                unit_price_product[product] = unit_price

            qty = random.randint(1, 10)

            product_data.append([product, unit_price, qty])

        return product_data


class TableData:
    def __init__(self, product_list):
        self.product_list = product_list
        self.id = Utils.form_id(product_list)
        self.date = Utils.form_date(product_list)
        self.region = random.choice(["region1", "region2", "region3", "region4"])
        self.city = random.choice(["city1", "city2", "city3", "city4"])
        self.qty_details = Utils.get_product_data(product_list)

    def to_table(self):
        base_df = pd.DataFrame(self.id, columns=["ID", "Product"])
        date_df = pd.DataFrame(self.date, columns=["Date", "Product"])
        qty_df = pd.DataFrame(
            self.qty_details, columns=["Product", "Unit Price", "Quantity"]
        )

        int_merge = base_df.merge(date_df, on="Product")
        final_df = int_merge.merge(qty_df, on="Product").drop_duplicates()

        regions = [
            random.choice(["region1", "region2", "region3", "region4"])
            for _ in range(len(final_df))
        ]
        cities = [
            random.choice(["city1", "city2", "city3", "city4"])
            for _ in range(len(final_df))
        ]

        final_df["Total"] = None
        final_df["Region"] = regions
        final_df["City"] = cities

        for idx, col in final_df.iterrows():
            final_df.at[idx, "Total"] = col["Unit Price"] * col["Quantity"]
        return final_df


def main():
    plist = [
        "Apples",
        "Pears",
        "Peanuts",
        "Butter",
        "Milk",
        "Flour",
        "Eggs",
        "Tires",
        "Chips",
        "Peanuts",
        "Flour",
        "Potatoes",
        "Milk",
        "Apples",
        "Pears",
        "Peanuts",
        "Flour",
        "Eggs",
        "Chips",
    ]

    TableData(plist).to_table().to_excel(
        "SearchBox.xlsx", engine="openpyxl", index=False
    )


if __name__ == "__main__":
    main()
