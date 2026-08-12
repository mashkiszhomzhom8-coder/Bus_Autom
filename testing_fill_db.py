import sqlite3
import sys

from datetime import date
from pathlib import Path

import xlwings as xw

from exceptions import DatabaseSystemError
from database import connect_db
from excel_reader import products_reader


def clean_db(conn):
    conn.executescript(
        """
        DELETE FROM order_items;
        DELETE FROM prod_movement;
        DELETE FROM supply_items;
        DELETE FROM inventory_items;
        DELETE FROM change_log;

        DELETE FROM orders;
        DELETE FROM supply_batch;
        DELETE FROM inventory_check;
        DELETE FROM season_opening_stock;

        DELETE FROM seasons;
        DELETE FROM products;
        """
    )


def append_fst_season(conn, workbook):
    excel_file = Path(workbook.fullname).name

    cursor = conn.execute(
        """
        INSERT INTO seasons (
            start_date,
            excel_file,
            status
        )
        VALUES (?, ?, ?);
        """,
        (
            "22.07.2026",
            excel_file,
            "active"
        )
    )

    return cursor.lastrowid


def validate_products(conn, products):
    if not products:
        raise DatabaseSystemError("В Excel не найдено товаров для первичного заполнения БД")

    seen_products = {}

    for product in products:
        key = (product["name"], product["size"])

        if key in seen_products:
            first_product_id = seen_products[key]
            raise DatabaseSystemError(f"В Excel повторяется товар: название={product['name']!r}, размер={product['size']!r}, ID={first_product_id} и ID={product['id']}")

        seen_products[key] = product["id"]

    products_count = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM products;
        """
    ).fetchone()["count"]

    if products_count != 0:
        raise DatabaseSystemError("Первичное заполнение невозможно: таблица products уже содержит товары")


def fill_products_from_excel(conn, products):
    validate_products(conn, products)

    try:
        for product in products:
            conn.execute(
                """
                INSERT INTO products (
                    id,
                    name,
                    size,
                    price,
                    current_stock,
                    active
                )
                VALUES (?, ?, ?, ?, ?, ?);
                """,
                (
                    product["id"],
                    product["name"],
                    product["size"],
                    product["price"],
                    product["current_stock"],
                    1
                )
            )

    except sqlite3.IntegrityError as exc:
        raise DatabaseSystemError("Не удалось заполнить products: обнаружен конфликт ID или связки name + size") from exc

    return len(products)


def fill_season_opening_stock(conn, products, season_id):
    for product in products:
        conn.execute(
            """
            INSERT INTO season_opening_stock (
                season_id,
                product_id,
                opening_stock
            )
            VALUES (?, ?, ?);
            """,
            (
                season_id,
                product["id"],
                product["remain"]
            )
        )


def create_initial_inventory(conn, products, season_id):
    cursor = conn.execute(
        """
        INSERT INTO inventory_check (
            season_id,
            date
        )
        VALUES (?, ?);
        """,
        (
            season_id,
            date.today().isoformat()
        )
    )

    inventory_check_id = cursor.lastrowid

    for product in products:
        conn.execute(
            """
            INSERT INTO inventory_items (
                inventory_check_id,
                product_id,
                expected_remain,
                actual_remain,
                difference
            )
            VALUES (?, ?, ?, ?, ?);
            """,
            (
                inventory_check_id,
                product["id"],
                product["remain"],
                product["remain"],
                0
            )
        )

    return inventory_check_id


def initialize_products(excel_path):
    workbook = xw.Book(str(excel_path))
    conn = connect_db()

    try:
        products = products_reader(workbook)

        season_id = append_fst_season(conn, workbook)
        inserted_count = fill_products_from_excel(conn, products)
        fill_season_opening_stock(conn, products, season_id)
        create_initial_inventory(conn, products, season_id)

        conn.commit()

        return inserted_count

    except Exception:
        conn.rollback()

        try:
            clean_db(conn)
            conn.commit()
        except Exception:
            conn.rollback()

        raise

    finally:
        conn.close()

if __name__ == "__main__":
    initialize_products("Шаблон_прайса.xlsm")
    print('Успешно заполнено')


# def run_from_python(excel_file):
#     excel_path = Path(excel_file).resolve()

#
#     try:
#         inserted_count = initialize_products(workbook)
#         print(f"База данных успешно инициализирована. Добавлено товаров: {inserted_count}")
#
#     finally:
#         workbook.close()
#
#
# if __name__ == "__main__":
#     if len(sys.argv) != 2:
#         raise RuntimeError("Укажите путь к тестовому Excel-файлу .xlsm")
#
#     run_from_python(sys.argv[1])