from pathlib import Path

import xlwings as xw

from exceptions import ProductAlreadyDeletedError, ProductDeletionInputError, ProductIdNotFoundError, ProductDeletionSystemError
from database import connect_db
from excel_reader import get_worksheet, price_sheet
from logging_config import setup_logging
from sync_utils import create_backup, restore_excel_file

logger = setup_logging()




def validate_product_ids(conn, raw_ids):
    if raw_ids is None:
        return None

    raw_ids = raw_ids.strip()

    if not raw_ids:
        raise ProductDeletionInputError("Не указано ни одного ID товара")

    parts = [part.strip() for part in raw_ids.split(",")]

    if any(not part for part in parts):
        raise ProductDeletionInputError("ID товаров должны быть указаны через запятую")

    if any(not part.isdigit() for part in parts):
        raise ProductDeletionInputError("ID товара должен быть целым положительным числом")

    product_ids = [int(part) for part in parts]

    if any(product_id <= 0 for product_id in product_ids):
        raise ProductDeletionInputError("ID товара должен быть целым положительным числом")

    if len(product_ids) != len(set(product_ids)):
        raise ProductDeletionInputError("Один и тот же ID указан несколько раз")

    placeholders = ",".join("?" for _ in product_ids)

    rows = conn.execute(
        f"""
        SELECT id, active
        FROM products
        WHERE id IN ({placeholders});
        """,
        product_ids
    ).fetchall()

    db_products = {row["id"]: row["active"] for row in rows}

    missing_ids = [product_id for product_id in product_ids if product_id not in db_products]

    if missing_ids:
        raise ProductIdNotFoundError(f"Товары с ID {', '.join(map(str, missing_ids))} отсутствуют в базе данных")

    deleted_ids = [product_id for product_id in product_ids if db_products[product_id] == 0]

    if deleted_ids:
        raise ProductAlreadyDeletedError(f"Товары с ID {', '.join(map(str, deleted_ids))} уже удалены")

    return product_ids

# def get_product_ids_for_deletion(conn):
#     while True:
#         raw_ids = ask_product_ids()
#
#         if raw_ids is None:
#             return None
#
#         try:
#             return validate_product_ids(conn, raw_ids)
#
#         except ProductDeletionInputError as exc:
#             messagebox.showerror("Ошибка удаления товара", str(exc))


def deactivate_products(conn, product_ids):
    placeholders = ",".join("?" for _ in product_ids)

    cursor = conn.execute(
        f"""
        UPDATE products
        SET active = 0
        WHERE id IN ({placeholders})
          AND active = 1;
        """,
        product_ids
    )

    if cursor.rowcount != len(product_ids):
        raise ProductDeletionSystemError("Не удалось деактивировать все выбранные товары")

def find_product_rows_in_excel(worksheet, product_ids):
    rows_by_id = {}

    last_row = worksheet.range(f"A{worksheet.cells.last_cell.row}").end("up").row

    for row_num in range(2, last_row + 1):
        raw_product_id = worksheet.range(f"A{row_num}").value

        if raw_product_id is None:
            continue

        try:
            product_id = int(raw_product_id)
        except (TypeError, ValueError):
            continue

        if product_id in product_ids:
            rows_by_id[product_id] = row_num

    missing_ids = [product_id for product_id in product_ids if product_id not in rows_by_id]

    if missing_ids:
        raise ProductDeletionSystemError(f"В Excel не найдены товары с ID: {', '.join(map(str, missing_ids))}")

    return rows_by_id

def delete_product_rows_from_excel(worksheet, product_ids):
    rows_by_id = find_product_rows_in_excel(worksheet, product_ids)
    row_numbers = sorted(rows_by_id.values(), reverse=True)

    for row_num in row_numbers:
        worksheet.range(f"{row_num}:{row_num}").delete()

def show_deletion_success(app, product_ids):
    ids_text = ", ".join(map(str, product_ids))
    app.alert(f"Товары успешно удалены.\n\nУдаленные ID: {ids_text}", title="Удаление товаров", buttons="ok", mode="info")


def show_deletion_error(app, error, excel_restored):
    if excel_restored:
        message = f"Удаление товаров не выполнено.\n\nИзменения в базе данных отменены, Excel-файл восстановлен.\n\nОшибка: {error}"
    else:
        message = f"Удаление товаров не выполнено.\n\nИзменения в базе данных отменены, но автоматически восстановить Excel-файл не удалось.\n\nОшибка: {error}"

    app.alert(message, title="Ошибка удаления товаров", buttons="ok", mode="critical")


def delete_products(workbook, raw_ids):
    excel_path = Path(workbook.fullname)
    backup_path = excel_path.with_name(f"{excel_path.stem}_backup{excel_path.suffix}")
    app = workbook.app
    conn = connect_db()
    backup_created = False

    logger.info("Запуск удаления товаров: excel_file=%s", excel_path)

    try:
        product_ids = validate_product_ids(conn, raw_ids)
        logger.info("Получены ID товаров для удаления: %s", product_ids)

        workbook.save()

        create_backup(excel_path, backup_path)
        backup_created = True

        worksheet = get_worksheet(workbook, price_sheet)

        deactivate_products(conn, product_ids)
        delete_product_rows_from_excel(worksheet, product_ids)

        workbook.save()
        conn.commit()

    except Exception as error:
        conn.rollback()

        excel_restored = False

        if backup_created and backup_path.exists():
            try:
                restore_excel_file(workbook, app, excel_path, backup_path)
                excel_restored = True
            except Exception:
                logger.exception("Не удалось автоматически восстановить Excel-файл")

        show_deletion_error(app, error, excel_restored)
        raise

    else:
        if backup_path.exists():
            backup_path.unlink()

        show_deletion_success(app, product_ids)

    finally:
        conn.close()

def main():
    workbook = xw.Book.caller()
    delete_products(workbook)

