import pandas as pd
from time import time
import dateutil.relativedelta as difftime
from datetime import datetime


def save_to_exc(df: pd.DataFrame, filename: str):
    """
    save_to_exc(df: pd.DataFrame, filename: str)
    Данная функция принемает на вход объект класса pd.DataFrame и строку,
    содержвщую путь к файлу.
    Сохраняет pd.DataFrame как Excel файл
    """
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, startrow=0, header=True)
    writer.save()


def merge_from_exc(input_file1: str, input_file2: str, filtered=False):
    """
    merge_from_exc(input_file1: str, input_file2: str)
    Данная функция приниает на вход пути к ексель файлам, возвращает объект
    класса pd.DataFrame, содержащий таблицу результат
    слияния двух объектов pd.DataFrame
    """
    df_o = pd.read_excel(input_file1, parse_dates=True)
    if filtered:
        df_o = get_last_month(df_o)
    df_ol = pd.read_excel(input_file2)
    return pd.merge(df_o, df_ol)


def get_last_month(df):
    """
    Ввиду того, что DateTime не является индексом
    (поле потенциально не уникальное)
    пришлось отсекать "вручную", метод Last не работает,
    если поле с датой не индекс.
    """
    last_month = datetime.today() - difftime.relativedelta(months=1)
    return df[df.DateTime > "{:%Y-%m-%d}".format(last_month)]


def combine_columns(inputdb):
    """
    combine_columns(inputdb)
    Данная функция принимает на вход объект класса pd.DataFrame,
    В результате выполнения, возвращает объект, содержащий информацию
    согласно приложенному заданию:

    самые популярные за последний месяц продукты

    суммарная выручка по каждому такому продукту

    средний чек заказов, в которых есть такие продукты
    """
    grouped = inputdb.groupby(["ProductId"])

    # Считаем количество сгруппированных элементов и сортируем
    # в контексте задачи вместо "OrderId" можно взять любое поле.
    results = grouped["OrderId"].count()

    # переименовываем поле для дальшнейшего удобства
    results.name = "Count"

    results = results.sort_values(ascending=False)
    results = pd.merge(
                          results,
                          grouped["Price"].sum(),
                          left_index=True,
                          right_index=True
                        )

    # Добавляем поле
    results["avg_price"] = results["Price"]/results["Count"]
    return results


def main():
    merged_df = merge_from_exc("orders.xlsx", "order_lines.xlsx", True)
    print(merged_df)
    results_df = combine_columns(merged_df)
    print(results_df)
    save_to_exc(results_df, "results.xlsx")


if __name__ == "__main__":
    start_time = time()
    main()
    print(time() - start_time)
