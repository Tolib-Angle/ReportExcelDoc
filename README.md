# ReportExcelDoc
### Author: [Tolib](https://github.com/Tolib-Angle) Date of loading: 08.06.2023
Console application designed to compile a hierarchical report in C # and output them to an Excel document

Necessary packages to run the program: EPPlus 6.2.4 + , RecyclableMemoryStream.1.4.1

The starting point of the programm: [Program.cs](https://github.com/Tolib-Angle/ReportExcelDoc/blob/main/Program.cs)

--------------------------------------

# Task

Для реализации необходимо использовать C# (любой удобный фрэйм) и SQL (MS SQL, SQL Lite). БД должна быть с набитыми данными для полной проверки.

Есть 2 таблицы:

Izdel

| Столбец | Тип данных     | Описание                     |
|---------|----------------|------------------------------|
| Id      | bigint         | Уникальный идентификатор (Primary Key) |
| Name    | varchar(100)   | Наименование изделия          |
| Price   | Decimal(20,2)  | Цена покупки/сборки изделия   |

Links

| Столбец | Тип данных     | Описание                      |
|---------|----------------|-------------------------------|
| IzdelUp | bigint         | Ссылка на вышестоящее изделие |
| Izdel   | bigint         | Ссылка на текущее изделие     |
| kol     | int            | Количество текущих изделий, входящих в вышестоящее |


Необходимо сформировать иерархический отчет в MS Excel, в котором отразить количество и стоимость изделий, находящихся на первых трех уровнях.
Количество изделия на первом уровне иерархии принимаем равным 1. Стоимость изделия - количество по изделию, умноженное на цену по изделию, плюс стоимость всех изделий нижестоящих уровней.

Пример отчета:

| Изделие   | Кол-во | Стоимость | Цена  |
|-----------|--------|-----------|-------|
| Изделие 1 | 1      | 3000      | 800   |
| Изделие 2 | 10     | 1000      | 1000  |
| Изделие 3 | 2      | 1000      | 400   |
| Изделие 5 | 2      | 600       | 300   |
| Изделие 4 | 1      | 600       | 400   |
| Изделие 2 | 1      | 100       | 100   |
| Изделие 6 | 5      | 100       | 20    |
| Изделие 7 | 1      | 7000      | 1000  |
| Изделие 8 | 20     | 2000      | 100   |
| Изделие 3 | 10     | 4000      | 400   |


Example of a generated report using the program: [Отчет.xlsx](https://github.com/Tolib-Angle/ReportExcelDoc/blob/main/Отчет.xlsx)

### Attention!
The contents of the [Отчет.xlsx](https://github.com/Tolib-Angle/ReportExcelDoc/blob/main/Отчет.xlsx) have been generated and are random. Any coincidence is just a coincidence, also the report could be generated incorrectly due to the randomness of the data (the links could be repeated or contradict each other).

The report was compiled on a database containing 500 records in each table

If you have any additional questions, write to the author: [Tolib](https://github.com/Tolib-Angle)