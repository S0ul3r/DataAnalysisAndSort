import pandas as pd
from pandas.core.frame import DataFrame

# Tworzymy dictionary z excela
orders = pd.read_excel('analizaDanych.xlsx', sheet_name='Sheet1')
# Tworzymy listę wszystkich przedstawicieli z PRZEDSTAWICIEL_HANDL
names = list(dict.fromkeys(orders['PRZEDSTAWICIEL_HANDL'].tolist()))
# Wartość ilu mamy przedstawicieli (na później)
number_names = len(names)


# Tworzymy dictionary dla każdego pracownika z jego imieniem
przedstawiciel = {}
for name in names:
    przedstawiciel[str(name)] = []

# loopujemy po każdym z pracowników
for i in range(number_names):
    #loopujemy po każdym wierszu danych z excela
    for index, row in orders.iterrows():
        # Sprawdzamy do którego przedstawiciela zapisać dany wiersz z excela
        if names[i] == orders._get_value(index, "PRZEDSTAWICIEL_HANDL"):
            # Przypisujemy dane z excela do danego przedstawiciela
            przedstawiciel[names[i]].append(row)

# Przypisujemy ścieżkę zapisu i writer używany do zapisu .xslx używając biblioteki pandas
path = "Dane.xlsx"
writer = pd.ExcelWriter(path, engine='xlsxwriter')
# Loopujemy po kazdym z pracowników
for i in range(number_names):
    # Zameiniamy listę zleceń dla danego pracownika na dataframe z biblioteki pandas
    df = pd.DataFrame(przedstawiciel[names[i]])
    # Zapis do excela
    df.to_excel(writer, sheet_name = str(names[i]), index=False)
writer.save()


###### Jeśli chcielibyśmy by każdy pracownik został zapisany do osobnego pliku, zakładka nazwana odpowiednio jego imieniem

#writer = pd.ExcelWriter(path, engine='xlsxwriter')
# for i in range(number_names):
#     df = pd.DataFrame(przedstawiciel[names[i]])
#     path = str(names[i]) + ".xlsx"
#     df.to_excel(writer, sheet_name = str(names[i]), index=False)
#     writer.save()