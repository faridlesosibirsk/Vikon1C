# -*- coding: utf-8 -*-

class Abit:
    __abiturient = []
    def append(self, x, line, employees_sheet) -> None:
        __abiturient = []
        if line != None and isinstance(line, int): 
            abit = []
            for y in range(1, employees_sheet.max_column+1):
                field = employees_sheet.cell(row=x, column=y).value
                abit.append(field)
            self.__abiturient.append(abit)
    def abiturient(self) -> []:
        return self.__abiturient
