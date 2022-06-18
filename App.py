# -*- coding: utf-8 -*-
from Excel import Excel
from CheckList import CheckList
from Abit import Abit

class App:
    __e: Excel
    def __init__(self) -> None:
        self.__e = Excel()
    def run(self) -> None:
        #c = CompetitionGroup() 
        ch = CheckList()
        for x in self.__e.rangeMax_row():
            #c.setName(self.__e.getCell(x, 1))
            a = Abit()
            a.append(x, self.__e.getCell(x, 1), self.__e.getEmployeesSheet())
            
            ch.create(self.__e.getCell(x, 1), a, x, self.__e.getEmployeesSheet())

