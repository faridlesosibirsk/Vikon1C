# -*- coding: utf-8 -*-

class CompetitionGroup:
    __name: str = ""
    def __init__(self) -> None:
        pass
    def setName(self, line) -> None:
        if line != None and isinstance(line, str):
            if "Конкурсная" in line:
                self.__name = line.split('Конкурсная группа - ')[1]
                print(self.__name)
    def getName(self) -> str:
        return self.__name


