import abc
from typing import Any, Dict, List, Optinoal, Set, Tuple, Union


class ABC_Excel(abc):

    def get_table(self):
        pass
    
    def extract_table_to_csv(self):
        pass
    
    def save(self):
        pass
