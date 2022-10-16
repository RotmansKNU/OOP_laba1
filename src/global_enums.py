from enum import Enum


class GlobalErrorMessages(Enum):
    AdditionOperationError = 'Expression with plus sign is wrong!'
    SubtractionOperationError = 'Expression with minus sign is wrong!'
    MultiplicationOperationError = 'Expression with multiplying sign is wrong!'
    DivisionOperationError = 'Expression with dividing sign is wrong!'
    MaxOperationError = 'Expression for searching max value is wrong!'
    MinOperationError = 'Expression for searching min value is wrong!'
    ExponentiantOperationError = 'Expression with exponent sign is wrong!'
    ReplacementError = 'Expression for replacing the cell is wrong!'
