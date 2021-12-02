#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
Settings object to hold various dataclasses

"""
from dataclasses import dataclass
from .consumables import Consumable
from .hourlyrates import HourlyRate
from .markups import MarkUp

@dataclass
class Settings():
    consumables: Consumable
    hourly_rates: HourlyRate
    mark_ups: MarkUp
    
if __name__ == "__main__":
    pass
