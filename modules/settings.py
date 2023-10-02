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
    consumables: dict[str, Consumable]
    hourly_rates: dict[str, HourlyRate]
    mark_ups: dict[str, MarkUp]
    flags: dict
    
if __name__ == "__main__":
    pass
